# -*- coding: utf-8 -*-
"""
公共图床上传模块（图床）
============================================================================
本模块的方案参考自用户现有项目：

    C:\\Users\\Administrator\\Desktop\\跨境电商\\跨境电商\\ozon_uploader\\media.py
    C:\\Users\\Administrator\\Desktop\\跨境电商\\跨境电商\\ozon_uploader\\image_tools.py

Ozon / Wildberries 助手上架时，平台服务器必须能自己抓到商品图，所以图片要先
传到「任何人都能访问的公网地址」。原项目沉淀下来的做法是：

  * 一张图先压缩再上传 —— 免费图床普遍限制 1MB，而商品原图常有 2MB；
  * 多家图床依次尝试（FreeImage → BeeIMG → Uguu → 自建图床），
    谁家今天抽风就换下一家，不让一张图失败拖垮整批；
  * 失败按「可重试 / 不可重试」分类：
      - 连接断开、超时、5xx      → 重试（退避 1.5s 起，带随机抖动）
      - 文件被拒、超限、配额用尽 → 立即换下一家，不浪费重试次数
  * 配额用尽的图床本批次直接跳过，不再逐张去打它；
  * 同一个文件只上传一次（带缓存 + 线程安全的「正在上传」表），
    多个线程共用同一张图时不会重复上传。

与原项目的差异：本模块不依赖任何第三方包（只用标准库），
Pillow 存在时压缩图片，不存在时原样上传（大图也比没图强）。

【依赖】
    pip install pillow      # 可选，用于压缩
"""

from __future__ import annotations

import json
import logging
import mimetypes
import random
import threading
import time
import uuid
from pathlib import Path
from urllib.request import Request, urlopen

logger = logging.getLogger("shopify_generator")

#: 失败后重试次数
UPLOAD_ATTEMPTS = 4
#: 重试退避基数（秒），实际等待 = 基数 × 尝试次数 + 随机抖动
UPLOAD_BACKOFF_SECONDS = 1.5
#: 单次上传超时（秒）。免费图床经常慢，能等就多等一会儿
UPLOAD_TIMEOUT_SECONDS = 120.0
#: 压缩后长边像素上限
MAX_SIDE = 1200
#: JPEG 质量
JPEG_QUALITY = 85
#: 小于这个体积的图片原样上传，不必再压
SKIP_BELOW_BYTES = 200 * 1024

#: 图床返回「配额/额度用尽」时的关键字 —— 命中就不再打这家了
QUOTA_MARKERS = (
    "quota",
    "limit reached",
    "too many requests",
    "rate limit",
    "bandwidth exceeded",
    "storage limit",
    "每日上传量已达上限",
    "配额",
)

#: 永久性失败 —— 再试多少次也一样，直接换下一家
PERMANENT_MARKERS = (
    "rejected the file",
    "http error 4",
    "http error 413",
    "no space left",
    "file too large",
    # BeeIMG 的 1MB 限制就是这么报的，再传一次也不会变小
    "file size error",
    "file extension not allowed",
    "file type not allowed",
    "not a valid image",
)

#: 可重试失败 —— 网络抖动，换个时间再试
RETRYABLE_MARKERS = (
    "eof occurred",
    "timed out",
    "timeout",
    "connection reset",
    "connection refused",
    "connection aborted",
    "remote end closed",
    "ssl",
    "badstatusline",
    "incompleteread",
    "temporary failure in name resolution",
    "network is unreachable",
    "urlerror",
    "http error 5",
)


class ImageHostError(RuntimeError):
    """图床上传失败。"""


def pillow_available() -> bool:
    try:
        import PIL  # noqa: F401
    except ImportError:
        return False
    return True


def _mime_for(path: Path) -> str:
    return mimetypes.guess_type(path.name)[0] or "application/octet-stream"


def compress_image(
    path: str | Path,
    max_side: int = MAX_SIDE,
    quality: int = JPEG_QUALITY,
) -> tuple[bytes, str]:
    """把图片压成 JPEG 字节，返回 (内容, mime类型)。

    已经是小图、或者没装 Pillow 时，原样返回。
    """
    source = Path(path)
    content = source.read_bytes()
    if len(content) <= SKIP_BELOW_BYTES:
        return content, _mime_for(source)

    try:
        from PIL import Image
    except ImportError:
        logger.info("未安装 Pillow，%s 原样上传", source.name)
        return content, _mime_for(source)

    try:
        with Image.open(source) as image:
            if image.mode not in ("RGB", "L"):
                image = image.convert("RGB")
            width, height = image.size
            if max(width, height) > max_side:
                scale = max_side / max(width, height)
                image = image.resize(
                    (max(1, int(width * scale)), max(1, int(height * scale))),
                    Image.LANCZOS,
                )
            import io

            buffer = io.BytesIO()
            image.save(buffer, "JPEG", quality=quality, optimize=True)
    except Exception as exc:
        logger.warning("无法压缩 %s (%s)，原样上传", source.name, exc)
        return content, _mime_for(source)

    compressed = buffer.getvalue()
    if len(compressed) >= len(content):
        return content, _mime_for(source)
    logger.info(
        "已压缩 %s: %d KB -> %d KB",
        source.name, len(content) // 1024, len(compressed) // 1024,
    )
    return compressed, "image/jpeg"


def _multipart(
    field: str,
    filename: str,
    content: bytes,
    mime_type: str,
    extra: dict[str, str] | None = None,
) -> tuple[bytes, str]:
    """拼 multipart/form-data 请求体。

    普通字段（比如 API key）必须放在文件前面 —— 有些图床一旦读到文件部分
    就停止解析表单，key 放在后面就永远收不到。
    """
    boundary = uuid.uuid4().hex
    extra_part = b""
    for name, value in (extra or {}).items():
        extra_part += (
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="{name}"\r\n\r\n'
            f"{value}\r\n"
        ).encode()
    head = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="{field}"; filename="{filename}"\r\n'
        f"Content-Type: {mime_type}\r\n\r\n"
    ).encode()
    body = extra_part + head + content + f"\r\n--{boundary}--\r\n".encode()
    return body, f"multipart/form-data; boundary={boundary}"


def _find_url(payload, path: str = "url") -> str:
    """从图床返回里找出图片地址，兼容 url / data.url / files.url 等常见位置。"""
    if isinstance(payload, str):
        text = payload.strip()
        return text if text.startswith("http") else ""
    if isinstance(payload, dict):
        if path in payload:
            value = payload[path]
            return str(value) if value else ""
        for key in ("url", "link", "src", "image", "image_url"):
            value = payload.get(key)
            if isinstance(value, str) and value.startswith("http"):
                return value
        for key in ("data", "files", "result", "file"):
            nested = payload.get(key)
            if nested is not None:
                found = _find_url(nested, path)
                if found:
                    return found
    if isinstance(payload, list) and payload:
        return _find_url(payload[0], path)
    return ""


class ImageHost:
    """一个可以接收图片并返回公网地址的图床。"""

    #: True 表示国内访问快
    domestic = False

    def __init__(self, name: str, url: str, field: str = "file", api_key: str = ""):
        self.name = name
        self.url = url
        self.field = field
        self.api_key = api_key

    def upload(self, content: bytes, filename: str, mime_type: str) -> str:
        raise NotImplementedError

    def _post(self, body: bytes, content_type: str) -> bytes:
        request = Request(
            self.url,
            data=body,
            headers={
                "Content-Type": content_type,
                "User-Agent": "ShopifyTableGen/1.5",
                "Connection": "close",
            },
            method="POST",
        )
        with urlopen(request, timeout=UPLOAD_TIMEOUT_SECONDS) as response:
            return response.read()


class FreeImageHost(ImageHost):
    """freeimage.host (iili.io) —— 最慷慨的免费图床。

    实测一批 30 张图、8 张并发，30 张全部成功，单张约 0.6 秒。
    不需要账号，图片长期保留，单文件上限 64MB。
    """

    def __init__(self, base: str = "https://freeimage.host/api/1/upload",
                 api_key: str = "", name: str = "FreeImage"):
        super().__init__(name, base, "source", api_key or "6d207e02198a847aa98d0a2a901485a5")

    def upload(self, content: bytes, filename: str, mime_type: str) -> str:
        body, content_type = _multipart(
            self.field, filename, content, mime_type, extra={"key": self.api_key}
        )
        result = json.loads(self._post(body, content_type).decode("utf-8"))
        image = result.get("image") or {}
        if result.get("status_code") != 200 or not image.get("url"):
            error = result.get("error") or result
            raise ImageHostError(f"{self.name} 拒绝了该文件: {error}")
        return str(image["url"])


class BeeImgHost(ImageHost):
    """BeeIMG —— 国内图床，速度快，不需要 key。

    免费版拒绝 1MB 以上的文件，所以图片必须先压缩。
    """

    domestic = True

    def __init__(self, base: str = "https://beeimg.com/api/upload/file/json/",
                 api_key: str = "", name: str = "BeeIMG"):
        super().__init__(name, base, "file", api_key)

    def upload(self, content: bytes, filename: str, mime_type: str) -> str:
        extra = {"apikey": self.api_key} if self.api_key else None
        body, content_type = _multipart(self.field, filename, content, mime_type, extra=extra)
        result = json.loads(self._post(body, content_type).decode("utf-8"))
        files = result.get("files") or {}
        if str(files.get("code")) != "200" or not files.get("url"):
            raise ImageHostError(
                f"{self.name} 拒绝了该文件: {files.get('status', result)}"
            )
        return str(files["url"])


class UguuHost(ImageHost):
    """Uguu —— Ozon 助手最早用的图床，快但在海外。"""

    def __init__(self):
        super().__init__("Uguu", "https://uguu.se/upload.php", "files[]")

    def upload(self, content: bytes, filename: str, mime_type: str) -> str:
        body, content_type = _multipart(self.field, filename, content, mime_type)
        result = json.loads(self._post(body, content_type).decode("utf-8"))
        if not result.get("success") or not result.get("files"):
            raise ImageHostError(f"{self.name} 拒绝了该文件: {result}")
        return str(result["files"][0]["url"])


class CustomHost(ImageHost):
    """自建图床 / 兼容通用 multipart 协议的自有服务。

    地址就是一个普通 multipart 接口：文件放 ``file`` 字段，
    返回 JSON 里带 ``url`` 即可。字段名和取值路径都可调。
    """

    domestic = True

    def __init__(self, url: str, api_key: str = "", name: str = "自建图床",
                 field: str = "file", url_path: str = "url", header: str = ""):
        super().__init__(name, url, field, api_key)
        self.url_path = url_path
        self.header = header

    def upload(self, content: bytes, filename: str, mime_type: str) -> str:
        body, content_type = _multipart(self.field, filename, content, mime_type)
        request = Request(
            self.url,
            data=body,
            headers={
                "Content-Type": content_type,
                "User-Agent": "ShopifyTableGen/1.5",
                "Connection": "close",
                **({"Authorization": self.header} if self.header else {}),
            },
            method="POST",
        )
        with urlopen(request, timeout=UPLOAD_TIMEOUT_SECONDS) as response:
            raw = response.read().decode("utf-8")
        try:
            result = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ImageHostError(f"{self.name} 没有返回 JSON: {raw[:200]}") from exc
        found = _find_url(result, self.url_path)
        if not found:
            raise ImageHostError(f"{self.name} 没有返回图片地址: {raw[:200]}")
        return found


def default_hosts() -> list[ImageHost]:
    """默认图床顺序（按实测可靠性与国内速度排列）。"""
    return [
        FreeImageHost(),
        BeeImgHost(),
        UguuHost(),
        BeeImgHost("https://img.beeimg.cn/api/upload/file/json/", name="BeeIMG-CN"),
    ]


def _is_quota_exhausted(exc: Exception) -> bool:
    text = str(exc).lower()
    return any(marker in text for marker in QUOTA_MARKERS)


def _is_retryable(exc: Exception) -> bool:
    """这个失败值不值得再试一次？"""
    text = str(exc).lower()
    if any(marker in text for marker in PERMANENT_MARKERS):
        return False
    if _is_quota_exhausted(exc):
        return False
    return any(marker in text for marker in RETRYABLE_MARKERS)


def upload_to_host(
    path: str | Path,
    hosts: list[ImageHost] | None = None,
    timeout_seconds: float = UPLOAD_TIMEOUT_SECONDS,
    attempts: int = UPLOAD_ATTEMPTS,
    on_progress=None,
) -> str:
    """把一张图片上传到公网图床，返回可访问的 URL。

    图片先压缩再上传；图床依次尝试，某家失败就换下一家；
    网络类错误重试若干次，被拒绝/超限这类错误立即换人。

    失败时抛 :class:`ImageHostError`。
    """
    source = Path(path)
    content, mime_type = compress_image(source)
    candidates = list(hosts) if hosts is not None else default_hosts()
    problems: list[str] = []
    exhausted: set[str] = set()

    for host in candidates:
        if host.name in exhausted:
            continue
        for attempt in range(1, max(1, attempts) + 1):
            logger.info("开始上传图片 %s via %s (第 %s 次)", source.name, host.name, attempt)
            try:
                image_url = host.upload(content, source.name, mime_type)
            except Exception as exc:
                problems.append(f"{host.name}: {exc}")
                if _is_quota_exhausted(exc):
                    exhausted.add(host.name)
                    logger.warning("图床 %s 配额已用尽，本批次不再使用", host.name)
                    break
                if attempt >= max(1, attempts) or not _is_retryable(exc):
                    logger.warning("上传到 %s 失败 (%s)，换下一家", host.name, exc)
                    break
                wait = UPLOAD_BACKOFF_SECONDS * attempt + random.uniform(0, 0.8)
                logger.warning("上传到 %s 失败 (%s)，%.1f 秒后重试", host.name, exc, wait)
                time.sleep(wait)
                continue
            logger.info("图片上传完成: %s -> %s", source.name, image_url)
            if on_progress:
                on_progress(source.name, image_url)
            return image_url

    detail = "; ".join(problems[:4])
    raise ImageHostError(f"图片托管失败: {detail}")


class ImageUploader:
    """上传图片并记住每个文件已经传过的地址。

    一个 99 款商品的文件夹若共用 4 张详情图，不缓存的话要传 495 次，
    有缓存只传 103 次：每款一张主图 + 共用的详情图。

    线程安全：正在上传的文件也会被记住，多线程共用同一张图时不会重复上传。
    """

    def __init__(self, timeout_seconds: float = UPLOAD_TIMEOUT_SECONDS, uploader=None):
        self.timeout_seconds = timeout_seconds
        self._uploader = uploader or (lambda p: upload_to_host(p, timeout_seconds=self.timeout_seconds))
        self._cache: dict[str, str] = {}
        self._in_flight: dict[str, threading.Event] = {}
        self._lock = threading.Lock()
        self.uploaded = 0

    def upload(self, path: str | Path) -> str:
        key = str(path)
        while True:
            with self._lock:
                cached = self._cache.get(key)
                if cached is not None:
                    return cached
                event = self._in_flight.get(key)
                if event is None:
                    event = threading.Event()
                    self._in_flight[key] = event
                    owner = True
                else:
                    owner = False
            if not owner:
                # 别的线程正在传这张图，等它传完直接用结果
                event.wait(timeout=self.timeout_seconds + 30)
                with self._lock:
                    cached = self._cache.get(key)
                if cached is not None:
                    return cached
                continue
            try:
                url = self._uploader(path)
            finally:
                with self._lock:
                    self._in_flight.pop(key, None)
                event.set()
            with self._lock:
                self._cache[key] = url
                self.uploaded += 1
            return url

    def upload_many(self, paths) -> list[str]:
        return [self.upload(p) for p in paths]
