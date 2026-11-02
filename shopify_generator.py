# -*- coding: utf-8 -*-
"""
Shopify表格生成 1.5  —  自动上新流程中的「表格生成」环节
============================================================================
本文件是从已编译的 shopify.exe (PyInstaller / Python 3.12) 中逆向复现的源码，
并按用户后续要求做了三处增强（下面「与原版的差异」一节逐条列出）。

【工具用途】
读取一份「数据表」(xlsx，内含 表格模板 / 依赖文件 / 标题文件 / Sheet1 四张表)、
一个「目标文件夹」(里面按 款号/parent_sku 分层放商品图片)、一个「尺码文件夹」
(各款式的尺码表图片，文件名=款号.jpg)，把图片上传到工厂图片服务器或公共图床，
最终生成上架表格，并另存为 <目标文件夹名>.xlsx。

【依赖】
    pip install openpyxl requests pillow

【运行】
    python shopify_generator.py
（重新打包成 shopify.exe：pyinstaller --onefile --noconsole --name shopify shopify_generator.py）

---------------------------------------------------------------------------
与原版的差异（本次新增/修改）
  1) 【结果文件只留 Sheet1】
     原版保存时把整本工作簿写出去，结果里带着 表格模板/依赖文件/标题文件 三张
     无关的表。现在结果文件**只保留 Sheet1 一张表**，且 Sheet1 采用与
     上架/DJLLWY.xlsx 完全一致的 53 列 TEMU 上架格式：
       产品标题 / 英文标题 / 产品描述 / 产品货号 / 变种名称 / 变种属性名称一二 /
       变种属性值一二 / 预览图 / 申报价格 / SKU货号 / 长宽高 / 重量 /
       识别码 / 站外产品链接 / 轮播图 / 产品素材图 / 外包装形状 / 外包装类型 /
       外包装图片 / 建议零售价 / 库存 / 发货时效 / 分类id / 产品属性 /
       SPU属性 / SKC属性 / SKU属性 / 站点价格 / 来源url / 产地 / 敏感属性 /
       备注 / SKU分类 / SKU分类数量 / SKU分类单位 / 独立包装 / 包装清单 /
       生命周期 / 视频Url / 运费模板 / 经营站点
     行数 = 每个 SKU 一行（不是原版的尺码数² 笛卡尔积），SKU货号 =
     图片名前缀 + "-" + 该行的「变种属性值二」。

  2) 【新增公共图床上传功能】
     新增 image_host.py：多图床容灾（FreeImage → BeeIMG → Uguu → 自建），
     方案参考用户现有项目 跨境电商/跨境电商/ozon_uploader/media.py。
     图床地址会写进结果 xlsx 的「预览图 / 轮播图 / 产品素材图 / 外包装图片 /
     SKC属性.previewImgUrls」等图片字段。工厂图床（config.json 的 host）
     仍然保留，可通过 config.json 的 upload_target 切换。

  3) 【SKU货号规则】
     按用户更正：SKU货号 = 图片名前缀 + "-" + 该行的「变种属性值二」。
     图片前缀从图片名反推：DJLLWY-MDL26050801-1.jpg → DJLLWY-MDL26050801。
     款号取前缀第一段（DJLLWY），用于匹配依赖文件。

---------------------------------------------------------------------------
复现保真度说明
  * 图片上传协议：来自 upload_image 字节码（POST host+'uploadImg'，files={'file':二进制}，
    data={'filename':文件名, 'user_key':''}，返回 host+'show_full_image/'+res['file']）。
  * 标题算法：random.sample + 空格拼接、长度 < 100，来自 create_title 字节码。
  * 依赖文件 / 标题文件解析：来自 read_template 字节码。
  * 目录层级约定：目标/款号/parent_sku/图片，来自 show_directory 字节码。

原版行为保留但已容错（代码内标注 #GUARD）：
  #GUARD-1  标题文件里出现但依赖文件里没有的款号，原版会 KeyError 崩溃 —— 改为跳过。
  #GUARD-2  变体图缺失时原版会 KeyError 崩溃 —— 改为留空。
  #GUARD-3  上传失败原版会「无限重试」导致卡死 —— 改为最多重试 3 次。
"""

import os
import re
import json
import sys
import random
import threading
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as filedialog
import tkinter.messagebox
import logging

import openpyxl
import requests

import image_host

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


# ---------------------------------------------------------------------------
# 运行日志文件（ground truth）
#   日志窗口依赖 tkinter，且 pywinauto 这类外部工具读不到它的内容。
#   这里额外把每条 log 写到「工作目录/shopify_run.log」，方便排障与自动化验证。
# ---------------------------------------------------------------------------
_RUN_LOG = None


def run_log(msg):
    global _RUN_LOG
    if _RUN_LOG is None:
        try:
            _RUN_LOG = open(os.path.join(os.getcwd(), "shopify_run.log"),
                            "a", encoding="utf-8")
        except Exception:
            _RUN_LOG = False
            return
    if _RUN_LOG is False:
        return
    try:
        _RUN_LOG.write(str(msg) + "\n")
        _RUN_LOG.flush()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 配置（从同目录 config.json 读取）
#   host           工厂图床地址
#   factory_name   厂名
#   upload_target  图片上传目标："factory"=工厂图床 / "public"=公共图床 / "both"=两者都传
#   public_hosts   公共图床列表（留空用默认顺序）
# ---------------------------------------------------------------------------
def _base_dir():
    """程序所在目录。

    PyInstaller 打包成 onefile 后，__file__ 指向临时解包目录
    (C:\\Users\\...\\Temp\\_MEIxxxx)，配置文件都在那里找不到。
    所以打包运行时要用 exe 自己的目录。
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


_BASE_DIR = _base_dir()
#: 配置文件。按顺序读取，后面的覆盖前面的。
#:   config.json    —— 原版自带的配置（厂名 + 工厂图床地址）
#:   settings.json  —— 可选，放 upload_target 等新配置项
#:                   （config.json 有时被 WPS/Excel 锁住无法写入，
#:                    新配置项放这里就能随时改，不受锁影响）
_CFG_FILES = ("config.json", "settings.json")
_cfg = {}
for _name in _CFG_FILES:
    _path = os.path.join(_BASE_DIR, _name)
    if not os.path.exists(_path):
        continue
    try:
        with open(_path, "r", encoding="utf-8-sig") as _f:
            _data = json.load(_f)
        if isinstance(_data, dict):
            _cfg.update(_data)
    except Exception as e:  # pragma: no cover
        logging.warning("读取 %s 失败: %s", _name, e)

jit_url = _cfg.get("host", "http://218.65.17.106:4999/").rstrip("/") + "/"
factory_name = _cfg.get("factory_name", "")
#: 图片上传目标："public_first"=公共图床优先（默认，仿 ozon 方案，
#:               全部失败回退工厂图床）/ "factory"=只传工厂图床 /
#:               "public"=只传公共图床 / "both"=两者都传
#: 想改的话，在同目录 settings.json 写 {"upload_target": "both"} 即可。
upload_target = _cfg.get("upload_target", "public_first")
#: 公共图床（None 表示用 image_host.default_hosts() 的默认顺序）
public_hosts = _cfg.get("public_hosts") or None

# 工作表名
SHEET_TEMPLATE = "表格模板"   # 输入：Shopify 33 列参考表
SHEET_DEPEND = "依赖文件"     # 输入：每款式主数据
SHEET_TITLE = "标题文件"     # 输入：每款式标题词
SHEET_OUTPUT = "Sheet1"       # 输出：53 列 TEMU 上架表（结果文件只保留这一张）

# 线程数（原版用 hash(款号)%thread_number 分派；单线程最稳妥）
THREAD_NUMBER = 1

# 全局状态
style = {}          # 款号 -> {class_id, attribute, global_product, price, storage,
                    #         weight, length, width, height, status, delivery, brand,
                    #         size:[...], title:[...]}
size_table = {}     # 款号 -> 尺码表图片 URL
success_img = {}    # 预留

# GUI 控件（运行时由 build_gui 赋值）
text = None
var_name = None        # 数据表路径
var_name2 = None       # 目标文件夹路径
var_name_style = None  # 尺码文件夹路径
converter_thread = None


#: 主窗口句柄。子线程要更新日志时通过它把操作投递回主线程执行。
_root_window = None


def log(msg):
    """写日志。

    ⚠️ 子线程不能直接操作 tkinter 控件（Text.insert / StringVar.set 都不是
    线程安全的），否则会卡在「main thread is not in main loop」。
    这里用 after(0, ...) 把写入投递回主线程，主线程空闲时就会执行。
    """
    line = str(msg) + "\n"
    logging.info("%s", msg)
    run_log(msg)
    win = _root_window
    if win is None:
        return
    try:
        win.after(0, _append_log, line)
    except Exception:
        pass


def _append_log(line):
    """在主线程里往日志区写一行。"""
    if text is None:
        return
    try:
        text.insert(tk.END, line)
        text.see(tk.END)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 图片上传
# ---------------------------------------------------------------------------
def upload_to_factory(path, file_name, _retry=0):
    """上传单张图片到工厂图床，返回可访问的图片 URL。逻辑来自 upload_image 字节码。"""
    try:
        with open(path, "rb") as fp:
            files = {"file": fp}
            payload = {"filename": file_name, "user_key": ""}
            # 工厂图床是直连 IP，必须绕过系统代理 —— 否则本机代理软件
            # （FlClash 等）会把请求拦下，报 ProxyError 10053/10054。
            res = json.loads(
                requests.post(
                    jit_url + "uploadImg",
                    files=files,
                    data=payload,
                    proxies={"http": None, "https": None},
                    timeout=60,
                ).content
            )
        url = jit_url + "show_full_image/" + res["file"].replace("\\", "/")
        log(url + "  " + file_name)
        return url
    except Exception as e:
        # GUARD-3: 原版会无限重试卡死，这里最多重试 3 次
        if _retry < 3:
            logging.warning("照片上传失败，重新上传: %s", e)
            log("照片上传失败，重新上传")
            return upload_to_factory(path, file_name, _retry + 1)
        logging.error("照片上传最终失败: %s", e)
        raise


def upload_to_public(path, file_name):
    """上传到公共图床（FreeImage/BeeIMG/Uguu… 多家容灾），返回公网 URL。

    实现见 image_host.py，方案参考 跨境电商/跨境电商/ozon_uploader/media.py。
    """
    hosts = None
    if public_hosts:
        hosts = [image_host.CustomHost(h["url"], h.get("api_key", ""),
                                        h.get("name", "自建图床"),
                                        h.get("field", "file"))
                 for h in public_hosts]
    url = image_host.upload_to_host(path, hosts=hosts)
    log(url + "  " + file_name)
    return url


def upload_image(path, file_name):
    """按配置把图片传到图床，返回最终写进表格的 URL。

    - public_first（默认，仿 ozon 方案）：先走公共图床
      （FreeImage → BeeIMG → Uguu → BeeIMG-CN 依次容灾），
      全部失败才回退工厂图床（218.65.17.106:4999）
    - factory：只传工厂图床
    - public ：只传公共图床
    - both   ：两者都传，表格里用公共图床的地址（公网可访问），
                工厂地址另存到「来源url」列
    """
    factory_url = ""
    public_url = ""
    if upload_target == "public_first":
        try:
            public_url = upload_to_public(path, file_name)
            return public_url, ""
        except Exception as e:
            logging.error("公共图床全部失败，回退工厂图床 %s: %s", file_name, e)
            log("公共图床全部失败，改用工厂图床 " + file_name)
            factory_url = upload_to_factory(path, file_name)
            return factory_url, factory_url
    if upload_target in ("factory", "both"):
        factory_url = upload_to_factory(path, file_name)
    if upload_target in ("public", "both"):
        try:
            public_url = upload_to_public(path, file_name)
        except Exception as e:
            logging.error("公共图床上传失败 %s: %s", file_name, e)
            log("公共图床上传失败 " + file_name)
            if not factory_url:
                raise
    return public_url or factory_url, factory_url


# ---------------------------------------------------------------------------
# 标题生成（来自 create_title 字节码）
# ---------------------------------------------------------------------------
def create_title(selected_style):
    """从 style[selected_style]['title'] 中随机抽样若干词，空格拼接，总长 < 100。"""
    words = style[selected_style].get("title") or []
    if not words:
        return ""
    sl = len(words)
    while True:
        rn = random.randint(1, sl)
        r_title = random.sample(words, rn)
        title_str = ""
        for i in r_title:
            title_str = title_str + i + " "
        if len(title_str.strip()) < 100:
            return title_str.strip()


# ---------------------------------------------------------------------------
# 读取数据表（来自 read_template 字节码）
# ---------------------------------------------------------------------------
def read_template(template_path):
    file_path = (template_path or "").strip()
    if not file_path:
        raise ValueError("请先选择数据表(xlsx)")
    if not os.path.isfile(file_path):
        raise FileNotFoundError("数据表不存在: " + file_path)
    wb = openpyxl.load_workbook(file_path)
    style.clear()

    # ---- 依赖文件 ----
    sheet = wb[SHEET_DEPEND]
    row = 2
    while True:
        key = sheet.cell(row=row, column=1).value
        if key is None:
            break
        style[key] = {
            "class_id": sheet.cell(row=row, column=2).value,
            "attribute": sheet.cell(row=row, column=3).value,
            "global_product": sheet.cell(row=row, column=4).value,
            "price": sheet.cell(row=row, column=5).value,
            "storage": sheet.cell(row=row, column=6).value,
            "weight": sheet.cell(row=row, column=7).value,
            "length": sheet.cell(row=row, column=8).value,
            "width": sheet.cell(row=row, column=9).value,
            "height": sheet.cell(row=row, column=10).value,
            "status": sheet.cell(row=row, column=11).value,
            "delivery": sheet.cell(row=row, column=12).value,
            "brand": sheet.cell(row=row, column=13).value,
            "size": [],
            "title": [],
        }
        col = 14
        while True:
            val = sheet.cell(row=row, column=col).value
            if val is None:
                break
            style[key]["size"].append(val)
            col += 1
        row += 1

    # ---- 标题文件（转置：第1行是款号，往下各列是该款号的标题词）----
    sheet2 = wb[SHEET_TITLE]
    col = 1
    while True:
        key = sheet2.cell(row=1, column=col).value
        if key is None:
            break
        # GUARD-1: 原版直接 style[key]['title']=[] 会在 key 不在 style 时崩溃
        if key not in style:
            style[key] = {"title": []}
        else:
            style[key]["title"] = []
        row = 3  # 字节码确认从第 3 行开始读标题词（第 2 行被跳过）
        while True:
            val = sheet2.cell(row=row, column=col).value
            if val is None:
                break
            style[key]["title"].append(val)
            row += 1
        col += 1

    log("读取依赖文件完成")


# ---------------------------------------------------------------------------
# 读取尺码表图片（来自 read_style_table 字节码）
# ---------------------------------------------------------------------------
def read_style_table(size_path):
    file_path = (size_path or "").strip()
    if not file_path:
        raise ValueError("请先选择尺码文件夹")
    if not os.path.isdir(file_path):
        raise NotADirectoryError("尺码文件夹不存在: " + file_path)
    image_list = os.listdir(file_path)
    for i in image_list:
        image_path = os.path.join(file_path, i)
        if not os.path.isfile(image_path):
            continue
        suffix = i.split(".")[-1]
        mutation_suffix = i[: -len(suffix) - 1]  # 去掉 .后缀，得到“文件名无扩展名”
        url, _factory = upload_image(image_path, i)
        size_table[mutation_suffix] = url
    log("上传尺码表照片完成")


# ---------------------------------------------------------------------------
# 图片名解析
# ---------------------------------------------------------------------------
IMG_EXT = (".jpg", ".jpeg", ".png", ".webp", ".gif")


def _is_image(name):
    return name.lower().rstrip().endswith(IMG_EXT)


def _sku_prefix_of(base_name):
    """图片基名 -> SKU 前缀。

    例： DJLLWY-MDL26050801-1 -> DJLLWY-MDL26050801（去掉末尾的 '-序号'）
        只有末尾是数字时才剥掉，避免把 'DJLLWY-MDL26050801' 本身截短。
        若文件名是「颜色,尺码」写法（老约定），前缀取逗号前的颜色名，
        这种情况由调用方改用文件夹里的款号。
    """
    if "," in base_name:
        return base_name.split(",")[0]
    stripped = re.sub(r"-\d+$", "", base_name)
    return stripped or base_name


#: 通用图片目录名（这些目录里的图是公用详情图，不算独立 SKU）
COMMON_DIRS = {"通用图片", "common", "公用", "公共", "详情图", "公用图片", "新建文件夹"}


def _is_common_dir(name):
    return name.strip() in COMMON_DIRS


# ---------------------------------------------------------------------------
# 结果表：53 列 TEMU 格式（列号 1-based，与 上架/DJLLWY.xlsx 一致）
# ---------------------------------------------------------------------------
OUT_COL = {
    "title": 1,          # 产品标题
    "en_title": 2,       # 英文标题
    "desc": 3,           # 产品描述
    "goods_code": 4,     # 产品货号
    "variant_name": 5,   # 变种名称
    "attr_name1": 6,     # 变种属性名称一
    "attr_value1": 7,    # 变种属性值一
    "attr_name2": 8,     # 变种属性名称二
    "attr_value2": 9,    # 变种属性值二
    "preview": 10,       # 预览图
    "price": 11,         # 申报价格
    "sku_code": 12,      # SKU货号
    "length": 13,        # 长(cm)
    "width": 14,         # 宽(cm)
    "height": 15,        # 高(cm)
    "weight": 16,        # 重量
    "barcode_type": 17,  # 识别码类型
    "barcode": 18,       # 识别码
    "out_link": 19,      # 站外产品链接
    "carousel": 20,      # 轮播图
    "material": 21,      # 产品素材图
    "pack_shape": 22,    # 外包装形状
    "pack_type": 23,     # 外包装类型
    "pack_image": 24,    # 外包装图片
    "retail_price": 25,  # 建议零售价
    "storage": 26,       # 库存
    "delivery": 27,      # 发货时效
    "class_id": 28,      # 分类id
    "attr": 29,          # 产品属性
    "spu_attr": 30,      # SPU属性
    "skc_attr": 31,      # SKC属性
    "sku_attr": 32,      # SKU属性
    "site_price": 33,    # 站点价格
    "source_url": 34,    # 来源url
    "origin": 35,        # 产地
    "sensitive": 36,     # 敏感属性
    "remark": 37,        # 备注
    "sku_class": 38,     # SKU分类
    "sku_class_qty": 39, # SKU分类数量
    "sku_class_unit": 40,# SKU分类单位
    "independ_pack": 41, # 独立包装
    "pack_list": 42,     # 包装清单
    "lifecycle": 43,     # 生命周期
    "video_url": 44,     # 视频Url
    "freight_tpl": 45,   # 运费模板（模板id）
    "site": 46,          # 经营站点
}


def _set(sheet, row, key, value):
    sheet.cell(row=row, column=OUT_COL[key]).value = value


def _num(value):
    """数值统一成一位小数（30 -> 30.0），与参考产物一致。"""
    if value is None or value == "":
        return value
    try:
        f = float(value)
    except (TypeError, ValueError):
        return value
    if f == int(f):
        return float(f)
    return f


def _blank(value):
    """空值统一成空串。

    参考产物 上架/DJLLWY.xlsx 里这些列是「首行空串、后续行 None」，
    那是 Excel 另存时留下的痕迹，语义上都是空。这里统一成空串。
    """
    return "" if value is None else value


def _attr_value(attr_json, prop_name):
    """从「产品属性」JSON 里取某个属性的值。

    例：产品属性 JSON 里 {"propName":"适用人群","propValue":"女童"} -> "女童"
    """
    if not attr_json:
        return ""
    try:
        data = json.loads(attr_json)
    except Exception:
        return ""
    if isinstance(data, list):
        for item in data:
            if isinstance(item, dict) and item.get("propName") == prop_name:
                return str(item.get("propValue") or "")
    return ""


def _derive_product_title(attr_json):
    """推导中文产品标题：适用人群 + 类型。

    参考产物 上架/DJLLWY.xlsx 里 产品标题 = '女童夹克'，
    正是素材库「产品属性」中 适用人群=女童 与 类型=夹克 的组合。
    """
    people = _attr_value(attr_json, "适用人群")
    kind = _attr_value(attr_json, "类型")
    title = (people + kind).strip()
    return title or None


def _update_skc_attr(raw, image_url):
    """SKC属性里的 previewImgUrls 换成上传后的图片地址。

    原值形如 [{"parentSpecId":1001,"parentSpecName":"颜色","specId":21263,
    "specName":"混合色","extCode":"","productSkcId":"","previewImgUrls":"||||"}]
    """
    if not raw:
        return raw
    try:
        data = json.loads(raw)
    except Exception:
        return raw
    if isinstance(data, list):
        for item in data:
            if isinstance(item, dict):
                if "previewImgUrls" in item:
                    item["previewImgUrls"] = image_url
                if "extCode" in item and not item["extCode"]:
                    item["extCode"] = "desktop"
    # 紧凑写法，与素材库/参考产物里的 JSON 字符串格式一致
    return json.dumps(data, ensure_ascii=False, separators=(",", ":"))

def _update_sku_attr(raw, size):
    """SKU属性里的尺码项换成当前行的尺码。"""
    if not raw:
        return raw
    try:
        data = json.loads(raw)
    except Exception:
        return raw
    if isinstance(data, list):
        for item in data:
            if isinstance(item, dict) and item.get("parentSpecName") == "尺码":
                item["specName"] = size
    return json.dumps(data, ensure_ascii=False, separators=(",", ":"))


def _template_row_for_sku(lib_rows, class_id, size):
    """在数据表 Sheet1 里找一条与（分类id, 尺码）匹配的参考行。

    数据表的 Sheet1 是 TEMU 上架表的「素材库」：25000 行、5000 个款式、每款 5 个尺码，
    里面带着完整的 产品属性 / SPU属性 / SKU属性 JSON。生成结果行时从这里取一行做底，
    再覆盖款式相关的字段即可（与 上架/DJLLWY.xlsx 的做法一致）。
    """
    for r in lib_rows:
        if str(r.get("分类id")) == str(class_id) and str(r.get("变种属性值二")) == str(size):
            return r
    return None


def show_directory(path, hash_num, template_path=None):
    """按目标文件夹生成结果表（图片驱动）。

    目录约定：目标/款号/parent_sku/图片。款号也可由图片名反推（前缀第一段）。
    每个 parent_sku 一组，生成「一个 SKU 一行」的结果行。
    """
    # 模板路径由主线程传入 —— 子线程不能碰 tkinter 的 StringVar
    wb_in = openpyxl.load_workbook((template_path or "").strip(), data_only=True)

    # ---- 读 Sheet1 素材库（只取需要的列）----
    lib_rows = []
    if SHEET_OUTPUT in wb_in.sheetnames:
        lib = wb_in[SHEET_OUTPUT]
        header = [c.value for c in lib[1]]
        idx = {name: i for i, name in enumerate(header) if name}
        need = ["产品标题", "英文标题", "产品描述", "产品货号", "变种名称",
                "变种属性名称一", "变种属性值一", "变种属性名称二", "变种属性值二",
                "申报价格", "长", "宽", "高", "重量", "库存", "发货时效",
                "分类id", "产品属性", "SPU属性", "SKC属性", "SKU属性", "站点价格",
                "产地", "敏感属性", "备注", "外包装形状", "外包装类型", "外包装图片",
                "建议零售价(建议零售价币种)", "SKU分类", "SKU分类数量",
                "SKU分类单位", "独立包装", "包装清单", "生命周期", "视频Url",
                "运费模板（模板id）", "经营站点"]
        colmap = {n: idx[n] for n in need if n in idx}
        for row in lib.iter_rows(min_row=2, values_only=True):
            if not any(v is not None for v in row):
                continue
            rec = {n: row[i] for n, i in colmap.items()}
            rec.setdefault("分类id", None)
            rec.setdefault("变种属性值二", None)
            lib_rows.append(rec)
        log("素材库读取完成：%d 行" % len(lib_rows))

    # ---- 输出工作簿：只有 Sheet1 一张表 ----
    wb_out = openpyxl.Workbook()
    sheet = wb_out.active
    sheet.title = SHEET_OUTPUT
    # 表头取自数据表的 Sheet1（53 列 TEMU 上架表头）
    if SHEET_OUTPUT in wb_in.sheetnames:
        lib = wb_in[SHEET_OUTPUT]
        for c in range(1, 54):
            sheet.cell(row=1, column=c).value = lib.cell(row=1, column=c).value

    row = 2

    # 1) 收集图片，按 (款号, SKU前缀, parent_sku) 分组
    groups = {}
    for root, _dirs, files in os.walk(path):
        rel = os.path.relpath(root, path)
        parts = [p for p in rel.replace("\\", "/").split("/") if p not in (".", "")]
        parent_sku = parts[-1] if len(parts) >= 2 else ""
        is_attach = os.path.basename(root.rstrip("/\\")) == "通用图片"
        for f in files:
            if not _is_image(f):
                continue
            full = os.path.join(root, f)
            base = f[: f.rfind(".")]
            prefix = _sku_prefix_of(base)
            sc = prefix.split("-")[0]
            key = (sc, prefix, parent_sku)
            groups.setdefault(key, {"main": [], "attach": []})
            groups[key]["attach" if is_attach else "main"].append(full)

    # 2) 逐组处理
    for (current_style, sku_prefix, parent_sku), g in groups.items():
        if hash(current_style) % THREAD_NUMBER != hash_num:
            continue
        if current_style not in style:
            log("依赖文件中不存在 " + current_style + " 跳过")
            continue
        if current_style not in size_table:
            log("尺码表图片中不存在 " + current_style + " 跳过")
            continue

        st = style[current_style]
        title = create_title(current_style)      # 英文标题

        # ---- 上传图片 ----
        mutation_dir = {}
        for p in g["main"]:
            fn = os.path.basename(p)
            base = fn[: fn.rfind(".")]
            url, factory_url = upload_image(p, fn)
            mutation_dir[base] = {"url": url, "factory": factory_url}

        attach = []
        for p in g["attach"]:
            fn = os.path.basename(p)
            url, factory_url = upload_image(p, fn)
            attach.append((url, factory_url))

        all_urls = [m["url"] for m in mutation_dir.values()] + [a[0] for a in attach]
        factory_urls = [m["factory"] for m in mutation_dir.values()] + [a[1] for a in attach]
        # 多张图用 '|' 连接（与参考产物 上架/DJLLWY.xlsx 一致）
        joined_urls = "|".join(all_urls)
        main_url = all_urls[0] if all_urls else ""
        factory_main = factory_urls[0] if factory_urls else ""

        # ---- 每个 SKU 一行 ----
        for size in st["size"]:
            ref = _template_row_for_sku(lib_rows, st["class_id"], size) or {}

            # 产品标题：素材库「产品属性」里的 适用人群+类型（与参考产物一致）
            product_title = (_derive_product_title(ref.get("产品属性"))
                            or parent_sku or current_style)

            _set(sheet, row, "title", product_title)
            _set(sheet, row, "en_title", title)
            # 产品描述 / 产品货号：沿用素材库原值
            _set(sheet, row, "desc", _blank(ref.get("产品描述")))
            _set(sheet, row, "goods_code", _blank(ref.get("产品货号")))
            # 变种名称：直接沿用素材库原值（如 '花色，S'），
            # 里面的尺码字母(S/M/L/XL/XXL) 与 变种属性值二(5-6Y…) 是两套标准，不能混用。
            color_name = ref.get("变种属性值一") or st["attribute"]
            variant_name = ref.get("变种名称")
            if not variant_name:
                variant_name = "%s，%s" % (color_name, size)
            _set(sheet, row, "variant_name", variant_name)
            _set(sheet, row, "attr_name1", "颜色")
            _set(sheet, row, "attr_value1", color_name)
            _set(sheet, row, "attr_name2", "尺码")
            _set(sheet, row, "attr_value2", size)
            _set(sheet, row, "preview", _blank(ref.get("预览图")))
            # 申报价格 / 库存 / 发货时效 / 重量 / 长宽高：全部沿用素材库原值
            # （素材库里的申报价格是 27、发货时效 9、重量按尺码 170~250，
            #   参考产物 上架/DJLLWY.xlsx 用的就是这套值）
            _set(sheet, row, "price", ref.get("申报价格", st["price"]))
            # SKU货号 = 图片名前缀 + "-" + 该行的变种属性值二
            _set(sheet, row, "sku_code", sku_prefix + "-" + size)
            _set(sheet, row, "length", ref.get("长", st["length"]))
            _set(sheet, row, "width", ref.get("宽", st["width"]))
            _set(sheet, row, "height", ref.get("高", st["height"]))
            _set(sheet, row, "weight", ref.get("重量", st["weight"]))
            _set(sheet, row, "barcode_type", _blank(ref.get("识别码类型")))
            _set(sheet, row, "barcode", _blank(ref.get("识别码")))
            _set(sheet, row, "out_link", _blank(ref.get("站外产品链接")))
            _set(sheet, row, "carousel", joined_urls)
            _set(sheet, row, "material", joined_urls)
            _set(sheet, row, "pack_shape", ref.get("外包装形状", "长方体"))
            _set(sheet, row, "pack_type", ref.get("外包装类型", "软包装+软物"))
            _set(sheet, row, "pack_image", ref.get("外包装图片", ""))
            _set(sheet, row, "retail_price", ref.get("建议零售价(建议零售价币种)"))
            _set(sheet, row, "storage", ref.get("库存", st["storage"]))
            # 发货时效取素材库值（该列在素材库里是统一值）
            _set(sheet, row, "delivery", ref.get("发货时效", st["delivery"]))
            _set(sheet, row, "class_id", ref.get("分类id", st["class_id"]))
            _set(sheet, row, "attr", ref.get("产品属性"))
            _set(sheet, row, "spu_attr", ref.get("SPU属性"))
            _set(sheet, row, "skc_attr", _update_skc_attr(ref.get("SKC属性"), joined_urls))
            _set(sheet, row, "sku_attr", _update_sku_attr(ref.get("SKU属性"), size))
            _set(sheet, row, "site_price", ref.get("站点价格"))
            _set(sheet, row, "source_url", _blank(ref.get("来源url")))
            _set(sheet, row, "origin", _blank(ref.get("产地")))
            _set(sheet, row, "sensitive", _blank(ref.get("敏感属性")))
            _set(sheet, row, "remark", _blank(ref.get("备注")))
            _set(sheet, row, "sku_class", _blank(ref.get("SKU分类", "单品")))
            _set(sheet, row, "sku_class_qty", ref.get("SKU分类数量", 1))
            _set(sheet, row, "sku_class_unit", _blank(ref.get("SKU分类单位", "件")))
            _set(sheet, row, "independ_pack", _blank(ref.get("独立包装", "是")))
            _set(sheet, row, "pack_list", _blank(ref.get("包装清单", "[]")))
            _set(sheet, row, "lifecycle", _blank(ref.get("生命周期")))
            _set(sheet, row, "video_url", _blank(ref.get("视频Url")))
            _set(sheet, row, "freight_tpl", _blank(ref.get("运费模板（模板id）")))
            _set(sheet, row, "site", _blank(ref.get("经营站点")))

            row += 1

    # 保存：结果文件只有 Sheet1 一张表
    excel_name = os.path.basename(path.rstrip("/\\")) or "output"
    wb_out.save(excel_name + ".xlsx")
    run_log("已保存结果文件: " + excel_name + ".xlsx")
    log("生成表格成功 " + excel_name + ".xlsx")


# ---------------------------------------------------------------------------
# 启动转换（来自 start_convert 字节码：按 THREAD_NUMBER 开线程）
# ---------------------------------------------------------------------------
def start_convert(template_path, target_path, size_path):
    """开始转换。

    ⚠️ 重要：三个路径必须由**主线程**取出来传进来，不能在这里读
    tkinter 的 StringVar。tkinter 的变量不是线程安全的 —— 子线程调用
    StringVar.get() 会阻塞在「main thread is not in main loop」上，
    表现就是点了「开始」之后毫无反应、日志区一片空白。
    """
    global converter_thread
    run_log("start_convert 开始: 数据表=%s 目标=%s 尺码=%s" % (
        template_path, target_path, size_path))
    read_template(template_path)
    read_style_table(size_path)
    path = target_path
    converter_thread = []
    for i in range(THREAD_NUMBER):
        t = __import__("threading").Thread(
            target=show_directory, args=(path, i, template_path))
        t.setDaemon(True)
        t.start()
        converter_thread.append(t)


# ---------------------------------------------------------------------------
# GUI（尽量还原原版窗口结构）
# ---------------------------------------------------------------------------
def build_gui():
    global text, var_name, var_name2, var_name_style, converter_thread, _root_window
    window = tk.Tk()
    _root_window = window
    window.title("Shopify表格生成 1.5")
    window.geometry("700x500")

    # 任何回调（按钮等）里未捕获的异常都写出来，避免“点了没反应、什么都没有”的情况
    def _report_exc(exc, val, tb):
        import traceback as _tb
        detail = "".join(_tb.format_exception(exc, val, tb))
        logging.error("未捕获异常: %s", detail)
        run_log("未捕获异常: " + detail)
        try:
            tk.messagebox.showerror("程序异常", detail[:2000])
        except Exception:
            pass

    window.report_callback_exception = _report_exc

    pw = tk.PanedWindow(window, orient=tk.VERTICAL)
    pw.pack(fill=tk.BOTH, expand=1)

    top = tk.Frame(pw)
    var_name = tk.StringVar()
    var_name2 = tk.StringVar()
    var_name_style = tk.StringVar()

    def choose_data():
        var_name.set(filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")]))

    def choose_target():
        var_name2.set(filedialog.askdirectory())

    def choose_size():
        var_name_style.set(filedialog.askdirectory())

    tk.Label(top, text="数据表:").grid(row=0, column=0, sticky="e")
    tk.Entry(top, textvariable=var_name, width=60).grid(row=0, column=1)
    tk.Button(top, text="选择", command=choose_data).grid(row=0, column=2)

    tk.Label(top, text="目标文件夹:").grid(row=1, column=0, sticky="e")
    tk.Entry(top, textvariable=var_name2, width=60).grid(row=1, column=1)
    tk.Button(top, text="选择", command=choose_target).grid(row=1, column=2)

    tk.Label(top, text="尺码文件夹:").grid(row=2, column=0, sticky="e")
    tk.Entry(top, textvariable=var_name_style, width=60).grid(row=2, column=1)
    tk.Button(top, text="选择", command=choose_size).grid(row=2, column=2)

    tk.Label(top, text="图片上传:").grid(row=3, column=0, sticky="e")
    ttk.Label(top, text="公共图床优先，失败自动换工厂图床（详见 settings.json）").grid(
        row=3, column=1, sticky="w")

    def on_start():
        """点「开始」。

        三个路径必须在**主线程**读取 —— tkinter 的 StringVar 不是线程安全的，
        在子线程里 get() 会直接卡死（main thread is not in main loop），
        表现就是点了按钮毫无反应。这里主线程取好再交给工作线程。
        """
        try:
            tpl = var_name.get().strip()
            tgt = var_name2.get().strip()
            siz = var_name_style.get().strip()
        except Exception as exc:
            tk.messagebox.showerror("错误", "读取输入框失败: %s" % exc)
            return

        missing = [n for n, v in (("数据表", tpl), ("目标文件夹", tgt), ("尺码文件夹", siz)) if not v]
        if missing:
            tk.messagebox.showwarning("提示", "请先选择：" + "、".join(missing))
            return

        run_log("on_start 触发, 数据表=%s 目标=%s 尺码=%s" % (tpl, tgt, siz))
        log("任务启动")
        threading.Thread(target=_run_convert, args=(tpl, tgt, siz), daemon=True).start()

    tk.Button(top, text="开始", command=on_start).grid(row=4, column=1)

    pw.add(top)

    bottom = tk.Frame(pw)
    text = tk.Text(bottom)
    text.pack(fill=tk.BOTH, expand=1)
    pw.add(bottom)

    window.mainloop()


def _run_convert(template_path, target_path, size_path):
    """工作线程：跑转换。任何异常都要让用户看见，不能默默死掉。"""
    try:
        start_convert(template_path, target_path, size_path)
        for t in list(converter_thread or []):
            t.join(timeout=1800)
    except Exception as exc:
        import traceback
        detail = traceback.format_exc()
        logging.error("任务失败: %s", detail)
        try:
            _root_window.after(0, _show_error, "任务失败", str(exc) + "\n\n" + detail)
        except Exception:
            pass


def _show_error(title, message):
    try:
        tk.messagebox.showerror(title, message)
    except Exception:
        pass


if __name__ == "__main__":
    build_gui()
