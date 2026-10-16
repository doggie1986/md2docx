# -*- coding: utf-8 -*-
"""设置层：样式预设、一份配置文件的形状、"最近打开的文件"那条表——**一段 Qt 都不碰**。

三块各管一处：

1. **样式预设**（内置技术文档 / 学术论文 / 简洁版，人还能自己存）—— :data:`BUILTIN_PRESETS`、
   :func:`presets`、:func:`add_custom`、:func:`remove_custom`、:func:`encode_custom` /
   :func:`decode_custom`。
2. **一份配置文件**（导出给另一台机器、另一版软件）—— :func:`to_document`、:func:`write_config`、
   :func:`read_config`、:func:`parse_document`。
3. **"最近打开的文件"那条表**—— :func:`push_recent`（新点中的排最前、同一个文件不占两格、
   超出上限从尾巴上砍）。

为什么不并进 :mod:`gui.preferences`：那边是 QSettings 的读写（Windows 上落在注册表里），
而注册表**不是一份能拷走、能进版本库的文件**——"导入 / 导出配置"与"自定义预设"必须有纯数据的一层，
本机存储和抄本两边共用同一份形状。这一层不 import Qt 也顺手保住了"``import main`` 不把 PySide6
拖进来"那条不变式（七期立下的），以后要接 ``--preset`` 直接接得上。

配置文件**带哪几格是要说得出理由的**：样式那三格与合并那四格跟着走；窗口大小位置、三栏宽度、
主题、输出目录、上次的工程文件路径都**不带**，逐条写在 :data:`NOT_CARRIED` 里。
认不出来就拒，并说清是哪一格——同一句理由管着这一份 .json 与工程文件那一份
（:mod:`gui.project_file`），两份 JSON 抄本各判一套就会出现"那份能导这份不能导"。
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, List, Optional, Sequence, Tuple

import docx_builder as db

__all__ = ["KIND", "VERSION", "TOP_KEYS", "SETTINGS_KEYS", "PRESET_KEYS", "NOT_CARRIED",
           "CARRY_NOTE",
           "SIZE_MIN", "SIZE_MAX", "RECENT_LIMIT",
           "StylePreset", "BUILTIN_PRESETS", "preset_names", "presets", "preset_named",
           "add_custom", "remove_custom", "builtin_names",
           "encode_custom", "decode_custom",
           "StyleValues", "to_document", "parse_document", "write_config", "read_config",
           "config_filter", "push_recent", "SettingsError"]

#: 认门用：这一份是 md2docx 的配置文件，不是别家也叫 ``.json`` 的什么东西（同工程文件那个 ``kind``）。
KIND = "md2docx.config"

#: 这一版写得出的最高版本。**读的时候只认 1 ~ VERSION**：往后的版本改了格子的形状，
#: 硬读会把"少一格"当成"关着"，那是往设置里塞假数据。
VERSION = 1

TOP_KEYS = ("kind", "version", "saved_at", "settings", "style_presets")

#: 跟着配置文件走的那七格。缺任何一格都拒，见 :func:`parse_document`。
SETTINGS_KEYS = ("body_font", "body_size", "heading_font",
                 "with_toc", "page_break", "source_label", "cover_page")
STYLE_KEYS = ("body_font", "body_size", "heading_font")
MERGE_KEYS = ("with_toc", "page_break", "source_label", "cover_page")

#: 一条预设认这四格。``builtin`` 与 ``note`` **不进抄本**：内置那三档在读的这一版里自己就带着
#: 名字和说明，抄本里再存一份"它是内置的"就成了一件事两处说；自定义档本来也没有说明。
PRESET_KEYS = ("name", "body_font", "body_size", "heading_font")

#: 不跟着走的几格，逐条写着为什么——导入那一路人一眼看到的就是这份名单。
NOT_CARRIED = {
    "window/geometry": "窗口大小与位置是这台机器那块屏幕的事，拷到另一台就是把它摆到屏幕外",
    "window/splitter_sizes": "三栏宽度同上",
    "appearance/theme": "主题管的是这个软件长什么样，不是稿子长什么样",
    "output/directory": "输出目录是本机路径，那边多半没有这个目录",
    "project/last_path": "上一份工程文件也是本机路径",
    "recent files": "最近打开的文件同理，而且它该由那台机器自己攒",
}

#: 导入那一路念给人听的一句：换了哪几格、没换哪几格。名单与 :data:`NOT_CARRIED` 是同一批格子，
#: 但这里要的是**中文叫法**——注册表键名摆到屏幕上，人读不出"输出目录"指的是哪一格。
CARRY_NOTE = ("这份抄本带的是排版三格、合并那四个开关和自定义预设；"
              "输出目录、主题、窗口尺寸与那几张路径表都不带——那些是这台机器上的事。")

#: 字号的范围。界面那一侧的旋钮读这一份，不另抄一份（两处各一套就会出现"预设填得进、旋钮拧不到"）。
SIZE_MIN = 6.0
SIZE_MAX = 72.0

#: 最近打开的文件一次记几条。是**显示闸门不是计数口径**：列表只演前面那几条，
#: 但删掉的那些不是"没打开过"，所以砍的时候从尾巴上砍、不重排前面。
RECENT_LIMIT = 8


class SettingsError(ValueError):
    """读不懂 / 写不出这一层的任何东西。句子是给**人**看的中文，带下一步该干什么。"""

    def __init__(self, message: str, source: Optional[Path] = None) -> None:
        super().__init__(message)
        self.source = Path(source) if source is not None else None

    @property
    def user_message(self) -> str:
        return str(self)


# --- 样式预设 ----------------------------------------------------------------

@dataclass(frozen=True)
class StylePreset:
    """一套排版：正文用哪支字体、多大、标题用哪支。标题各级的**字号**不在这儿——
    它由 :meth:`docx_builder.BuilderOptions.heading_size` 从正文档位推，所以档位动一档就整篇跟着动。
    """

    name: str
    body_font: str
    body_size: float
    heading_font: str
    builtin: bool = False
    note: str = ""

    def as_dict(self) -> dict:
        return {"name": self.name, "body_font": self.body_font,
                "body_size": self.body_size, "heading_font": self.heading_font}

    def describe(self) -> str:
        """给人念的一行：挑之前就该看见这套是什么排版。"""
        return "%s / %s pt，标题 %s" % (self.body_font, _size_text(self.body_size), self.heading_font)


#: 内置三档。**技术文档那一档就是现在的默认档**（取自 ``docx_builder``，不在这里抄一份数值），
#: 所以挑它等于不改，而"默认"这件事也只在 :mod:`docx_builder` 一处住着。
BUILTIN_PRESETS: Tuple[StylePreset, ...] = (
    StylePreset(name="技术文档", body_font=db.BODY_FONT_DEFAULT, body_size=db.BODY_SIZE_DEFAULT,
                heading_font=db.HEADING_FONT_DEFAULT, builtin=True,
                note="现在的默认档：正文宋体、标题黑体，投标与施工类稿子常用。"),
    StylePreset(name="学术论文", body_font="Times New Roman", body_size=12.0,
                heading_font="Times New Roman", builtin=True,
                note="正文与标题同一支衬线体，期刊投稿常要求这种配法。"),
    StylePreset(name="简洁版", body_font="等线", body_size=10.5, heading_font="等线",
                builtin=True,
                note="五号字、标题不另换字体，只靠字号差把层级拉开。"),
)


def builtin_names() -> List[str]:
    return [preset.name for preset in BUILTIN_PRESETS]


def add_custom(current: Sequence[StylePreset], preset: StylePreset) -> List[StylePreset]:
    """把一条自定义预设并进表里，交回**新的列表**（不改传入的那一份）。

    撞内置那三档的名字**当场拒**：顶掉之后"回到内置那一档"就没有地方了，而下拉里两条同名
    也点不出是哪一个。内置那三档本身不许改——它们是这一版的说明，改了就没有对照物。
    """
    name = _text(preset.name, "预设名字", allow_empty=False)
    _size(preset.body_size, "「%s」的字号" % name)
    _text(preset.body_font, "「%s」的正文字体" % name, allow_empty=False)
    _text(preset.heading_font, "「%s」的标题字体" % name, allow_empty=False)
    if name in builtin_names():
        raise SettingsError(
            "内置的「%s」已经有这一档，不能再存成同名的自定义预设。"
            "想改它就另起一个名字（比如「%s（我的）」）——内置那三档不会被改掉。" % (name, name))
    result = [item for item in current if not item.builtin]
    result = [item for item in result if item.name != name]
    result.append(StylePreset(name=name, body_font=preset.body_font, body_size=float(preset.body_size),
                              heading_font=preset.heading_font))
    return result


def remove_custom(current: Sequence[StylePreset], name: str) -> List[StylePreset]:
    """撤掉一条自定义预设。**内置那三档撤不掉**，明说一句而不是默默什么都不做。"""
    target = _text(name, "预设名字", allow_empty=False)
    if target in builtin_names():
        raise SettingsError("「%s」是内置的三档之一，撤不掉（挑别的档就等于把它换掉了）。" % target)
    rest = [item for item in current if not item.builtin and item.name != target]
    if len(rest) == len([item for item in current if not item.builtin]):
        raise SettingsError("自定义预设里没有「%s」这一档（内置的那三档不在这里）。" % target)
    return rest


def presets(custom: Sequence[StylePreset] = ()) -> List[StylePreset]:
    """内置三档在前、自定义按存的次序在后——下拉里就是这个次序。"""
    return list(BUILTIN_PRESETS) + [item for item in custom if not item.builtin]


def preset_names(custom: Sequence[StylePreset] = ()) -> List[str]:
    """下拉里那一列名字（与 :func:`presets` 同序，界面按次序配着画）。"""
    return [item.name for item in presets(custom)]


def preset_named(name: str, custom: Sequence[StylePreset] = ()) -> Optional[StylePreset]:
    for item in presets(custom):
        if item.name == name:
            return item
    return None


def encode_custom(custom: Sequence[StylePreset]) -> str:
    """自定义预设存进本机设置时的那一份文本（QSettings 只能存字符串）。"""
    return json.dumps([item.as_dict() for item in custom if not item.builtin],
                      ensure_ascii=False, sort_keys=True)


def decode_custom(text: str) -> List[StylePreset]:
    """读本机存的那一份自定义预设。空文本 = **一条都没有**（不是坏档）。

    读不懂就抛 :class:`SettingsError`——静默交回空列表等于"你自己存的那几档没了"，
    而那正是人最不会因为报错去查的一种损失。
    """
    raw = (text or "").strip()
    if not raw:
        return []
    try:
        body = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise SettingsError(
            "本机存的自定义预设读不进（第 %d 行第 %d 列：%s）。"
            "这一档先按没有处理；要重来就重新存一条。" % (exc.lineno, exc.colno, exc.msg)) from exc
    if not isinstance(body, list):
        raise SettingsError("本机存的自定义预设不是一个列表，现在是%s。" % _shape(body))
    result: List[StylePreset] = []
    for index, item in enumerate(body, 1):
        result.append(_preset_of(item, "本机存的第 %d 条预设" % index))
    names = [item.name for item in result]
    for name in names:
        if name in builtin_names():
            raise SettingsError("本机存的自定义预设里有一条叫「%s」，那是内置的档名。" % name)
    if len(set(names)) != len(names):
        raise SettingsError("本机存的自定义预设里有重名的那一条（同名就点不出是哪一档）。")
    return result


# --- 一份配置文件 ------------------------------------------------------------

@dataclass
class StyleValues:
    """配置文件里 ``settings`` 那一格认识的七样：样式三格 + 合并四格。"""

    body_font: str = db.BODY_FONT_DEFAULT
    body_size: float = db.BODY_SIZE_DEFAULT
    heading_font: str = db.HEADING_FONT_DEFAULT
    with_toc: bool = True
    page_break: bool = True
    source_label: bool = True
    cover_page: bool = False

    def as_dict(self) -> dict:
        return {key: getattr(self, key) for key in SETTINGS_KEYS}


def to_document(values: StyleValues, custom: Sequence[StylePreset] = (),
                now: Optional[datetime] = None) -> dict:
    """摆成一份要写盘的字典。``style_presets`` 只装自定义那几条（内置的在读的那一版里有）。"""
    return {
        "kind": KIND,
        "version": VERSION,
        "saved_at": (now or datetime.now()).strftime("%Y-%m-%dT%H:%M:%S"),
        "settings": values.as_dict(),
        "style_presets": [item.as_dict() for item in custom if not item.builtin],
    }


def parse_document(body: Any, source: Optional[Path] = None) -> Tuple[StyleValues, List[StylePreset]]:
    """把读回来的字典换成 ``(设置值, 自定义预设)``；任何一格读不懂都抛并点名是哪一格。

    ``settings`` 那七格**一格不许缺**：缺 ``with_toc`` 的这份文件到底是"目录关着"还是"没写"，
    没人替它猜。``style_presets`` **可以缺**——一份没有自定义预设的文件，"没有"就是它的真相。
    """
    if not isinstance(body, dict):
        raise SettingsError("这一份不是一个对象（{…} 那种形状），现在是%s。" % _shape(body), source)
    kind = body.get("kind")
    if kind != KIND:
        raise SettingsError(
            "这一份不是 md2docx 的配置文件：kind 那一格写的是 %r（应当是 %r）。" % (kind, KIND), source)
    version = body.get("version")
    if isinstance(version, bool) or not isinstance(version, int):
        raise SettingsError("version 那一格得是一个整数（现在是 %r）。" % (version,), source)
    if version < 1 or version > VERSION:
        raise SettingsError(
            "这一份配置文件的版本是 %d，这一版只写到 %d。"
            "请换新版软件来读，或者在旧版里另存一份。" % (version, VERSION), source)
    extra = [key for key in body if key not in TOP_KEYS]
    if extra:
        raise SettingsError("顶层有读不懂的键：%s（这一版只认 %s）。"
                            % ("、".join(sorted(extra)), "、".join(TOP_KEYS)), source)
    values = _values_of(body.get("settings"), source)
    return values, _presets_of(body.get("style_presets"), source)


def write_config(path: Path, values: StyleValues, custom: Sequence[StylePreset] = (),
                 now: Optional[datetime] = None) -> Path:
    """写一份配置文件。同一次调用给同一个 ``now`` 就**逐字节相同**（差异比对才读得出设置到底变没变）。"""
    where = Path(path)
    text = json.dumps(to_document(values, custom, now), ensure_ascii=False, indent=2,
                      sort_keys=True) + "\n"
    # newline="\n"：Windows 默认把每个 \n 翻成 \r\n，同一份配置导两次就该逐字节相同。
    where.write_text(text, encoding="utf-8", newline="\n")
    return where


def read_config(path: Path) -> Tuple[StyleValues, List[StylePreset]]:
    """读一份配置文件，交回 ``(设置值, 自定义预设)``；读不懂就抛 :class:`SettingsError`。"""
    where = Path(path)
    try:
        raw = where.read_text(encoding="utf-8")
    except FileNotFoundError as exc:
        raise SettingsError("这个文件不在那儿：%s" % where, where) from exc
    except UnicodeDecodeError as exc:
        raise SettingsError("这份配置文件不是 UTF-8 文本：%s（导出时编码选 UTF-8）。" % where,
                            where) from exc
    except OSError as exc:
        raise SettingsError("这份配置文件读不动：%s（%s）" % (where, exc), where) from exc
    try:
        body = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise SettingsError("这份配置文件不是合法的 JSON：%s（第 %d 行第 %d 列：%s）"
                            % (where, exc.lineno, exc.colno, exc.msg), where) from exc
    return parse_document(body, where)


def config_filter() -> str:
    """文件框里那一栏的过滤器（与工程文件那个 ``project_filter`` 同一写法）。"""
    return "md2docx 配置文件 (*.json);;所有文件 (*.*)"


# --- 最近打开的文件 -----------------------------------------------------------

def push_recent(existing: Iterable[str], paths: Iterable[str],
                limit: int = RECENT_LIMIT) -> List[str]:
    """新点中的排最前、同一个文件不占两格、超出上限从尾巴上砍，交回**新的列表**。

    比"同不同一条"用的是**规范化之后的路径**（``normcase`` + ``normpath``，不碰磁盘）：
    ``C:\\temp\\a.json`` 与 ``C:\\temp\\sub\\..\\a.json`` 是同一个文件，大小写在 Windows 上也不算两回事。
    **摆上表的也是规范化后的写法**（保留原本的字母大小写，只把 ``/`` 与 ``..`` 收平）：这一条表是给人
    在菜单里读的，留着 ``…\子目录\..\a.json`` 那种写法，同一个文件在表上今天一个样子明天另一个样子。
    一次只喂"刚点中的那一条"时最自然；一批喂几条就按给的次序摆在前头（谁算"最近"归调用方判，
    这一层不猜）。这里只定"一条表的规则"，喂它工程文件还是 .md 稿子由界面那一步决定。
    """
    limit = int(limit)
    if limit <= 0:
        raise SettingsError("最近打开的文件至少留 1 条（现在是 %d）。" % limit)
    result: List[str] = []
    seen: List[str] = []
    for candidate in list(paths) + list(existing):
        key = _path_key(candidate)
        if not key or key in seen:
            continue
        seen.append(key)
        result.append(os.path.normpath(_clean_text(candidate)))
        if len(result) >= limit:
            break
    return result


# --- 内部 ---------------------------------------------------------------------

def _values_of(body: Any, source: Optional[Path]) -> StyleValues:
    if not isinstance(body, dict):
        raise SettingsError("settings 那一格不是一个对象（应当有 %s）。"
                            % "、".join(SETTINGS_KEYS), source)
    missing = [key for key in SETTINGS_KEYS if key not in body]
    if missing:
        raise SettingsError("settings 里缺了这几格：%s（这一版写出的文件都带齐那七格，"
                            "缺任何一格都没人替它猜是「关着」还是「没写」）。"
                            % "、".join(missing), source)
    extra = [key for key in body if key not in SETTINGS_KEYS]
    if extra:
        raise SettingsError("settings 里有读不懂的键：%s（这一版只认 %s）。"
                            % ("、".join(sorted(extra)), "、".join(SETTINGS_KEYS)), source)
    return StyleValues(
        body_font=_text(body["body_font"], "正文字体（body_font）", source),
        body_size=_size(body["body_size"], "字号（body_size）", source),
        heading_font=_text(body["heading_font"], "标题字体（heading_font）", source),
        with_toc=_flag(body["with_toc"], "生成目录（with_toc）", source),
        page_break=_flag(body["page_break"], "插入分页符（page_break）", source),
        source_label=_flag(body["source_label"], "标注源文件名（source_label）", source),
        cover_page=_flag(body["cover_page"], "生成封面页（cover_page）", source),
    )


def _presets_of(body: Any, source: Optional[Path]) -> List[StylePreset]:
    if body is None:
        return []
    if not isinstance(body, list):
        raise SettingsError("style_presets 那一格得是一个列表，现在是%s。" % _shape(body), source)
    result: List[StylePreset] = []
    for index, item in enumerate(body, 1):
        result = add_custom(result, _preset_of(item, "第 %d 条自定义预设" % index, source))
    return result


def _preset_of(item: Any, where: str, source: Optional[Path] = None) -> StylePreset:
    if not isinstance(item, dict):
        raise SettingsError("%s不是一个对象（应当有 %s），现在是%s。"
                            % (where, "、".join(PRESET_KEYS), _shape(item)), source)
    missing = [key for key in PRESET_KEYS if key not in item]
    if missing:
        raise SettingsError("%s里缺了这几格：%s。" % (where, "、".join(missing)), source)
    extra = [key for key in item if key not in PRESET_KEYS]
    if extra:
        raise SettingsError("%s里有读不懂的键：%s（一条预设只认 %s）。"
                            % (where, "、".join(sorted(extra)), "、".join(PRESET_KEYS)), source)
    name = _text(item["name"], "%s的名字" % where, source)
    return StylePreset(
        name=name,
        body_font=_text(item["body_font"], "「%s」的正文字体" % name, source),
        body_size=_size(item["body_size"], "「%s」的字号" % name, source),
        heading_font=_text(item["heading_font"], "「%s」的标题字体" % name, source),
    )


def _text(value: Any, label: str, source: Optional[Path] = None,
          allow_empty: bool = False) -> str:
    if not isinstance(value, str):
        raise SettingsError("%s那一格得是一段文本，现在是%s。" % (label, _shape(value)), source)
    result = value.strip()
    if not result and not allow_empty:
        raise SettingsError("%s那一格是空的——空着一格与填了值是两回事，"
                            "这一版不替人补一个（补了就等于悄悄改了而没人知道）。" % label, source)
    return result


def _size(value: Any, label: str, source: Optional[Path] = None) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise SettingsError("%s那一格得是一个数（现在是 %r）。" % (label, value), source)
    result = float(value)
    if not SIZE_MIN <= result <= SIZE_MAX:
        raise SettingsError("%s那一格得在 %s ~ %s 之间（现在是 %s）。"
                            % (label, _size_text(SIZE_MIN), _size_text(SIZE_MAX), _size_text(result)),
                            source)
    return result


def _flag(value: Any, label: str, source: Optional[Path] = None) -> bool:
    if not isinstance(value, bool):
        raise SettingsError("%s那一格得是真或假（true / false），现在是 %r。" % (label, value), source)
    return value


def _path_key(value: Any) -> str:
    text = _clean_text(value)
    if not text:
        return ""
    return os.path.normcase(os.path.normpath(text))


def _clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _size_text(value: float) -> str:
    text = ("%.1f" % float(value))
    return text[:-2] if text.endswith(".0") else text


def _shape(value: Any) -> str:
    if isinstance(value, dict):
        return "对象"
    if isinstance(value, list):
        return "列表"
    if isinstance(value, bool):
        return "真或假"
    if isinstance(value, (int, float)):
        return "数"
    if isinstance(value, str):
        return "文本"
    if value is None:
        return "空"
    return type(value).__name__
