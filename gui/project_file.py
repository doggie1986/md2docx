# -*- coding: utf-8 -*-
"""工程文件（.json）：把左边那一列列表存成一份能拷走、能进版本库的抄本。

**这里存的是"清单"，不是"稿子内容"**——哪几个 .md 参加、按什么顺序、勾着其中哪几条、
分组线在哪儿、每一组叫什么名。这四样是"屏幕上看得见而别处查不到"的事实，所以要存；
文件内容、样式、转换结果都不在这一份里（样式与窗口尺寸是 QSettings 那一档，见 :mod:`gui.preferences`）。

**状态不存**：某一条 .md 还在不在盘上、读不读得动，那是盘上的事实，打开时重核一遍才对得上。
存了就会出现"档案里写着在、其实早就被挪走了"这种假账——而且列表照样把那一行摆出来
（带着"已经不在了"那枚图标），所以不存也照样看得见。

**认不出来就拒**，并说清是哪一格：一份读不懂的工程文件被"尽力解析"成半条列表，
比直接报错难查得多（人只看见列表少了几行，不知道少的是哪几行、为什么少）。
同一句理由管着顶层那四格与每一条条目里的键——拼错了的键静默忽略，等于你以为存了、
下一版其实读不到。

路径可以写相对的，**相对的是这份 .json 自己所在的目录**，所以整个文件夹拷到别的机器上照样打得开；
:func:`load` 交回的条目一律已是绝对路径，列表那一侧不需要知道这份文件放在哪儿
（所以本模块不 import :mod:`gui.file_panel`——条目那两三种形状是唯一的接缝）。
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

__all__ = ["KIND", "VERSION", "TOP_KEYS", "Project", "ProjectFileError",
           "save", "load", "project_filter", "default_project_path"]

#: 认门用：这一份是 md2docx 的工程文件，不是别家也叫 `.json` 的什么东西。
KIND = "md2docx.project"

#: 这一版写得出的最高版本。**读的时候只认 1 ~ VERSION**：往后的版本改了条目形状，
#: 硬读会把"少一格"当成"没勾"，那是往列表里塞假数据。
VERSION = 1

TOP_KEYS = ("kind", "version", "saved_at", "entries")
GROUP_KEYS = ("group",)
FILE_KEYS = ("path", "checked")

ENTRY_NOTE = ("每一条要么是分组（只带 group），要么是文件（带 path，checked 可省）。"
              "序号不存：它就是列表里的位置，摆回列表会重新数一遍。")


class ProjectFileError(ValueError):
    """读不懂 / 写不出这份工程文件。句子是给**人**看的中文，带下一步该干什么。"""

    def __init__(self, message: str, source: Optional[Path] = None) -> None:
        super().__init__(message)
        self.source = Path(source) if source is not None else None

    @property
    def user_message(self) -> str:
        return str(self)


@dataclass(frozen=True)
class Project:
    """一份读回来的工程文件。``entries`` 已经是 :meth:`gui.file_panel.FileListWidget.load_entries`
    认识的形状（path 全为绝对路径），主窗口拿到它就一行命令交给列表。"""

    entries: List[dict] = field(default_factory=list)
    version: int = VERSION
    saved_at: str = ""
    source: Path = Path(".")

    def counts(self) -> Dict[str, int]:
        """文件几条、分组几条、一共几行。**只数条目**：已经不存在的 .md 照样算一条——
        那一行会在列表上摆出来，回执念少了就和屏幕上对不上。"""
        groups = sum(1 for entry in self.entries if "group" in entry)
        return {"files": len(self.entries) - groups, "groups": groups, "rows": len(self.entries)}

    def receipt(self) -> str:
        """一行念清"这一份里有什么"，打开之后写进日志区。"""
        count = self.counts()
        tail = "｜%d 组" % count["groups"] if count["groups"] else ""
        return "%d 个文件%s（这一份写于 %s）" % (
            count["files"], tail, self.saved_at or "没登记时间")


def project_filter() -> str:
    """文件框那个"只看工程文件"的过滤串。"""
    return "md2docx 工程文件 (*.json);;所有文件 (*)"


def default_project_path(paths: Sequence[Path]) -> Path:
    """「另存为…」的默认落点：**跟着稿子走，不跟着输出目录走**。

    工程文件记的是"这一批稿子"，所以摆在稿子公共父目录旁边；输出目录里全是成品，
    把清单混进去会让人以为那也是一份成果。取不到公共父目录（没文件、或跨盘）就落到当前目录。
    名字一律叫 ``工程.json``：这里不拿文件夹名当文件名（根目录、``.`` 这种名字要么太长要么没有），
    同一个目录里第二份就靠保存框里改名——Qt 那个框本来就会问"要覆盖吗"。
    """
    taken = [Path(path) for path in paths if str(path).strip()]
    if not taken:
        return Path.cwd() / "工程.json"
    base = _common_parent(taken)
    return base / "工程.json"


def _common_parent(paths: Sequence[Path]) -> Path:
    r"""这些稿子的公共父目录；跨盘、或只公共到盘根就退回当前目录。

    默认落点猜歪了没什么代价（保存框里那一格本来就能改），**抛异常才有代价**——
    它会把「另存为…」这一整条路堵死。盘根也不给：那里多半没有写权限，
    而且把 ``工程.json`` 丢在 ``E:\`` 上比丢在稿子旁边难找。
    Windows 上同一个目录可能写成 ``E:\x`` 也可能写成 ``e:\x``，所以按 ``normcase`` 比、
    交回的是先到那一条的原样写法。
    """
    chains = [[part for part in (path if path.is_absolute() else Path.cwd() / path).parent.parts]
              for path in paths]
    shared: List[str] = []
    for chunk in zip(*chains):
        if len({os.path.normcase(name) for name in chunk}) != 1:
            break
        shared.append(chunk[0])
    if len(shared) <= 1:
        return Path.cwd()
    return Path(*shared)


# ---------------------------------------------------------------- 写


def save(entries: Sequence[dict], target: Path, saved_at: Optional[str] = None) -> Path:
    """把列表交回的条目写成一份 .json，交回真写了哪一个文件。

    写出去的形状与 :func:`load` 认的形状是同一条 :func:`_entry_for_write`——
    能写出去的一定能读回来，否则"存了"这件事要等下次打开才发现。
    """
    where = Path(target)
    where.parent.mkdir(parents=True, exist_ok=True)
    body = {
        "kind": KIND,
        "version": VERSION,
        "saved_at": saved_at or datetime.now().isoformat(timespec="seconds"),
        "entries": [_entry_for_write(entry, where.parent, index)
                    for index, entry in enumerate(entries, start=1)],
    }
    text = json.dumps(body, ensure_ascii=False, indent=2) + "\n"
    # newline="\n"：Windows 默认会把每个 \n 翻成 \r\n，同一份列表存两次就该逐字节相同，
    # 否则"这份工程文件到底改没改"这种事就再也比不出来了。
    where.write_text(text, encoding="utf-8", newline="\n")
    return where


def _entry_for_write(entry: dict, base: Path, index: int) -> Dict[str, Any]:
    if not isinstance(entry, dict):
        raise ProjectFileError("第 %d 条不是一项（这一份是列表交回来的，不该出现这种形状）。" % index)
    if "group" in entry:
        name = str(entry.get("group") or "").strip()
        if not name:
            raise ProjectFileError("第 %d 条是一条分组线，但组名是空的——组名就是那一组的输出文件名。" % index)
        if "path" in entry:
            raise ProjectFileError("第 %d 条同时写了 group 和 path：一条要么是分组要么是文件。" % index)
        return {"group": name}
    raw = str(entry.get("path") or "").strip()
    if not raw:
        raise ProjectFileError("第 %d 条既没有 group 也没有 path：%s" % (index, ENTRY_NOTE))
    out: Dict[str, Any] = {"path": _stored_path(Path(raw), base)}
    if "checked" in entry:
        value = entry.get("checked")
        if not isinstance(value, bool):
            raise ProjectFileError("第 %d 条的 checked 得是真或假（现在是 %r）。" % (index, value))
        out["checked"] = value
    extra = sorted(key for key in entry if key not in FILE_KEYS)
    if extra:
        raise ProjectFileError("第 %d 条里有读不懂的键：%s。" % (index, "、".join(extra)))
    return out


def _stored_path(path: Path, base: Path) -> str:
    """能写成相对就相对（整目录拷走照样打得开），跨盘 / 不在这一目录下就老老实实写绝对路径。"""
    text = path if path.is_absolute() else (base / path)
    try:
        relative = Path(text).resolve().relative_to(Path(base).resolve())
    except (ValueError, OSError):
        # ValueError = 不在这个父目录下（跨盘也算）；OSError = 这一台上根本比不了（路径坏了）。
        return Path(text).as_posix()
    return relative.as_posix()


# ---------------------------------------------------------------- 读


def load(source: Path) -> Project:
    """读一份工程文件，交回 :class:`Project`；读不懂就抛 :class:`ProjectFileError`。

    一层一层判、**一句一句短路**：连顶层都不认识的档案，"第 3 条的 checked 不是真假"是假问题。
    """
    where = Path(source)
    raw = _read_text(where)
    document = _parse(raw, where)
    if not isinstance(document, dict):
        raise ProjectFileError(
            "这份工程文件的顶层得是一个对象（{…}），现在是一个 %s。" % _shape(document), where)

    unknown = sorted(key for key in document if key not in TOP_KEYS)
    if unknown:
        raise ProjectFileError(
            "这份工程文件里有读不懂的键：%s。这一版只认 %s。"
            % ("、".join(unknown), "、".join(TOP_KEYS)), where)

    kind = document.get("kind")
    if str(kind or "").strip() != KIND:
        raise ProjectFileError(
            "这一份不是 md2docx 的工程文件：kind 写的是 %r，要的是 %r。"
            % (kind, KIND), where)

    version = _version_of(document.get("version"), where)
    entries = document.get("entries")
    if entries is None:
        raise ProjectFileError("这份工程文件没有 entries 那一格（列表就是存在这一格里的）。", where)
    if not isinstance(entries, list):
        raise ProjectFileError("entries 那一格得是一个列表，现在是一个 %s。" % _shape(entries), where)

    base = where.parent
    cleaned = [_entry_for_read(entry, base, where, index)
               for index, entry in enumerate(entries, start=1)]
    saved_at = document.get("saved_at")
    return Project(entries=cleaned, version=version,
                   saved_at="" if saved_at is None else str(saved_at), source=where)


def _read_text(where: Path) -> str:
    try:
        # utf-8-sig：记事本另存时可能带一道 BOM；带着 BOM 的 JSON 会让第一格键名变成 "\ufeffkind"，
        # 于是报出"读不懂的键"这种让人摸不着头脑的话。
        return where.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        # 那句下一步引的是**菜单上真写着的词**（「工程另存为…」）——让人去找一条不存在菜单项的话，
        # 比不指下一步更糟。
        raise ProjectFileError("这个文件不在那儿：%s（先点「文件 → 工程另存为…」存一份）。" % where,
                               where) from None
    except UnicodeDecodeError:
        raise ProjectFileError("这份工程文件不是 UTF-8 文本：%s（另存一份时编码选 UTF-8）。" % where,
                               where) from None
    except OSError as exc:
        raise ProjectFileError("这份工程文件读不动：%s（%s）"
                               % (where, exc.strerror or exc.__class__.__name__), where) from None


def _parse(raw: str, where: Path) -> Any:
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ProjectFileError("这份工程文件不是合法的 JSON：%s（第 %d 行第 %d 列：%s）"
                               % (where, exc.lineno, exc.colno, exc.msg), where) from None


def _version_of(value: Any, where: Path) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ProjectFileError("version 那一格得是一个整数（现在是 %r）。" % (value,), where)
    if value < 1 or value > VERSION:
        raise ProjectFileError(
            "这一版读不懂第 %d 版的工程文件（这一版最高到 %d）——换新版软件打开，"
            "或者在旧版里另存一份。" % (value, VERSION), where)
    return value


def _entry_for_read(entry: Any, base: Path, source: Path, index: int) -> Dict[str, Any]:
    where = "第 %d 条" % index
    if not isinstance(entry, dict):
        raise ProjectFileError("%s不是一项（{%s} 那种形状），现在是一个 %s。%s"
                               % (where, "…", _shape(entry), ENTRY_NOTE), source)
    keys = set(entry)
    if keys == set(GROUP_KEYS):
        name = str(entry.get("group") or "").strip()
        if not name:
            raise ProjectFileError("%s是一条分组线，但 group 那一格是空的——组名就是那一组的输出文件名。"
                                   % where, source)
        return {"group": name}
    if "path" not in keys:
        raise ProjectFileError("%s既不是分组也不是文件（要有 group，或者要有 path）：%s"
                               % (where, ENTRY_NOTE), source)
    extra = sorted(keys - set(FILE_KEYS))
    if extra:
        raise ProjectFileError("%s里有读不懂的键：%s（文件那一条只认 %s）。"
                               % (where, "、".join(extra), "、".join(FILE_KEYS)), source)
    raw = str(entry.get("path") or "").strip()
    if not raw:
        raise ProjectFileError("%s的 path 那一格是空的。" % where, source)
    out: Dict[str, Any] = {"path": _resolved(raw, base)}
    if "checked" in keys:
        value = entry.get("checked")
        if not isinstance(value, bool):
            raise ProjectFileError("%s的 checked 得是真或假（true / false），现在是 %r。"
                                   % (where, value), source)
        out["checked"] = value
    return out


def _resolved(raw: str, base: Path) -> str:
    """相对的按**这份 .json 所在的目录**接全；绝对的照原样。交回的一律是绝对路径文本。"""
    path = Path(raw)
    if not path.is_absolute():
        path = base / path
    return str(path)


def _shape(value: Any) -> str:
    """给人看的形状名：说"是一个 list"比说 "array" 好读，也不把 Python 类型名甩给人。"""
    if value is None:
        return "空值 null"
    if isinstance(value, bool):
        return "真/假"
    if isinstance(value, (int, float)):
        return "一个数"
    if isinstance(value, str):
        return "一段文本"
    if isinstance(value, list):
        return "列表"
    if isinstance(value, dict):
        return "对象"
    return type(value).__name__
