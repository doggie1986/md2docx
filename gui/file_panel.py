# -*- coding: utf-8 -*-
"""左侧那一栏：待转换的文件列表、分组、以及工程文件的进出口。

先说清这一栏的几条口径，免得日后当成 bug 去"修"：

* **复选框 = 这条参不参与这一批**，不是"选中"。选中（高亮）只决定预览区演哪一个文件、
  以及"移除 / 上移 / 下移 / 批量勾选"动的是哪几行——两件事分开放，勾成一排时手一抖不会把选中也挪走。
* **多选与列表内拖排并存**（以前是单选，为的是拖排时顺序对得上）。多选之后"上移/下移"
  动的是**整块**：选中三行就三行一起往上走一格，而不是三行各自挪一格——后者会把三行打散，
  中间夹进别人的行，而那一列序号就是合并顺序，打散了就是改错了顺序。
* **分组是一条"组头行"，不是 ``insertSeparator()``**：分隔符不带名字，而每一组要单独出一份
  .docx，那份文件的名字就是组名。组头不算文件、不算序号、选不中也拖不动。
* **组的归属看位置**：一个文件属于**它上方最近的那条组头**；排在第一条组头之前的文件是
  "未分组"。所以拖排不只是好看——把一行拖到另一条组头之上，它就是换了组、换了输出文件。
* **序号是文本前缀**（``1. 名字``）而不是自画的行号栏：给行画 delegate 要把绘制矩形缩进去，
  而复选框是 Qt 原生画的，缩进去之后命中区就跟视觉错位——点勾那儿就勾不上了。
* **外部拖进来的非 Markdown 文件直接跳过并念一句**，不默默收下——收了会在转换那一步
  报"读不动"，而那时人早忘了是自己拖错了东西。
"""

from __future__ import annotations

import contextlib
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, Sequence

from PySide6.QtCore import QItemSelectionModel, QSize, Qt, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (QAbstractItemView, QInputDialog, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QMenu, QVBoxLayout, QWidget)

from gui import theme

__all__ = ["MARKDOWN_SUFFIXES", "MAX_FILES_PER_DROP", "UNGROUPED", "ROW_FILE", "ROW_GROUP",
           "is_markdown", "looks_markdown", "FilePanel", "FileListWidget"]

MARKDOWN_SUFFIXES = frozenset({".md", ".markdown", ".mkd", ".mdown"})

#: 一次收档的上限。没有这一档，拖进来一整个文档目录会把界面钉住。
MAX_FILES_PER_DROP = 500

#: 排在第一条组头之前的那些文件用的组名（它们照样出得了文件，只是不能没名字）。
UNGROUPED = "未分组"

ROW_FILE = "file"
ROW_GROUP = "group"

_PATH_ROLE = Qt.ItemDataRole.UserRole
_KIND_ROLE = Qt.ItemDataRole.UserRole + 1
_GROUP_ROLE = Qt.ItemDataRole.UserRole + 2
_STATUS_ROLE = Qt.ItemDataRole.UserRole + 3


def is_markdown(path: Path) -> bool:
    """拖放那一侧的口径：后缀对**且**文件真在那儿。"""
    return path.is_file() and path.suffix.lower() in MARKDOWN_SUFFIXES


def looks_markdown(path: Path) -> bool:
    """只看后缀，不看盘。工程文件恢复时要认得"一条已经不存在的 .md"（那一行该带着
    "已经不在了"的图标摆出来，而不是被当成不是 Markdown 悄悄丢掉），
    而 :func:`is_markdown` 要求文件真的在那儿。"""
    return Path(path).suffix.lower() in MARKDOWN_SUFFIXES


def markdown_in(folder: Path) -> Iterator[Path]:
    """目录里的 Markdown，按路径排序（``rglob`` 的默认次序不稳定，列表会一次一个样）。"""
    for candidate in sorted(folder.rglob("*"), key=lambda item: str(item).lower()):
        if is_markdown(candidate):
            yield candidate


class FileListWidget(QListWidget):
    """文件列表本体：外部拖入、列表内拖排、多选批量、分组、状态图标、右键菜单都长在这个类上。"""

    #: 给用户看的一句话（跳过了什么、收了多少、改了哪几行），主窗口拿去写日志区。
    noticed = Signal(str)
    #: 右键"单独转换这一条"，交回那一条的路径。
    convert_single = Signal(object)
    #: 选中行换了（交回 Path；组头与空列表都**不发**这一条，预览区不该被一条组头清空）
    selected_changed = Signal(object)
    #: 条数或勾选状态变了（有几个文件，勾了其中几个）——**文件数不含组头**。
    counts_changed = Signal(int, int)
    #: 空白处右键那两条：主窗口认识"要开文件框 / 要开目录框 / 要清空"，列表不认识文件对话框。
    add_files_requested = Signal()
    add_folder_requested = Signal()
    clear_requested = Signal()
    #: 列表内容变了（增删、改勾、挪序、分组都算）。主窗口拿它标"工程文件还没存"那颗星号。
    changed = Signal()

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setObjectName("fileList")
        self._theme = theme.DEFAULT_THEME
        self._bulk = 0            # 批量改动的嵌套计数，见 :meth:`_batch`
        self._renumbering = False
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        # 注意是 DragDropMode 不是 DropAction：InternalMove 属于前者，
        # 写成 DropAction 会在运行时抛 AttributeError（构造列表那一刻就炸，跑不起来了）。
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setDropIndicatorShown(True)
        self.setAcceptDrops(True)
        # 刻意**不**开隔行底色：Qt 在这一格用的是系统调色板的 AlternateBase（这台机器上
        # 是一眼看得见的绿色 #6cba74，QSS 里没写它、主题管不到），而更要紧的是它会盖住
        # 组头自己那格底色——"这一组单独出一份"全靠那条带子，落在奇数行就没了。
        self.setUniformItemSizes(True)
        self.setIconSize(QSize(theme.STATUS_ICON_EDGE, theme.STATUS_ICON_EDGE))
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_menu)
        self.currentItemChanged.connect(self._echo_selection)
        self.itemChanged.connect(self._on_item_changed)

    # ---- 增删 -------------------------------------------------------------

    def add_paths(self, paths: Iterable[Path]) -> int:
        """收下这些路径（目录会翻进去找 Markdown），交回真加进去的条数。"""
        had_current = self.currentItem() is not None
        added = 0
        skipped_kind = 0
        skipped_dupe = 0
        overflow = 0
        # announce=False：一行没收（整目录没有 .md、交回来的全是重复）不该把工程标成
        # "有没保存的改动"——星号说的是"列表和刚存进盘的那份不一样"，而它现在还是原样。
        with self._batch(announce=False):
            for given in paths:
                path = Path(given)
                if path.is_dir():
                    found = list(markdown_in(path))
                    if not found:
                        self.noticed.emit("目录 %s 里没有 Markdown 文件。" % path)
                    for found_path in found:
                        if self.count() >= MAX_FILES_PER_DROP:
                            overflow += 1
                            continue
                        if self._append(found_path):
                            added += 1
                    continue
                if not is_markdown(path):
                    skipped_kind += 1
                    continue
                if self.count() >= MAX_FILES_PER_DROP:
                    overflow += 1
                    continue
                if self._append(path):
                    added += 1
                else:
                    skipped_dupe += 1
        if skipped_kind:
            self.noticed.emit("跳过了 %d 个不是 Markdown 的文件（只收 %s）。"
                              % (skipped_kind, "、".join(sorted(MARKDOWN_SUFFIXES))))
        if skipped_dupe:
            self.noticed.emit("跳过了 %d 个列表里已经有的文件。" % skipped_dupe)
        if overflow:
            self.noticed.emit("到了 %d 条上限，后面 %d 个没收——请分批拖，或把范围收窄一点。"
                              % (MAX_FILES_PER_DROP, overflow))
        if added:
            self.changed.emit()
            self.noticed.emit("加入了 %d 个文件。" % added)
            self.refresh_status(announce=False)
        if not had_current and self.file_count():
            # 第一批文件落进空列表时顺手选上第一条：不选的话中间那块还停在欢迎页，
            # 人刚把文件拖进来却什么都看不见，会以为拖放没生效。
            self.select_first()
        return added

    def _append(self, path: Path, checked: bool = True) -> bool:
        """一条一个复选框，默认勾上。同一路径不重复收（比的是解析之后的绝对路径）。"""
        resolved = _canonical(path)
        if self._find(resolved) is not None:
            return False
        item = QListWidgetItem(resolved.name, self)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsDragEnabled)
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        item.setData(_PATH_ROLE, str(resolved))
        item.setData(_KIND_ROLE, ROW_FILE)
        self._set_status(item, theme.STATUS_UNKNOWN)
        return True

    def remove_selected(self) -> List[Path]:
        """移除选中的那几行（组头选不中，所以这里全是文件行）。没选中交回**空表**，
        由主窗口念"先选中一行"——交回 ``None`` 与交回一条空列表在"移除 0 行"上是两回事。"""
        rows = self.selected_rows()
        if not rows:
            return []
        paths = [self.path_of(self.item(row)) for row in rows]
        with self._batch():
            for row in reversed(rows):
                self.takeItem(row)
        self.noticed.emit("移除了 %d 行（盘上的文件一个字都没动）。" % len(paths))
        return paths

    def remove_row(self, row: int) -> Optional[Path]:
        item = self.item(row)
        if item is None:
            return None
        if self.is_group(item):
            return self.remove_group(row)
        path = self.path_of(item)
        with self._batch():
            self.takeItem(row)
        self.noticed.emit("移除了 %s。" % path.name)
        return path

    def clear_files(self) -> int:
        """清空列表（组头也一并去掉），交回去掉的文件条数——主窗口的日志念的就是这个数。"""
        removed = self.file_count()
        was = self.count()
        with self._batch(announce=False):
            self.clear()
        # 判"列表变没变"看行数而不是文件条数：列表里只剩一条分组线时文件条数是 0，行数却确实少了。
        if self.count() != was:
            self.changed.emit()
        return removed

    def move_selected(self, upwards: bool) -> int:
        """把选中的行往上/往下挪一格，**连续的那几行当一个整块走**。
        交回真挪了几行：0 表示"要么没选中，要么这一行已经到顶/到底，要么再挪就越过组头了"。"""
        runs = self._runs(self.selected_rows())
        if not runs:
            return 0
        moved = 0
        with self._batch(announce=False):
            for items in runs:
                moved += self._move_run(items, upwards)
        if moved:
            self.changed.emit()
            self.noticed.emit("选中的 %d 行%s挪了一格（合并与分组的顺序就是这一列的次序）。"
                              % (moved, "上" if upwards else "下"))
        return moved

    def set_all_checked(self, checked: bool) -> int:
        word = "参与" if checked else "不参与"
        wanted = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        changed = 0
        with self._batch(announce=False):
            for item in self._file_items():
                if item.checkState() != wanted:
                    changed += 1
                    item.setCheckState(wanted)
        # 一格都没改动时（本来就已经全是这一种勾选）交回 0、念"把 0 个文件改成…"、星号不亮。
        self.noticed.emit("把 %d 个文件改成「%s这一批」。" % (changed, word))
        if changed:
            self.changed.emit()
        return changed

    def set_checked_for_rows(self, rows: Sequence[int], checked: bool) -> int:
        """批量勾选的唯一实现：给一个行号表就改那几行（选中行、某一组的成员行都走这里）。"""
        word = "参与" if checked else "不参与"
        changed = 0
        with self._batch(announce=False):
            for row in rows:
                item = self.item(int(row))
                if item is None or self.is_group(item):
                    continue
                wanted = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
                if item.checkState() != wanted:
                    changed += 1
                item.setCheckState(wanted)
        self.noticed.emit("把 %d 个文件改成「%s这一批」。" % (changed, word))
        if changed:
            self.changed.emit()
        return changed

    # ---- 分组 -------------------------------------------------------------

    def groups(self) -> List[str]:
        """组名，按列表里的先后次序（那个次序就是出文件的次序）。"""
        return [self.group_name_of(item) for item in self._group_items()]

    def has_groups(self) -> bool:
        return any(self.is_group(self.item(row)) for row in range(self.count()))

    def group_rows(self) -> List[int]:
        return [row for row in range(self.count()) if self.is_group(self.item(row))]

    def add_group(self, name: str) -> bool:
        """新建一组。选中了几行就插在**那几行下面**（"把这几行分一组"是这一条最常见的来路），
        否则排在列表末尾。"""
        name = (name or "").strip()
        if not name:
            self.noticed.emit("组名不能空着——每一组单独出一份 .docx，那个文件名就是组名。")
            return False
        if name in self.groups():
            self.noticed.emit("已经有「%s」这一组了，换个名字（两组出两个同名文件会互相顶掉）。" % name)
            return False
        rows = self.selected_rows()
        at = max(rows) + 1 if rows else self.count()
        with self._batch():
            self._insert_group(at, name)
        self.noticed.emit("新建了「%s」这一组，排在它下面的文件都归它（组名就是那份 .docx 的文件名）。"
                          % name)
        return True

    def rename_group(self, row: int, name: str) -> bool:
        name = (name or "").strip()
        item = self.item(row)
        if item is None or not self.is_group(item):
            return False
        old = self.group_name_of(item)
        if not name:
            self.noticed.emit("组名不能空着——「%s」这一组照旧。" % old)
            return False
        if name in [value for value in self.groups() if value != old]:
            self.noticed.emit("已经有「%s」这一组了，「%s」没改。" % (name, old))
            return False
        if name == old:
            # 敲了同一个名字：列表一个字没变，所以既不亮星号也不能说"改名叫 X 了"。
            self.noticed.emit("「%s」这一组名字没变。" % old)
            return True
        with self._batch():
            item.setData(_GROUP_ROLE, name)
        self.noticed.emit("「%s」这一组改名叫「%s」了（下一份输出的文件名跟着改）。" % (old, name))
        return True

    def remove_group(self, row: int) -> Optional[str]:
        """只去掉那条分组线，底下的文件一行都不动——它们并进上面那一组（或变成未分组）。"""
        item = self.item(row)
        if item is None or not self.is_group(item):
            return None
        name = self.group_name_of(item)
        members = len(self._members(row))
        with self._batch():
            self.takeItem(row)
        self.noticed.emit("去掉了「%s」的分组线，底下 %d 个文件本身没动（它们并入上一组）。"
                          % (name, members))
        return name

    def group_members(self, row: int) -> List[Path]:
        """那一行（组头）底下的文件，一直数到下一条组头或列表末尾。"""
        return [self.path_of(item) for item in self._members(row)]

    def empty_groups(self) -> List[str]:
        """一个文件都没勾上的组名。分组输出时这些组**出不了文件**，必须念出来——
        只给一份"出了 N 份"的清单，看的人会以为每一组都有东西。"""
        return [self.group_name_of(self.item(row)) for row in self.group_rows()
                if not self._checked_members(row)]

    def group_plans(self) -> List[tuple]:
        """按组切好的 ``(组名, [路径, …])``，只含勾上的文件；空组不摆（由 :meth:`empty_groups` 念）。"""
        plans: List[tuple] = []
        name = UNGROUPED
        bucket: List[Path] = []
        for row in range(self.count()):
            item = self.item(row)
            if self.is_group(item):
                if bucket:
                    plans.append((name, bucket))
                name = self.group_name_of(item)
                bucket = []
            elif item.checkState() == Qt.CheckState.Checked:
                bucket.append(self.path_of(item))
        if bucket:
            plans.append((name, bucket))
        return plans

    def _insert_group(self, row: int, name: str) -> QListWidgetItem:
        item = QListWidgetItem(name)
        # 只给 ItemIsEnabled：选不中（所以批量操作永远碰不到组头）、拖不动（所以不会把分组线
        # 拖到文件堆里）；留 ItemIsDropEnabled 是让文件能往它上面那条指示线落。
        item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsDropEnabled)
        item.setData(_KIND_ROLE, ROW_GROUP)
        item.setData(_GROUP_ROLE, name)
        item.setToolTip("这一组单独出一份 .docx，文件名就是组名。归组看位置：排在它下面的文件都属于它。")
        font = item.font()
        font.setBold(True)
        item.setFont(font)
        background, foreground = theme.group_colors(self._theme)
        item.setBackground(background)
        item.setForeground(foreground)
        self.insertItem(row, item)
        return item

    def _members(self, header_row: int) -> List[QListWidgetItem]:
        found: List[QListWidgetItem] = []
        for row in range(header_row + 1, self.count()):
            item = self.item(row)
            if self.is_group(item):
                break
            found.append(item)
        return found

    def _checked_members(self, header_row: int) -> List[QListWidgetItem]:
        return [item for item in self._members(header_row)
                if item.checkState() == Qt.CheckState.Checked]

    def _group_member_rows(self, header_row: int) -> List[int]:
        """那一行的组头**管到哪儿**：往下数到下一条组头为止的文件行号。
        直接 ``range(header_row + 1, count())`` 是错的——那会把下一组的文件一起改勾，
        而"这一组全部不参与"点下去动了别组，是最难当场发现的一种改法。"""
        return [self.row(item) for item in self._members(header_row)]

    # ---- 选中 -------------------------------------------------------------

    def selected_rows(self) -> List[int]:
        """选中的行号，按列表次序。组头选不中，所以这里永远是文件行。

        行号要从**项**反查：``QListWidget.row()`` 在 PySide6 里只认 ``QListWidgetItem``，
        直接把 ``QModelIndex`` 递进去是 ``TypeError``（C++ 那个重载没绑过来）。
        """
        rows = set()
        for index in self.selectedIndexes():
            item = self.itemFromIndex(index)
            if item is not None:
                rows.add(self.row(item))
        return sorted(rows)

    def selected_paths(self) -> List[Path]:
        return [self.path_of(self.item(row)) for row in self.selected_rows()]

    def selected_count(self) -> int:
        return len(self.selected_rows())

    # ---- 状态图标 ----------------------------------------------------------

    def refresh_status(self, announce: bool = True) -> Dict[str, int]:
        """逐行核一遍"这个文件还在不在、读不读得动"，把结果画成行首那枚图标。

        读一个字节而不是整份文件：几百行 × 几 MB 会真把界面钉住，而"读不动"在这一栏
        要说的是"打开都打不开"（没了、没权限、被别的程序锁住），不是"内容有问题"——
        内容那一层的检查归「校验」。
        """
        counts = {theme.STATUS_READY: 0, theme.STATUS_MISSING: 0, theme.STATUS_UNREADABLE: 0}
        # 逐行 _set_status 会一格派一次 itemChanged，那一回调里要重排序号、发计数、标脏：
        # 几百行就是几百次全表重排，还会把"工程文件还没存"那颗星号闪几百遍。压成一批。
        with self._batch(announce=False):
            for item in self._file_items():
                path = self.path_of(item)
                if not path.is_file():
                    kind = theme.STATUS_MISSING
                elif not _readable(path):
                    kind = theme.STATUS_UNREADABLE
                else:
                    kind = theme.STATUS_READY
                counts[kind] += 1
                self._set_status(item, kind)
        if announce:
            line = "核过 %d 个文件：%d 个能读" % (sum(counts.values()), counts[theme.STATUS_READY])
            if counts[theme.STATUS_MISSING]:
                line += "，%d 个已经不在了" % counts[theme.STATUS_MISSING]
            if counts[theme.STATUS_UNREADABLE]:
                line += "，%d 个读不动（没权限或被占着）" % counts[theme.STATUS_UNREADABLE]
            self.noticed.emit(line + "。")
        return counts

    def status_of(self, item: Optional[QListWidgetItem]) -> str:
        if item is None:
            return ""
        return str(item.data(_STATUS_ROLE) or theme.STATUS_UNKNOWN)

    def status_at(self, row: int) -> str:
        return self.status_of(self.item(row))

    def _set_status(self, item: QListWidgetItem, kind: str) -> None:
        item.setData(_STATUS_ROLE, kind)
        item.setIcon(theme.status_icon(kind, self._theme))
        tip = "%s\n状态：%s" % (self.path_of(item), theme.status_label(kind))
        if kind == theme.STATUS_UNKNOWN:
            tip += "（还没核过——右键「重新核对文件状态」或按 Ctrl+R）"
        item.setToolTip(tip)

    # ---- 读 ---------------------------------------------------------------

    def path_of(self, item: QListWidgetItem) -> Path:
        return Path(str(item.data(_PATH_ROLE) or ""))

    def group_name_of(self, item: QListWidgetItem) -> str:
        return str(item.data(_GROUP_ROLE) or "")

    def is_group(self, item: Optional[QListWidgetItem]) -> bool:
        return item is not None and item.data(_KIND_ROLE) == ROW_GROUP

    def row_kind(self, row: int) -> str:
        item = self.item(row)
        return ROW_GROUP if self.is_group(item) else ROW_FILE

    def file_count(self) -> int:
        return sum(1 for row in range(self.count()) if self.row_kind(row) == ROW_FILE)

    def row_count(self) -> int:
        return self.count()

    def checked_paths(self) -> list:
        return [self.path_of(item) for item in self._file_items()
                if item.checkState() == Qt.CheckState.Checked]

    def selected_path(self) -> Optional[Path]:
        """当前那一行的文件；选中组头时交回 ``None``（预览区该停在刚才那个文件上，
        而不是被一条组头清空——组头没有内容可演）。"""
        item = self.currentItem()
        return None if item is None or self.is_group(item) else self.path_of(item)

    def checked_count(self) -> int:
        return sum(1 for item in self._file_items()
                   if item.checkState() == Qt.CheckState.Checked)

    def select_first(self) -> Optional[Path]:
        """选上第一个**文件**行（第一条可能是组头，而组头不该决定预览区演什么）。"""
        for row in range(self.count()):
            if self.row_kind(row) == ROW_FILE:
                self.setCurrentRow(row)
                return self.selected_path()
        return None

    def _file_items(self) -> List[QListWidgetItem]:
        return [self.item(row) for row in range(self.count())
                if self.row_kind(row) == ROW_FILE]

    def _group_items(self) -> List[QListWidgetItem]:
        return [self.item(row) for row in range(self.count()) if self.is_group(self.item(row))]

    # ---- 工程文件（那一列序号与勾选就是它的全部内容） ----------------------

    def entries(self) -> List[dict]:
        """按列表次序交回一条条档案条目：组头给 ``{"group": 名}``，文件给 ``{"path": …, "checked": …}``。

        顺序与勾选状态**都**在里面——这两个就是这一栏唯一"屏幕上看得见、别处查不到"的事实。
        状态不存：那是盘上的事实，下次打开重核一遍才对得上。
        """
        out: List[dict] = []
        for row in range(self.count()):
            item = self.item(row)
            if self.is_group(item):
                out.append({"group": self.group_name_of(item)})
            else:
                out.append({"path": str(self.path_of(item)),
                            "checked": item.checkState() == Qt.CheckState.Checked})
        return out

    def load_entries(self, entries: Sequence[dict], replace: bool = True) -> dict:
        """把档案条目摆回列表。已经不存在的 .md **照样摆出来**（带着"已经不在了"那枚图标），
        因为"你的工程里那条文件被挪走了"是必须看得见的事实，不是可以顺手丢掉的一行。"""
        files = checked = groups = 0
        missing: List[str] = []
        skipped: List[str] = []
        overflow = 0
        with self._batch(announce=False):
            if replace:
                self.clear()
            for entry in entries:
                if self.count() >= MAX_FILES_PER_DROP:
                    overflow += 1
                    continue
                name = entry.get("group")
                if name is not None:
                    text = str(name).strip()
                    if not text or text in self.groups():
                        skipped.append("分组「%s」" % text)
                        continue
                    self._insert_group(self.count(), text)
                    groups += 1
                    continue
                raw = str(entry.get("path") or "")
                path = Path(raw)
                if not looks_markdown(path):
                    skipped.append("%s（不是 Markdown 后缀）" % (path.name or raw))
                    continue
                resolved = _canonical(path)
                if self._find(resolved) is not None:
                    skipped.append("%s（列表里已经有了）" % resolved.name)
                    continue
                if not resolved.is_file():
                    missing.append(resolved.name)
                self._append(resolved, checked=bool(entry.get("checked", True)))
                files += 1
                if bool(entry.get("checked", True)):
                    checked += 1
        if overflow:
            skipped.append("超过 %d 条上限的那 %d 条" % (MAX_FILES_PER_DROP, overflow))
        self.refresh_status(announce=False)
        return {"files": files, "checked": checked, "groups": groups,
                "missing": missing, "skipped": skipped}

    # ---- 拖放 -------------------------------------------------------------

    def dragEnterEvent(self, event) -> None:  # noqa: D102
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:  # noqa: D102
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event) -> None:  # noqa: D102
        if not event.mimeData().hasUrls():
            super().dropEvent(event)  # 列表内拖排：整套交给 InternalMove
            # 序号跟着位置走，所以挪完必须重排一遍号；Qt 那边是 remove+insert，
            # 没有"移动完"这个信号可以接，就在这一行后面自己补一次。
            with self._batch():
                pass
            event.acceptProposedAction()
            return
        paths = [Path(url.toLocalFile()) for url in event.mimeData().urls() if url.isLocalFile()]
        if paths:
            self.add_paths(paths)
        event.acceptProposedAction()

    # ---- 右键菜单 ---------------------------------------------------------

    def _show_menu(self, position) -> None:
        self._build_menu(position).exec(self.mapToGlobal(position))

    def _build_menu(self, position) -> QMenu:
        """建右键那张菜单。摆哪几项与点了干什么都在这一个方法里，``exec`` 只在上面那一行——
        分开是因为菜单一 exec 就挡住事件循环，脚本与测试核不了里面的条目。"""
        item = self.itemAt(position)
        menu = QMenu(self)
        if item is not None and self.is_group(item):
            self._group_menu(menu, position)
            return menu
        if item is not None:
            self._row_menu(menu, position)
        else:
            self._blank_menu(menu)
        return menu

    def _row_menu(self, menu: QMenu, position) -> None:
        item = self.itemAt(position)
        row = self.row(item)
        path = self.path_of(item)
        menu.addAction("单独转换这一条", lambda: self.convert_single.emit(path))
        menu.addSeparator()
        rows = self.selected_rows()
        if len(rows) > 1 and row in rows:
            count = len(rows)
            menu.addAction("选中的 %d 行都参与转换" % count,
                           lambda target=rows: self.set_checked_for_rows(target, True))
            menu.addAction("选中的 %d 行都不参与" % count,
                           lambda target=rows: self.set_checked_for_rows(target, False))
            menu.addAction("移除选中的 %d 行" % count, self.remove_selected)
        else:
            checked = item.checkState() == Qt.CheckState.Checked
            menu.addAction("这一条不参与转换" if checked else "这一条参与转换",
                           lambda target=row, now=checked: self._set_checked(target, not now))
            menu.addAction("从列表移除", lambda target=row: self.remove_row(target))
        menu.addSeparator()
        status = theme.status_label(self.status_of(item))
        menu.addAction("文件状态：%s" % status, lambda: self.refresh_status())
        menu.addAction("打开所在目录", lambda target=path: self.open_folder(target))

    def _group_menu(self, menu: QMenu, position) -> None:
        item = self.itemAt(position)
        row = self.row(item)
        name = self.group_name_of(item)
        members = self._members(row)
        menu.addAction("改「%s」的名字…" % name, lambda target=row: self.prompt_group(target))
        menu.addSeparator()
        menu.addAction("「%s」全部参与（%d 个）" % (name, len(members)),
                       lambda target=row: self.set_checked_for_rows(self._group_member_rows(target), True))
        menu.addAction("「%s」全部不参与（%d 个）" % (name, len(members)),
                       lambda target=row: self.set_checked_for_rows(self._group_member_rows(target), False))
        menu.addSeparator()
        menu.addAction("去掉这一组的分组线", lambda target=row: self.remove_group(target))

    def _blank_menu(self, menu: QMenu) -> None:
        menu.addAction("加入 Markdown 文件…", self.add_files_requested.emit)
        menu.addAction("加入整个文件夹…", self.add_folder_requested.emit)
        menu.addSeparator()
        menu.addAction("全部参与转换", lambda: self.set_all_checked(True))
        menu.addAction("全部不参与", lambda: self.set_all_checked(False))
        menu.addSeparator()
        menu.addAction("新建分组…", self.prompt_group)
        menu.addAction("重新核对文件状态", lambda: self.refresh_status())
        menu.addAction("清空列表", self.clear_requested.emit)

    def prompt_group(self, row: Optional[int] = None) -> None:
        """新建/改名共用的那一步：问一个名字。``row`` 给了就是改那一组的名。"""
        if row is None:
            self.add_group(self.ask_text("新建分组", "组名（也就是这一组那份 .docx 的文件名）："))
            return
        item = self.item(row)
        if item is None:
            return
        self.rename_group(row, self.ask_text("改这一组的名字", "新的组名：",
                                             self.group_name_of(item)))

    def ask_text(self, title: str, label: str, value: str = "") -> str:
        """弹一张"就敲一个名字"的框；取消交回空串（空串在 :meth:`add_group` 那一关会被拒）。

        单独列成一个方法是给测试留的口：菜单项的执行体挂在 lambda 里，而这一框一 ``exec``
        就挡住事件循环，脚本核不了"敲了个名"之后列表变成什么样。
        """
        text, ok = QInputDialog.getText(self, title, label, QLineEdit.EchoMode.Normal, value)
        return text if ok else ""

    def open_folder(self, path: Path) -> None:
        """在资源管理器里打开那个目录；打不开要念出来，不能一声不响。"""
        folder = _canonical(path).parent
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder))):
            self.noticed.emit("打不开这个目录：%s（这台机器上没有能打开它的程序？）" % folder)

    # ---- 主题 -------------------------------------------------------------

    def apply_theme(self, name: str) -> None:
        """换主题时组头的底色与状态图标都得重画一遍：它们走的是**行自己**的颜色，
        QSS 管不到"某一行"。"""
        self._theme = name if name in theme.theme_names() else theme.DEFAULT_THEME
        background, foreground = theme.group_colors(self._theme)
        for row in range(self.count()):
            item = self.item(row)
            if self.is_group(item):
                item.setBackground(background)
                item.setForeground(foreground)
            else:
                self._set_status(item, self.status_of(item))

    # ---- 内部 -------------------------------------------------------------

    @contextlib.contextmanager
    def _batch(self, announce: bool = True) -> Iterator[None]:
        """一次改动包成一批：序号只重排一遍、计数信号只发一次、dirty 只标一次。

        计数是嵌套的：``clear_files()`` 外面可能已经套了一层 ``load_entries()``，
        只有最外面那一层退出时才该重排与发信号，否则中间态会被画成一列空序号。

        ``announce=False`` 给"有可能一笔也没改成"的那几条（上移移不动、批量勾选本来就全勾着）：
        那种时候由调用方数完自己决定发不发 ``changed``。星号一旦亮了就得靠保存或关掉才灭，
        所以它只能代表"列表真的变了"。
        """
        self._bulk += 1
        try:
            yield
        finally:
            self._bulk -= 1
            if self._bulk == 0:
                self._renumber()
                self._emit_counts()
                if announce:
                    self.changed.emit()

    def _runs(self, rows: Sequence[int]) -> List[List[QListWidgetItem]]:
        """把行号表切成一段段连续的行，每段交出**项本身**（挪过之后行号会变，项不会）。"""
        runs: List[List[int]] = []
        for row in sorted(set(int(value) for value in rows)):
            if self.row_kind(row) != ROW_FILE:
                continue
            if runs and row == runs[-1][-1] + 1:
                runs[-1].append(row)
            else:
                runs.append([row])
        return [[self.item(value) for value in run] for run in runs]

    def _move_run(self, items: List[QListWidgetItem], upwards: bool) -> int:
        """一段连续的行整体挪一格；挪不动（到顶/到底/再挪就越过组头）交回 0。"""
        if not items:
            return 0
        rows = sorted(self.row(item) for item in items)
        top, bottom = rows[0], rows[-1]
        target = top - 1 if upwards else bottom + 1
        if target < 0 or target >= self.count():
            return 0
        if self.is_group(self.item(target)):
            # 越过组头就不是"挪一格"而是"换一组"，那等于悄悄改了它输出的那份文件；
            # 要换组是把整行拖过去（看得见那条指示线），不是按一下下箭头。
            return 0
        ordered = [self.item(value) for value in rows]
        for item in ordered:
            self.takeItem(self.row(item))
        insert_at = top - 1 if upwards else top + 1
        for offset, item in enumerate(ordered):
            self.insertItem(insert_at + offset, item)
        self.clearSelection()
        for item in ordered:
            item.setSelected(True)
        # 当前行要用 NoUpdate 挪：``setCurrentItem`` 走的是 ClearAndSelect，会把上面刚勾上的
        # 那几行清空——实测"选中三行下移"之后只剩一行还亮着，再按一次上移就只挪回一行。
        current = ordered[-1] if upwards else ordered[0]
        self.selectionModel().setCurrentIndex(
            self.model().index(self.row(current), 0),
            QItemSelectionModel.SelectionFlag.NoUpdate)
        return len(ordered)

    def _find(self, resolved: Path) -> Optional[QListWidgetItem]:
        for item in self._file_items():
            if self.path_of(item) == resolved:
                return item
        return None

    def _set_checked(self, row: int, checked: bool) -> None:
        item = self.item(row)
        if item is not None and not self.is_group(item):
            item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)

    def _renumber(self) -> None:
        """把序号与组头的计数重排一遍。序号只数**文件行**——它就是合并顺序，
        一条组头占了号会让人以为少了一个文件。"""
        if self._renumbering:
            return
        self._renumbering = True
        previous, self._bulk = self._bulk, self._bulk + 1
        try:
            order = 0
            for row in range(self.count()):
                item = self.item(row)
                if self.is_group(item):
                    members = self._members(row)
                    joined = sum(1 for member in members
                                 if member.checkState() == Qt.CheckState.Checked)
                    item.setText("%s｜%d 个，参与 %d 个" % (self.group_name_of(item),
                                                           len(members), joined))
                else:
                    order += 1
                    item.setText("%d. %s" % (order, self.path_of(item).name))
        finally:
            self._bulk = previous
            self._renumbering = False

    def _on_item_changed(self, _item: QListWidgetItem) -> None:
        if self._bulk:
            return
        # 用户勾掉一格，它所在组头那一行的"参与 N 个"跟着变，所以要重排一遍；
        # _renumber 自己压着 _bulk 并防重入，这里不会转圈。
        self._renumber()
        self._emit_counts()
        self.changed.emit()

    def _echo_selection(self, current: Optional[QListWidgetItem], _previous) -> None:
        if current is not None and self.is_group(current):
            return
        self.selected_changed.emit(None if current is None else self.path_of(current))

    def _emit_counts(self) -> None:
        if self._bulk:
            return
        self.counts_changed.emit(self.file_count(), self.checked_count())


def _readable(path: Path) -> bool:
    try:
        with open(path, "rb") as handle:
            handle.read(1)
        return True
    except OSError:
        return False


def _canonical(path: Path) -> Path:
    """绝对化 + 折叠 ``..``：同一个文件写成两种路径不该在列表里出现两行。"""
    try:
        return Path(path).expanduser().resolve()
    except OSError:  # 网络盘断开时 resolve 会抛，那时按原样收，别把拖放整个弄崩
        return Path(path).expanduser()


class FilePanel(QWidget):
    """文件列表那一栏：列表 + 底下那行计数。主窗口只跟这一个类打交道。"""

    noticed = Signal(str)
    convert_single = Signal(object)
    selected_changed = Signal(object)
    counts_changed = Signal(int, int)
    add_files_requested = Signal()
    add_folder_requested = Signal()
    clear_requested = Signal()
    changed = Signal()

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setObjectName("filePanel")
        self.list = FileListWidget(self)
        self.count_label = QLabel("共 0 个，勾选 0 个", self)
        self.count_label.setObjectName("hint")
        self.count_label.setWordWrap(True)

        self.list.noticed.connect(self.noticed.emit)
        self.list.convert_single.connect(self.convert_single.emit)
        self.list.selected_changed.connect(self.selected_changed.emit)
        self.list.counts_changed.connect(self._on_counts)
        self.list.add_files_requested.connect(self.add_files_requested.emit)
        self.list.add_folder_requested.connect(self.add_folder_requested.emit)
        self.list.clear_requested.connect(self.clear_requested.emit)
        self.list.changed.connect(self.changed.emit)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        layout.addWidget(self.list, 1)
        layout.addWidget(self.count_label)

    # ---- 转发给主窗口的口 ---------------------------------------------------

    def add_paths(self, paths: Iterable[Path]) -> int:
        return self.list.add_paths(paths)

    def remove_selected(self) -> List[Path]:
        return self.list.remove_selected()

    def remove_row(self, row: int) -> Optional[Path]:
        return self.list.remove_row(row)

    def move_selected(self, upwards: bool) -> int:
        return self.list.move_selected(upwards)

    def clear_files(self) -> int:
        return self.list.clear_files()

    def checked_paths(self) -> list:
        return self.list.checked_paths()

    def selected_path(self) -> Optional[Path]:
        return self.list.selected_path()

    def selected_rows(self) -> List[int]:
        return self.list.selected_rows()

    def selected_paths(self) -> List[Path]:
        return self.list.selected_paths()

    def selected_count(self) -> int:
        return self.list.selected_count()

    def set_all_checked(self, checked: bool) -> int:
        return self.list.set_all_checked(checked)

    def set_checked_for_selected(self, checked: bool) -> int:
        return self.list.set_checked_for_rows(self.list.selected_rows(), checked)

    def set_checked_for_rows(self, rows: Sequence[int], checked: bool) -> int:
        return self.list.set_checked_for_rows(rows, checked)

    def add_group(self, name: str) -> bool:
        return self.list.add_group(name)

    def prompt_group(self, row: Optional[int] = None) -> None:
        self.list.prompt_group(row)

    def groups(self) -> List[str]:
        return self.list.groups()

    def has_groups(self) -> bool:
        return self.list.has_groups()

    def group_rows(self) -> List[int]:
        return self.list.group_rows()

    def group_members(self, row: int) -> List[Path]:
        return self.list.group_members(row)

    def empty_groups(self) -> List[str]:
        return self.list.empty_groups()

    def group_plans(self) -> List[tuple]:
        return self.list.group_plans()

    def entries(self) -> List[dict]:
        return self.list.entries()

    def load_entries(self, entries: Sequence[dict], replace: bool = True) -> dict:
        return self.list.load_entries(entries, replace)

    def refresh_status(self, announce: bool = True) -> Dict[str, int]:
        return self.list.refresh_status(announce)

    def status_at(self, row: int) -> str:
        return self.list.status_at(row)

    def apply_theme(self, name: str) -> None:
        self.list.apply_theme(name)

    def select_first(self) -> Optional[Path]:
        return self.list.select_first()

    def count(self) -> int:
        """**文件**条数（不含组头）——日志与状态栏念的"列表里有几个"就是这个数。"""
        return self.list.file_count()

    def row_count(self) -> int:
        return self.list.row_count()

    def _on_counts(self, total: int, checked: int) -> None:
        groups = self.list.groups()
        text = "共 %d 个，勾选 %d 个" % (total, checked)
        if groups:
            text += "｜%d 组：%s" % (len(groups), "、".join(groups))
        self.count_label.setText(text)
        self.counts_changed.emit(total, checked)
