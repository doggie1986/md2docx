# -*- coding: utf-8 -*-
"""跑批完成后弹的那一张报告：只画 :mod:`gui.tasks` 交回的统计，自己不数一个数。

需求第 5 条点名的四个数（成功数、失败数、公式数、表格数）在这里是 ``head`` 那几对，
逐份的账是 ``columns`` + ``rows`` 那张表，坏掉的那些是 ``problem_columns`` + ``problems``，
``notes`` 装"这一批要单独念一句、但不属于某一条问题"的那些说明（转不出的公式只演了前几条、
读回核对差了几条）。列名与单元格都在执行体里定死
（:data:`gui.tasks.COLUMNS_OF_OUTPUT`），这里不拼句子也不数"第几列是什么"——
两处各数一遍就会出现"日志说 3 个文件、报告说 2 个"。

三条不是显而易见的取舍：

* **``open()`` 而不是 ``exec()``**：``exec()`` 开的是应用级模态，会把事件循环关进去，
  ``--check``、冒烟测试与自动化脚本就全卡在这张窗上。``open()`` 是**窗口级**模态——
  报告开着时主窗口点不动（人总该先看完这一批的账），但事件循环照跑，
  所以主窗口 ``close()`` 与测试都能往下走。
* **一扇窗重复用**：主窗口只造一个实例、每一批用 :meth:`set_stats` 换内容。
  每批另开一扇的话，连点三次「开始转换」屏幕上是三张报告，人分不清哪张是最新那一批，
  而关窗时也得挨个关。
* **空的那一栏不摆空表**：一条都没有时演一句"为什么没有"（取消那一路换一句说法）。
    空表会被读成"这一栏坏了"，而真实原因（一个都没勾上 / 全读不进 / 刚点了取消）人看不见。
"""

from __future__ import annotations

from typing import Optional, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QAbstractItemView, QDialog, QGridLayout, QGroupBox, QHBoxLayout,
                               QHeaderView, QLabel, QPushButton, QTableWidget, QTableWidgetItem,
                               QVBoxLayout, QWidget)

__all__ = ["ReportDialog", "NO_ROWS_TEXT", "NO_ROWS_CANCELLED_TEXT", "NO_PROBLEM_TEXT",
           "NO_PROBLEM_CANCELLED_TEXT", "NO_NOTES_TEXT"]

#: 一张表一条都没有时念的那一句。说清"没有"是这一批的结果，不是这一栏坏了。
NO_ROWS_TEXT = "这一批一条都没走成（原因看下面的「问题与提醒」与「说明」）。"
NO_ROWS_CANCELLED_TEXT = "刚点了取消，一份都还没走完，所以下面这张表是空的。"
NO_PROBLEM_TEXT = "没有问题：勾上的文件每个都读进、转完、写出去了。"
NO_PROBLEM_CANCELLED_TEXT = "已经走到的那些没有问题（没走到的那些这次不算）。"
NO_NOTES_TEXT = "没有别的说明——上面两张表就是这一批全部的账。"


def _make_table(wrap_last: bool) -> QTableWidget:
    """两张表同样的只读画法，只有"最后一格装的是整句话还是一条路径"这一档不同。

    「逐份的账」八列里有整条输出路径，撑开反而读不完，所以给它横向滚动条（``False``）；
    「问题与提醒」只有两列、第二列装的是整句话，所以那一格换行、行高跟着长（``True``）——
    撑到窗口边仍读不完的那一句会被画成"…（NoAvailableTokensError: 无…"，
    而这一栏存在的理由就是把"为什么"念给人听，裁在中间等于没写。
    """
    table = QTableWidget(0, 0)
    table.setObjectName("reportTable")
    table.verticalHeader().setVisible(False)
    table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
    table.setWordWrap(wrap_last)
    table.setDragEnabled(False)
    table.setAcceptDrops(False)
    table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
    header = table.horizontalHeader()
    header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
    header.setStretchLastSection(wrap_last)
    header.setSectionsMovable(False)
    if wrap_last:
        # 换行之后格子的高度由内容算：窗口拉宽/压窄时行高跟着重算，不然拖一下窗口就裁字。
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
    else:
        table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
    return table


def _clear(layout: QGridLayout) -> None:
    while layout.count():
        item = layout.takeAt(0)
        widget = item.widget()
        if widget is not None:
            widget.deleteLater()


def _fill(table: QTableWidget, columns: Sequence[str], rows: Sequence[Sequence[str]]) -> None:
    table.clear()
    table.setColumnCount(len(columns))
    table.setHorizontalHeaderLabels([str(column) for column in columns])
    table.setRowCount(len(rows))
    for row_index, row in enumerate(rows):
        for column_index in range(len(columns)):
            text = str(row[column_index]) if column_index < len(row) else ""
            item = QTableWidgetItem(text)
            item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            table.setItem(row_index, column_index, item)
    table.resizeColumnsToContents()
    if table.wordWrap():
        # 列宽定下来才算行高（反过来的话量到的是"整句话一行放下"的那个高度，等于没换行）。
        # 这一句其实不是唯一的一条路：verticalHeader 已经是 ResizeToContents，Qt 自己也会重算——
        # 留着是因为换一批内容时"先量列宽再量行高"这个次序写在代码里才看得见。
        table.resizeRowsToContents()


class ReportDialog(QDialog):
    """报告对话框。主窗口每批只发一次 :meth:`set_stats` 再 ``open()``。"""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setObjectName("reportDialog")
        self.setWindowTitle("报告")
        self.resize(980, 720)
        self._stats: dict = {}

        self.receipt = QLabel("", self)
        self.receipt.setObjectName("panelTitle")
        self.receipt.setWordWrap(True)
        self.receipt.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        self.head_box = QGroupBox("这一批的账", self)
        self.head_grid = QGridLayout(self.head_box)
        self.head_grid.setColumnStretch(1, 1)
        self.head_grid.setColumnStretch(3, 1)

        self.output_box = QGroupBox("逐份的账", self)
        output_layout = QVBoxLayout(self.output_box)
        self.output_table = _make_table(wrap_last=False)
        self.output_empty = QLabel(NO_ROWS_TEXT, self.output_box)
        self.output_empty.setObjectName("hint")
        self.output_empty.setWordWrap(True)
        output_layout.addWidget(self.output_table)
        output_layout.addWidget(self.output_empty)

        self.problem_box = QGroupBox("问题与提醒", self)
        problem_layout = QVBoxLayout(self.problem_box)
        self.problem_table = _make_table(wrap_last=True)
        self.problem_empty = QLabel(NO_PROBLEM_TEXT, self.problem_box)
        self.problem_empty.setObjectName("hint")
        self.problem_empty.setWordWrap(True)
        problem_layout.addWidget(self.problem_table)
        problem_layout.addWidget(self.problem_empty)

        self.notes_box = QGroupBox("说明", self)
        self.notes_grid = QGridLayout(self.notes_box)
        self.notes_grid.setColumnStretch(0, 1)

        self.close_button = QPushButton("关闭", self)
        self.close_button.setObjectName("act_close_report")
        self.close_button.clicked.connect(self.close)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        buttons.addWidget(self.close_button)

        root = QVBoxLayout(self)
        root.addWidget(self.receipt)
        root.addWidget(self.head_box)
        root.addWidget(self.output_box, 3)
        root.addWidget(self.problem_box, 2)
        root.addWidget(self.notes_box)
        root.addLayout(buttons)

    # ---- 填内容 -------------------------------------------------------------

    def set_stats(self, stats: Optional[dict]) -> None:
        """把执行体交回的那一份 dict 画到屏上；``None`` 就当没这批（不改已有内容）。"""
        if not stats:
            return
        self._stats = dict(stats)
        cancelled = bool(self._stats.get("cancelled"))

        self.setWindowTitle(str(self._stats.get("title") or "报告"))
        self.receipt.setText(str(self._stats.get("receipt") or ""))
        # 取消那一路要一眼和普通报告分开：红色加粗，而不是让人读完那句"已取消"才反应过来。
        # 换 objectName 之后必须 unpolish + polish：样式表是建控件那一刻按名字匹配的，
        # 光改名字 Qt 不会重算——这一窗是复用同一实例的（一扇窗每批换内容），所以每一批都得重算一次。
        wanted = "problem" if cancelled else "panelTitle"
        if self.receipt.objectName() != wanted:
            self.receipt.setObjectName(wanted)
            self.receipt.style().unpolish(self.receipt)
            self.receipt.style().polish(self.receipt)

        _clear(self.head_grid)
        for index, pair in enumerate(self._pairs("head")):
            row, column = divmod(index, 2)
            name = QLabel(pair[0], self.head_box)
            name.setObjectName("hint")
            body = QLabel(pair[1], self.head_box)
            body.setWordWrap(True)
            body.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            self.head_grid.addWidget(name, row, column * 2)
            self.head_grid.addWidget(body, row, column * 2 + 1)

        rows = self._rows("rows", "columns")
        self.output_table.setVisible(bool(rows))
        self.output_empty.setVisible(not rows)
        if rows:
            _fill(self.output_table, list(self._stats.get("columns") or []), rows)
        else:
            self.output_empty.setText(NO_ROWS_CANCELLED_TEXT if cancelled else NO_ROWS_TEXT)

        problems = self._rows("problems", "problem_columns")
        self.problem_table.setVisible(bool(problems))
        self.problem_empty.setVisible(not problems)
        if problems:
            _fill(self.problem_table, list(self._stats.get("problem_columns") or []), problems)
        else:
            self.problem_empty.setText(NO_PROBLEM_CANCELLED_TEXT if cancelled else NO_PROBLEM_TEXT)

        _clear(self.notes_grid)
        notes = [str(note) for note in (self._stats.get("notes") or [])]
        if not notes:
            notes = [NO_NOTES_TEXT]
        for index, note in enumerate(notes):
            body = QLabel(note, self.notes_box)
            body.setObjectName("hint")
            body.setWordWrap(True)
            self.notes_grid.addWidget(body, index, 0)

    def show_report(self) -> None:
        """窗口级模态地弹出来：挡住主窗口，但不关事件循环。"""
        self.open()
        self.raise_()
        self.activateWindow()

    # ---- 读回来给测试用 ------------------------------------------------------

    def stats(self) -> dict:
        return dict(self._stats)

    def receipt_text(self) -> str:
        return self.receipt.text()

    def head_pairs(self) -> list[list[str]]:
        return [[str(pair[0]), str(pair[1]) if len(pair) > 1 else ""]
                for pair in self._pairs("head")]

    def output_rows(self) -> list[list[str]]:
        return self._table_values(self.output_table)

    def problem_rows(self) -> list[list[str]]:
        return self._table_values(self.problem_table)

    def note_texts(self) -> list[str]:
        texts = []
        for index in range(self.notes_grid.count()):
            widget = self.notes_grid.itemAt(index).widget()
            if isinstance(widget, QLabel):
                texts.append(widget.text())
        return texts

    # ---- 内部 ---------------------------------------------------------------

    def _pairs(self, key: str) -> list[list[str]]:
        result = []
        for pair in (self._stats.get(key) or []):
            if isinstance(pair, (list, tuple)):
                result.append([str(part) for part in pair])
            else:
                result.append([str(pair), ""])
        return result

    def _rows(self, rows_key: str, columns_key: str) -> list[list[str]]:
        columns = list(self._stats.get(columns_key) or [])
        result = []
        for row in (self._stats.get(rows_key) or []):
            values = [str(cell) for cell in row]
            if len(values) < len(columns):
                values += [""] * (len(columns) - len(values))
            result.append(values[:len(columns)] if columns else values)
        return result

    @staticmethod
    def _table_values(table: QTableWidget) -> list[list[str]]:
        values = []
        for row in range(table.rowCount()):
            cells = []
            for column in range(table.columnCount()):
                item = table.item(row, column)
                cells.append(item.text() if item is not None else "")
            values.append(cells)
        return values
