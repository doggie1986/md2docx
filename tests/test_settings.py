# -*- coding: utf-8 -*-
"""设置这一层的自检：样式预设、一份配置文件的形状、"最近打开的文件"那条表。

这一层是**纯数据**，判据也就全在它自己的形状上，分四组：

1. **内置三档**——有哪三档、次序、以及"技术文档那一档 == ``docx_builder`` 的默认"。
   最后这一条钉的是"默认值只有一个来源"：界面、命令行、预设三处各抄一份数值，
   就会出现"不改预设也和改了预设一样"或者反过来。
2. **自定义预设的存与撤**——同名的两条不该长出两条、内置那三档改不掉也撤不掉、
   本机那份文本坏掉时**念一句而不是当成没存过**（"你自己存的那几档没了"不是人该丢的东西）。
3. **一份配置文件**——往返无损、同一次调用两次写盘逐字节相同、
   八档毛病各念自己那一句，以及**本机那几格（窗口尺寸、输出目录、上次工程文件）压根不写进去**。
   导入那一路走的也是 :func:`config.settings.parse_document`，所以"先看一眼"与"导入"
   天然同一套判定、同一句句子——这一组里那条重名的用例钉的就是这一句两边一致。
4. **最近打开的文件**——新点中的排最前、同一个文件不占两格（含大小写与 ``..`` 的写法）、
   超出上限从尾巴上砍、不改传入的那一份。

外加一条 :class:`NoQtTest`：这一层不许把 PySide6 拖进来（七期立下的那条不变式，
命令行 ``import main`` 那一路跟着它）。
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import docx_builder as db  # noqa: E402
from config import settings as st  # noqa: E402

NOW = datetime(2026, 10, 15, 9, 30, 0)

CUSTOM = st.StylePreset(name="我的标书档", body_font="仿宋", body_size=14.0,
                        heading_font="黑体")


def payload_of(values: st.StyleValues = None, custom=()) -> dict:
    """一份"这一版写得出来"的抄本，用例在上面逐格改坏。"""
    return st.to_document(values or st.StyleValues(), custom, NOW)


def message_of(action) -> str:
    """跑一遍 :class:`SettingsError` 那一路，把句子交回来；没抛就是测试失败。"""
    try:
        action()
    except st.SettingsError as exc:
        return str(exc)
    raise AssertionError("这一档本该被拒，却走通了")


class BuiltinPresetTest(unittest.TestCase):
    def test_the_three_builtin_presets_are_named_in_order(self) -> None:
        self.assertEqual(st.preset_names(), ["技术文档", "学术论文", "简洁版"])

    def test_technical_doc_equals_the_builder_default(self) -> None:
        """预设里的"技术文档"不许自己抄一份数值——默认住在 docx_builder，两处各一份就会分叉。"""
        tech = st.preset_named("技术文档")
        self.assertEqual((tech.body_font, tech.body_size, tech.heading_font),
                         (db.BODY_FONT_DEFAULT, db.BODY_SIZE_DEFAULT, db.HEADING_FONT_DEFAULT))

    def test_every_builtin_is_marked_and_described(self) -> None:
        for preset in st.BUILTIN_PRESETS:
            self.assertTrue(preset.builtin)
            self.assertTrue(preset.note, preset.name)
            self.assertIn("pt", preset.describe())

    def test_the_three_are_actually_different(self) -> None:
        """三档要是排出来一个样，那一栏就是骗人的。"""
        shapes = {(p.body_font, p.body_size, p.heading_font) for p in st.BUILTIN_PRESETS}
        self.assertEqual(len(shapes), 3)

    def test_custom_goes_behind_the_builtins(self) -> None:
        names = st.preset_names([CUSTOM])
        self.assertEqual(names[:3], st.builtin_names())
        self.assertEqual(names[3:], ["我的标书档"])

    def test_a_preset_without_a_custom_list_still_reads_the_builtins(self) -> None:
        self.assertEqual(len(st.presets()), 3)
        self.assertIsNone(st.preset_named("我的标书档"))


class CustomPresetTest(unittest.TestCase):
    def test_adding_keeps_the_builtin_and_appends(self) -> None:
        result = st.add_custom(st.presets(), CUSTOM)
        self.assertEqual([item.name for item in result][-1:], ["我的标书档"])
        self.assertFalse(result[-1].builtin)

    def test_a_custom_one_never_carries_the_builtin_note(self) -> None:
        result = st.add_custom([], CUSTOM)
        self.assertEqual(result[0].note, "")

    def test_same_name_replaces_instead_of_growing_a_second_row(self) -> None:
        again = st.StylePreset(name="我的标书档", body_font="楷体", body_size=10.0,
                               heading_font="楷体")
        result = st.add_custom(st.add_custom([], CUSTOM), again)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].body_font, "楷体")

    def test_a_custom_name_may_not_shadow_a_builtin_one(self) -> None:
        """顶掉内置那一档之后"回到内置"就没地方了，而下拉里两条同名点不出是哪一个。"""
        clone = st.StylePreset(name="学术论文", body_font="Arial", body_size=11.0,
                               heading_font="Arial")
        text = message_of(lambda: st.add_custom([], clone))
        self.assertIn("内置的「学术论文」", text)
        self.assertIn("不会被改掉", text)

    def test_the_input_is_never_mutated(self) -> None:
        current = [CUSTOM]
        st.add_custom(current, st.StylePreset(name="另一档", body_font="宋体", body_size=9.0,
                                              heading_font="宋体"))
        self.assertEqual(len(current), 1)

    def test_removing_a_builtin_is_refused_not_ignored(self) -> None:
        """撤不掉也要说一句：什么都不做会被读成"撤掉了，怎么还在"。"""
        text = message_of(lambda: st.remove_custom(st.presets(), "简洁版"))
        self.assertIn("内置的三档之一", text)

    def test_removing_a_missing_name_says_so(self) -> None:
        text = message_of(lambda: st.remove_custom([CUSTOM], "没存过的那一档"))
        self.assertIn("自定义预设里没有", text)

    def test_removing_works(self) -> None:
        self.assertEqual(st.remove_custom([CUSTOM], "我的标书档"), [])

    def test_the_three_filled_cells_are_checked(self) -> None:
        for bad, word in (
            (st.StylePreset(name="", body_font="宋体", body_size=12.0, heading_font="宋体"), "空的"),
            (st.StylePreset(name="坏", body_font="", body_size=12.0, heading_font="宋体"), "正文字体"),
            (st.StylePreset(name="坏", body_font="宋体", body_size=12.0, heading_font=""), "标题字体"),
            (st.StylePreset(name="坏", body_font="宋体", body_size=90.0, heading_font="宋体"), "得在"),
            (st.StylePreset(name="坏", body_font="宋体", body_size=2.0, heading_font="宋体"), "得在"),
        ):
            text = message_of(lambda b=bad: st.add_custom([], b))
            self.assertIn(word, text, word)


class CustomStorageTest(unittest.TestCase):
    def test_encode_decode_round_trip(self) -> None:
        text = st.encode_custom([CUSTOM])
        back = st.decode_custom(text)
        self.assertEqual(len(back), 1)
        self.assertEqual((back[0].name, back[0].body_font, back[0].body_size,
                          back[0].heading_font),
                         ("我的标书档", "仿宋", 14.0, "黑体"))
        self.assertFalse(back[0].builtin)

    def test_empty_text_means_no_custom_preset(self) -> None:
        """这一格空着是"没存过"，不是坏档：首屏就该看见内置那三档而不是报错。"""
        self.assertEqual(st.decode_custom(""), [])
        self.assertEqual(st.decode_custom("   "), [])

    def test_builtins_are_not_written_into_the_local_blob(self) -> None:
        body = json.loads(st.encode_custom(st.presets([CUSTOM])))
        self.assertEqual([item["name"] for item in body], ["我的标书档"])

    def test_a_broken_blob_is_said_out_loud(self) -> None:
        """静默交回空列表 = "你自己存的那几档没了"，而人不会为一句报错去查这件事。"""
        text = message_of(lambda: st.decode_custom("{不是 JSON"))
        self.assertIn("读不进", text)
        self.assertIn("第 ", text)

    def test_a_builtin_name_in_the_blob_is_refused(self) -> None:
        blob = json.dumps([{"name": "简洁版", "body_font": "宋体", "body_size": 12.0,
                            "heading_font": "宋体"}], ensure_ascii=False)
        self.assertIn("内置的档名", message_of(lambda: st.decode_custom(blob)))

    def test_a_doubled_name_in_the_blob_is_refused(self) -> None:
        one = {"name": "同一档", "body_font": "宋体", "body_size": 12.0, "heading_font": "宋体"}
        self.assertIn("重名", message_of(lambda: st.decode_custom(json.dumps([one, one]))))

    def test_a_bad_number_in_the_blob_is_named(self) -> None:
        blob = json.dumps([{"name": "坏档", "body_font": "宋体", "body_size": "十二",
                            "heading_font": "宋体"}], ensure_ascii=False)
        self.assertIn("字号", message_of(lambda: st.decode_custom(blob)))


class ConfigFileTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="md2docx-config-"))
        # 临时目录归临时目录：不清就会每跑一条用例在 %TEMP% 里落一个，跑测的人不会知道
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.where = self.tmp / "settings.json"

    def test_round_trip_keeps_all_seven_cells_and_the_customs(self) -> None:
        values = st.StyleValues(body_font="楷体", body_size=10.5, heading_font="微软雅黑",
                                with_toc=False, page_break=True, source_label=False,
                                cover_page=True)
        st.write_config(self.where, values, [CUSTOM], NOW)
        back, custom = st.read_config(self.where)
        self.assertEqual(back.as_dict(), values.as_dict())
        self.assertEqual([item.name for item in custom], ["我的标书档"])

    def test_the_same_call_writes_the_same_bytes(self) -> None:
        """差异比对要读得出"设置到底变没变"，所以同一次调用导两次必须逐字节相同。"""
        st.write_config(self.where, st.StyleValues(), [CUSTOM], NOW)
        first = self.where.read_bytes()
        st.write_config(self.where, st.StyleValues(), [CUSTOM], NOW)
        self.assertEqual(first, self.where.read_bytes())

    def test_newlines_stay_plain_lf_on_windows(self) -> None:
        st.write_config(self.where, st.StyleValues(), (), NOW)
        raw = self.where.read_bytes()
        self.assertNotIn(b"\r\n", raw)

    def test_the_document_has_the_four_top_cells_and_the_seven_settings(self) -> None:
        body = st.to_document(st.StyleValues(), [CUSTOM], NOW)
        self.assertEqual(sorted(body), sorted(st.TOP_KEYS))
        self.assertEqual(sorted(body["settings"]), sorted(st.SETTINGS_KEYS))
        self.assertEqual(body["kind"], st.KIND)
        self.assertEqual(body["version"], st.VERSION)

    def test_a_file_without_the_custom_cell_is_valid(self) -> None:
        """一份文件没有自定义预设是它的真相，不该当成缺了一格。"""
        body = payload_of()
        del body["style_presets"]
        values, custom = st.parse_document(body)
        self.assertEqual(custom, [])
        self.assertEqual(values.body_font, db.BODY_FONT_DEFAULT)

    def test_the_machine_local_cells_are_not_written(self) -> None:
        """窗口尺寸 / 输出目录 / 上次工程文件都是这台机器的事，带过去就是指一个不存在的路径。"""
        body = json.dumps(payload_of(), ensure_ascii=False)
        for word in ("geometry", "splitter", "output", "directory", "last_path", "theme",
                     "window", "recent"):
            self.assertNotIn(word, body, word)

    def test_not_carried_and_carried_do_not_overlap(self) -> None:
        """以后有人想把本机那一格加进抄本，这条会红——那是设计决定，不是顺手加的一格。"""
        joined = " ".join(st.NOT_CARRIED).lower()
        for key in st.SETTINGS_KEYS:
            self.assertNotIn(key, joined, key)

    def test_config_filter_names_json(self) -> None:
        self.assertIn("*.json", st.config_filter())


class ConfigRefusalTest(unittest.TestCase):
    """八档毛病各念自己那一句，而且**点名是哪一格**——"这一份读不懂"等于没说。"""

    def parse(self, mutate):
        body = payload_of(custom=[CUSTOM])
        mutate(body)
        return message_of(lambda: st.parse_document(body))

    def test_wrong_kind_says_what_it_found(self) -> None:
        text = self.parse(lambda body: body.__setitem__("kind", "md2docx.project"))
        self.assertIn("不是 md2docx 的配置文件", text)
        self.assertIn("md2docx.project", text)

    def test_a_future_version_is_refused(self) -> None:
        self.assertIn("只写到", self.parse(lambda body: body.__setitem__("version", 99)))

    def test_a_non_integer_version_is_refused(self) -> None:
        self.assertIn("得是一个整数", self.parse(lambda body: body.__setitem__("version", "1")))

    def test_an_unknown_top_cell_is_refused(self) -> None:
        text = self.parse(lambda body: body.__setitem__("zoom", 2))
        self.assertIn("顶层有读不懂的键", text)
        self.assertIn("zoom", text)

    def test_a_missing_settings_cell_is_refused(self) -> None:
        """缺 with_toc 到底是"目录关着"还是"没写"，没人替它猜。"""
        text = self.parse(lambda body: body["settings"].pop("with_toc"))
        self.assertIn("settings 里缺了这几格", text)
        self.assertIn("with_toc", text)

    def test_an_unknown_settings_cell_is_refused(self) -> None:
        self.assertIn("settings 里有读不懂的键",
                      self.parse(lambda body: body["settings"].__setitem__("code_font", "Consolas")))

    def test_a_switch_that_is_not_true_or_false_is_refused(self) -> None:
        """1 也算拒：True 是 1 的对子，"关着"与"没写"在这一格里就该分得开。"""
        self.assertIn("得是真或假",
                      self.parse(lambda body: body["settings"].__setitem__("page_break", 1)))

    def test_a_blank_font_name_is_refused(self) -> None:
        self.assertIn("正文字体", self.parse(lambda body: body["settings"].__setitem__("body_font", "  ")))

    def test_a_size_out_of_range_is_refused(self) -> None:
        text = self.parse(lambda body: body["settings"].__setitem__("body_size", 200))
        self.assertIn("得在 6 ~ 72 之间", text)
        self.assertIn("200", text)

    def test_a_size_that_is_text_is_refused(self) -> None:
        self.assertIn("得是一个数", self.parse(lambda body: body["settings"].__setitem__("body_size", "12")))

    def test_a_preset_row_that_is_not_an_object_is_refused(self) -> None:
        self.assertIn("不是一个对象",
                      self.parse(lambda body: body.__setitem__("style_presets", ["字符串"])))

    def test_a_preset_row_missing_a_cell_is_refused(self) -> None:
        text = self.parse(lambda body: body["style_presets"][0].pop("heading_font"))
        self.assertIn("第 1 条自定义预设里缺了这几格", text)
        self.assertIn("heading_font", text)

    def test_a_preset_row_with_an_extra_cell_is_refused(self) -> None:
        self.assertIn("读不懂的键",
                      self.parse(lambda body: body["style_presets"][0].__setitem__("note", "我的")))

    def test_a_preset_row_sharing_a_builtin_name_is_refused(self) -> None:
        """导入那一路与「另存为预设」那一路念的是同一句——两道闸各判一套就出现
        "档案里进得来、下拉里点不出"。"""
        clone = dict(name="技术文档", body_font="Arial", body_size=11.0, heading_font="Arial")
        text = self.parse(lambda body: body.__setitem__("style_presets", [clone]))
        self.assertIn("内置的「技术文档」", text)
        self.assertEqual(text, message_of(
            lambda: st.add_custom([], st.StylePreset(**clone))))

    def test_a_preset_cell_that_is_not_a_list_is_refused(self) -> None:
        self.assertIn("得是一个列表",
                      self.parse(lambda body: body.__setitem__("style_presets", {"name": "怪"})))

    def test_a_document_that_is_not_an_object_is_refused(self) -> None:
        self.assertIn("现在是列表", message_of(lambda: st.parse_document([1, 2])))

    def test_a_missing_file_is_named(self) -> None:
        text = message_of(lambda: st.read_config(Path(tempfile.gettempdir()) / "没有这个文件.json"))
        self.assertIn("这个文件不在那儿", text)

    def test_a_broken_json_file_reports_the_place(self) -> None:
        folder = Path(tempfile.mkdtemp(prefix="md2docx-config-bad-"))
        self.addCleanup(shutil.rmtree, folder, ignore_errors=True)
        tmp = folder / "b.json"
        tmp.write_text("{坏", encoding="utf-8", newline="\n")
        self.assertIn("不是合法的 JSON", message_of(lambda: st.read_config(tmp)))


class RecentFilesTest(unittest.TestCase):
    """表上存的是**规范化后的写法**（``\`` 分隔、``..`` 收平、大小写不动），比"是不是同一条"才判得准。"""

    def test_the_new_one_goes_to_the_front(self) -> None:
        self.assertEqual(st.push_recent(["C:\\旧\\a.json", "C:\\旧\\b.json"], ["C:/新/c.json"]),
                         ["C:\\新\\c.json", "C:\\旧\\a.json", "C:\\旧\\b.json"])

    def test_reopening_a_file_does_not_take_two_rows(self) -> None:
        result = st.push_recent(["C:\\旧\\a.json", "C:\\旧\\b.json"], ["C:\\旧\\a.json"])
        self.assertEqual(result, ["C:\\旧\\a.json", "C:\\旧\\b.json"])

    def test_the_same_file_written_two_ways_is_one_row(self) -> None:
        """大小写与 ``..`` 在 Windows 上是同一个文件；不规范化就会长出两格，砍的次序也跟着乱。"""
        self.assertEqual(st.push_recent(["C:\\旧\\a.json"], ["c:/旧/a.json"]),
                         ["c:\\旧\\a.json"])
        self.assertEqual(st.push_recent(["C:\\旧\\a.json"], ["c:/旧/子目录/../a.json"]),
                         ["c:\\旧\\a.json"])

    def test_the_cut_is_made_at_the_tail(self) -> None:
        """满表的时候再点一份新的，掉的是最旧那一条，不是中间某一条。"""
        existing = ["C:\\%d.json" % index for index in range(1, st.RECENT_LIMIT + 1)]
        result = st.push_recent(existing, ["C:\\100.json"])
        self.assertEqual(result[0], "C:\\100.json")
        self.assertEqual(len(result), st.RECENT_LIMIT)
        self.assertNotIn("C:\\%d.json" % st.RECENT_LIMIT, result)

    def test_a_batch_keeps_the_given_order(self) -> None:
        self.assertEqual(st.push_recent([], ["C:/a.json", "C:/b.json"]),
                         ["C:\\a.json", "C:\\b.json"])

    def test_blank_paths_are_not_kept(self) -> None:
        self.assertEqual(st.push_recent(["  ", "C:\\a.json"], [None, ""]), ["C:\\a.json"])

    def test_the_stored_list_is_not_mutated(self) -> None:
        existing = ["C:\\a.json"]
        st.push_recent(existing, ["C:\\b.json"])
        self.assertEqual(existing, ["C:\\a.json"])

    def test_a_limit_of_zero_is_refused(self) -> None:
        """0 不是"记 0 条"这个决定，多半是填错了：宁可当场问一句。"""
        self.assertIn("至少留 1 条", message_of(lambda: st.push_recent([], ["C:\\a.json"], 0)))


class NoQtTest(unittest.TestCase):
    """这一层不 import Qt——命令行与测试都靠它，"``import main`` 不把 PySide6 拖进来"是同一条。"""

    def test_importing_the_settings_layer_loads_no_pyside(self) -> None:
        probe = ("import sys; sys.path.insert(0, %r); import config; "
                 "print('PySide6' in sys.modules or 'PySide2' in sys.modules)"
                 % str(Path(__file__).resolve().parent.parent))
        done = subprocess.run([sys.executable, "-c", probe], capture_output=True, text=True)
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertEqual(done.stdout.strip(), "False")

    def test_the_panel_reads_its_size_range_from_here(self) -> None:
        """字号范围两处在，就会出现"预设里填得进、旋钮拧不到"。"""
        from gui import settings_panel as sp

        self.assertEqual((sp.SIZE_MIN, sp.SIZE_MAX), (st.SIZE_MIN, st.SIZE_MAX))
        self.assertEqual((st.SIZE_MIN, st.SIZE_MAX), (6.0, 72.0))


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
