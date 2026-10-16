# -*- coding: utf-8 -*-
"""设置与样式预设这一层的包口。

只 re-export :mod:`config.settings` 里对外的那几个名字，**这里不 import Qt**：
命令行（:mod:`main`）与测试要能读同一份形状，而"``import main`` 不把 PySide6 拖进来"
是七期立下的那条不变式。
"""

from __future__ import annotations

from config.settings import (BUILTIN_PRESETS, KIND, RECENT_LIMIT, SETTINGS_KEYS, SIZE_MAX,
                             SIZE_MIN, TOP_KEYS, VERSION, StylePreset, StyleValues,
                             SettingsError, add_custom, builtin_names, config_filter,
                             decode_custom, encode_custom, parse_document, preset_named,
                             preset_names, presets, push_recent, read_config, remove_custom,
                             to_document, write_config)

__all__ = ["KIND", "VERSION", "TOP_KEYS", "SETTINGS_KEYS", "SIZE_MIN", "SIZE_MAX",
           "RECENT_LIMIT", "StylePreset", "StyleValues", "SettingsError",
           "BUILTIN_PRESETS", "builtin_names", "preset_names", "presets", "preset_named",
           "add_custom", "remove_custom", "encode_custom", "decode_custom",
           "to_document", "parse_document", "write_config", "read_config", "config_filter",
           "push_recent"]
