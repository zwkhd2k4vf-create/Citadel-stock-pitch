# -*- coding: utf-8 -*-
"""v13 memo text (edited in memo_v13.json). Numbers tie to NCLH_model_v6.xlsx."""
import json
import os

_d = json.load(open(os.path.join(os.path.dirname(__file__), 'memo_v13.json'), encoding='utf-8'))
PAGE1 = [tuple(x) for x in _d['PAGE1']]
PAGE2 = [tuple(x) for x in _d['PAGE2']]
SPLIT_LABEL, SPLIT_TEXT, KPI_TABLE = _d['SPLIT_LABEL'], _d['SPLIT_TEXT'], _d['KPI_TABLE']
AFTER = [tuple(x) for x in _d['AFTER']]
SCEN_TABLE, SOURCES = _d['SCEN_TABLE'], _d['SOURCES']
