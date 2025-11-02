# tests/test_bxt.py
# -*- coding: utf-8 -*-

import math

from ausfunct import Bxt


def test_bxt_reference_case():
    # Referenz: vs=100_000 | age=40 | sex="M" | n=30 | t=20 | zw=12 | tarif="KLV"
    val = Bxt(100_000, 40, "M", 30, 20, 12, "KLV")
    assert math.isclose(val, 0.04226001, rel_tol=0.0, abs_tol=1e-8)
