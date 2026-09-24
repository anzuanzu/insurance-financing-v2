#!/usr/bin/env python3
"""Regression checks for the LibreOffice-backed premium-financing calculator."""

from __future__ import annotations

import copy
from pathlib import Path

import insurance_calculation_server as service


EXPECTED = {
    "ABA": {"age": 30, "gender": "male", "face": 95000, "rate": 4.35, "premium": 100291, "reference": 70720},
    "WEF": {"age": 19, "gender": "male", "face": 2000000, "rate": 4.15, "premium": 358380, "reference": 223007},
    "WJA": {"age": 1, "gender": "male", "face": 300000, "rate": 4.35, "premium": 30510, "reference": 24053},
    "WQ6": {"age": 19, "gender": "male", "face": 200000, "rate": 4.25, "premium": 135519, "reference": 95325},
    "WUS": {"age": 40, "gender": "male", "face": 1840000, "rate": 4.30, "premium": 500758, "reference": 359085},
}


def quote(profile: dict, expected: dict, rate: float | None = None) -> dict:
    return service.calculate_quote(profile, {
        "profileId": profile["id"],
        "age": expected["age"],
        "gender": expected["gender"],
        "faceAmount": expected["face"],
        "declaredRate": expected["rate"] if rate is None else rate,
    })


def main() -> None:
    store = service.ProfileStore()
    store.bootstrap_workspace_sources()
    profiles = {item["webCode"]: item for item in store._profiles}
    assert set(EXPECTED).issubset(profiles), "五項來源商品未完整載入"

    for code, expected in EXPECTED.items():
        source_path = service.profile_source_path(profiles[code])
        source_hash_before = service.sha256_file(source_path)
        result = quote(profiles[code], expected)
        assert service.sha256_file(source_path) == source_hash_before, f"{code} 的來源建議書不應被重算流程修改"
        assert profiles[code]["dividendOption"] == service.DIVIDEND_OPTION_CONTINUE, f"{code} 未固定為持續增購保額"
        assert "dividendOption" in profiles[code]["inputCells"], f"{code} 找不到給付方式輸入格"
        assert result["dividendOption"] == service.DIVIDEND_OPTION_CONTINUE, f"{code} 回傳的給付方式不正確"
        assert abs(result["premium"] - expected["premium"]) <= 1, f"{code} 保費不符"
        assert abs(result["referenceCashValue"] - expected["reference"]) <= 1, f"{code} 融資參考現金價值不符"
        assert result["benefitByYear"][0] is not None, f"{code} 第一年保障不可為空"
        assert result["cashValueByYear"][19] is not None, f"{code} 第二十年現金價值不可為空"
        print(f"PASS {code}: premium={result['premium']:,.0f}, reference={result['referenceCashValue']:,.0f}")

    aba = quote(profiles["ABA"], EXPECTED["ABA"])
    assert abs(aba["benefitByYear"][19] - 216607) <= 1, "ABA 持續增購保額後的第 20 年保障不符"
    assert abs(aba["cashValueByYear"][19] - 222779.85) <= 0.01, "ABA 持續增購保額後的第 20 年現金價值不符"
    print("PASS ABA forced to 第7年起持續增購保額")

    wus = profiles["WUS"]
    higher = quote(wus, EXPECTED["WUS"], 4.30)
    lower = quote(wus, EXPECTED["WUS"], 4.20)
    assert higher["cashValueByYear"][19] > lower["cashValueByYear"][19], "WUS 宣告利率提高後第 20 年現金價值應上升"
    assert higher["benefitByYear"][19] > lower["benefitByYear"][19], "WUS 宣告利率提高後第 20 年保障應上升"
    assert higher["referenceCashValue"] == lower["referenceCashValue"], "投保日融資參考現金價值不應受宣告利率影響"
    print("PASS WUS declared-rate sensitivity and fixed inception financing reference")

    saved = copy.deepcopy(store._profiles)
    try:
        probe = copy.deepcopy(profiles["WJA"])
        probe["name"] = "國泰人壽匯入驗證新商品"
        probe["normalizedName"] = service.normalized_text(probe["name"])
        probe["id"] = "product-import-verification"
        probe["sourceCode"] = "VERIFY"
        probe["webCode"] = "VERIFY"
        _, action = store.upsert(probe, persist=False)
        assert action == "created", "新商品名稱應新增商品"
        probe["declaredRate"] = 0.049
        probe["sourceHash"] = "verification-changed-source-hash"
        _, action = store.upsert(probe, persist=False)
        assert action == "updated", "同名商品應更新既有條件"
        print("PASS import naming rule: new name creates, same name updates")
    finally:
        store._profiles = saved
        store._save()

    assert not service.has_financing_filename("一般建議書.xlsx"), "檔名限制失效"
    assert service.has_financing_filename("商品_保費融資.xlsx"), "檔名限制誤拒"
    print("PASS filename gate")
    print("ALL VERIFICATIONS PASSED")


if __name__ == "__main__":
    main()

