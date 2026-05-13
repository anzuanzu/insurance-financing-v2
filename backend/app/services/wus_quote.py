from functools import lru_cache
from pathlib import Path

from app.schemas import Gender, QuoteRequest, QuoteResponse
from scripts import rebuild_wus_coverage


ROOT = Path(__file__).resolve().parents[3]
SOURCE_WORKBOOK = ROOT / "WUS(金控)國泰人壽鑽美滿利率變動型美元終身壽險(定期給付型)-11501版(宣告4.25%-1150101啟用)(保費融資)_i值.xlsx"
DISCOUNT_TIERS = (
    {"minTablePremium": 0, "discount": 0.0},
    {"minTablePremium": 500_000, "discount": 0.01},
)
BASE_FACE_AMOUNT = rebuild_wus_coverage.BASE_FACE_AMOUNT
UNIT_FACE = rebuild_wus_coverage.UNIT_FACE
def _round_int(value: float) -> int:
    return int(rebuild_wus_coverage.xl_round(value, 0))


def resolve_paid_amount_from_table_premium(table_premium: float) -> dict:
    rounded_table_premium = _round_int(table_premium)
    matched = DISCOUNT_TIERS[0]
    for tier in DISCOUNT_TIERS:
        if rounded_table_premium >= tier["minTablePremium"]:
            matched = tier
    discount_amount = _round_int(rounded_table_premium * matched["discount"])
    return {
        "tablePremium": rounded_table_premium,
        "paidPremium": rounded_table_premium - discount_amount,
        "discount": matched["discount"],
    }


def resolve_table_premium_from_paid_amount(paid_premium: float) -> dict:
    for idx in range(len(DISCOUNT_TIERS) - 1, -1, -1):
        current = DISCOUNT_TIERS[idx]
        next_tier = DISCOUNT_TIERS[idx + 1] if idx + 1 < len(DISCOUNT_TIERS) else None
        table_premium = paid_premium / (1 - current["discount"])
        if table_premium >= current["minTablePremium"] and (not next_tier or table_premium < next_tier["minTablePremium"]):
            return {
                "tablePremium": table_premium,
                "discount": current["discount"],
            }
    fallback = DISCOUNT_TIERS[0]
    return {
        "tablePremium": paid_premium / (1 - fallback["discount"]),
        "discount": fallback["discount"],
    }


def normalize_face_amount(face_amount: float) -> int:
    units = max(1, _round_int(face_amount / UNIT_FACE))
    return units * UNIT_FACE


@lru_cache(maxsize=1)
def load_tables():
    return rebuild_wus_coverage.load_tables()


@lru_cache(maxsize=200)
def build_series(gender_code: int, age: int):
    t501, t517, gp = load_tables()
    return rebuild_wus_coverage.build_series(gender_code, age, t501, t517, gp)


@lru_cache(maxsize=200)
def rate_for(gender_code: int, age: int) -> float:
    _, _, gp = load_tables()
    return rebuild_wus_coverage.gp_rate(gender_code, age, gp)


def quote_wus(request: QuoteRequest, engine_name: str = "rebuild_wus_coverage") -> QuoteResponse:
    gender_code = 1 if request.gender == Gender.male else 2
    rate = rate_for(gender_code, request.age)
    notes: list[str] = []

    if request.faceAmount:
        face_amount = normalize_face_amount(request.faceAmount)
        if face_amount != int(request.faceAmount):
            notes.append(f"Face amount normalized to nearest {UNIT_FACE:,} USD.")
        table_premium = face_amount / UNIT_FACE * rate
        paid_info = resolve_paid_amount_from_table_premium(table_premium)
    else:
        premium_info = resolve_table_premium_from_paid_amount(request.premium or 0)
        raw_units = premium_info["tablePremium"] / rate
        rounded_units = max(1, _round_int(raw_units))
        face_amount = rounded_units * UNIT_FACE
        paid_info = resolve_paid_amount_from_table_premium(rounded_units * rate)
        if paid_info["paidPremium"] != int(request.premium or 0):
            notes.append("Premium normalized to workbook-supported face amount units.")

    series = build_series(gender_code, request.age)
    if request.coverageYear > len(series):
        raise ValueError(f"Coverage year {request.coverageYear} exceeds supported policy year {len(series)} for age {request.age}.")
    base_benefit = series[request.coverageYear - 1]
    projected_benefit = _round_int(base_benefit * face_amount / BASE_FACE_AMOUNT)
    premium = paid_info["paidPremium"]
    financing_amount = _round_int(premium * request.ltvRatio)
    self_fund_amount = premium - financing_amount
    coverage_before = round(projected_benefit / premium, 4) if premium else 0.0
    coverage_after = round(projected_benefit / self_fund_amount, 4) if self_fund_amount else 0.0

    return QuoteResponse(
        productCode=request.productCode,
        financingType=request.financingType.value,
        engine=engine_name,
        sourceWorkbook=SOURCE_WORKBOOK.name,
        dividendOption=request.dividendOption.value,
        gender=request.gender.value,
        age=request.age,
        coverageYear=request.coverageYear,
        faceAmount=face_amount,
        premium=premium,
        tablePremium=paid_info["tablePremium"],
        projectedBenefit=projected_benefit,
        coverageBefore=coverage_before,
        coverageAfter=coverage_after,
        financingAmount=financing_amount,
        selfFundAmount=self_fund_amount,
        currency="USD",
        notes=notes,
    )
