from dataclasses import dataclass
from pathlib import Path

from app.schemas import Gender, QuoteRequest, QuoteResponse
from app.services.workbook_runtime import (
    WorkbookCellMap,
    WorkbookRangeMap,
    create_temp_workbook_copy,
    read_range_value,
    read_scalar,
    recalc_with_libreoffice,
    write_inputs,
)


ROOT = Path(__file__).resolve().parents[3]
SOURCE_WORKBOOK = ROOT / "WUS(金控)國泰人壽鑽美滿利率變動型美元終身壽險(定期給付型)-11501版(宣告4.25%-1150101啟用)(保費融資)_i值.xlsx"


@dataclass(frozen=True)
class WusWorkbookConfig:
    input_sheet: str = "大表 (2)"
    cells: WorkbookCellMap = WorkbookCellMap(
        gender="B2",
        age="B3",
        face_amount="B6",
        dividend_option="B13",
    )
    outputs: WorkbookRangeMap = WorkbookRangeMap(
        paid_premium="大表 (2)!B35",
        table_premium="大表 (2)!B34",
        total_benefit="利益分析表_基本!R5:R109",
    )


CONFIG = WusWorkbookConfig()
DIVIDEND_OPTION_CODES = {
    "第7年起持續增購保額": 1,
    "第7年起儲存生息": 2,
    "第7年起當年度給付": 3,
}


def _round_int(value: float) -> int:
    return int(value + 0.5)


def quote_wus_from_workbook(request: QuoteRequest, engine_name: str = "libreoffice_workbook") -> QuoteResponse:
    if request.premium is not None and request.faceAmount is None:
        raise ValueError("Workbook engine currently requires faceAmount input for WUS.")

    gender_code = 1 if request.gender == Gender.male else 2
    face_amount = int(request.faceAmount or 0)
    if face_amount <= 0:
        raise ValueError("faceAmount must be greater than 0.")

    dividend_option_code = DIVIDEND_OPTION_CODES[request.dividendOption.value]
    temp_workbook = create_temp_workbook_copy(SOURCE_WORKBOOK)
    write_inputs(
        temp_workbook,
        CONFIG.input_sheet,
        CONFIG.cells,
        gender_code=gender_code,
        age=request.age,
        face_amount=face_amount,
        dividend_option_code=dividend_option_code,
    )
    recalculated = recalc_with_libreoffice(temp_workbook)

    premium = _round_int(read_scalar(recalculated, CONFIG.outputs.paid_premium))
    table_premium = _round_int(read_scalar(recalculated, CONFIG.outputs.table_premium))
    projected_benefit = _round_int(read_range_value(recalculated, CONFIG.outputs.total_benefit, request.coverageYear - 1))
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
        tablePremium=table_premium,
        projectedBenefit=projected_benefit,
        coverageBefore=coverage_before,
        coverageAfter=coverage_after,
        financingAmount=financing_amount,
        selfFundAmount=self_fund_amount,
        currency="USD",
        notes=[],
    )
