from enum import Enum
from typing import Literal, Optional

from pydantic import BaseModel, Field, model_validator


class Gender(str, Enum):
    male = "male"
    female = "female"


class FinancingType(str, Enum):
    premium = "premium"


class DividendOption(str, Enum):
    buyup_from_year_7 = "第7年起持續增購保額"


class QuoteRequest(BaseModel):
    productCode: Literal["WUS"] = "WUS"
    financingType: FinancingType = FinancingType.premium
    gender: Gender = Gender.male
    age: int = Field(..., ge=0, le=75)
    coverageYear: int = Field(20, ge=1, le=100)
    faceAmount: Optional[float] = Field(default=None, gt=0)
    premium: Optional[float] = Field(default=None, gt=0)
    ltvRatio: float = Field(0.5, gt=0, lt=1)
    dividendOption: DividendOption = DividendOption.buyup_from_year_7

    @model_validator(mode="after")
    def validate_amounts(self):
        if self.faceAmount is None and self.premium is None:
            raise ValueError("Either faceAmount or premium is required.")
        return self


class QuoteResponse(BaseModel):
    productCode: str
    financingType: str
    engine: str
    sourceWorkbook: str
    dividendOption: str
    gender: str
    age: int
    coverageYear: int
    faceAmount: int
    premium: int
    tablePremium: int
    projectedBenefit: int
    coverageBefore: float
    coverageAfter: float
    financingAmount: int
    selfFundAmount: int
    currency: str
    notes: list[str] = Field(default_factory=list)


class HealthResponse(BaseModel):
    status: str
    workbookEngine: str
    supportedProducts: list[str]
