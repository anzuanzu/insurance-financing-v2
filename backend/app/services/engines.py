import os
import shutil
from abc import ABC, abstractmethod

from app.schemas import QuoteRequest, QuoteResponse
from app.services import wus_quote
from app.services.wus_workbook import quote_wus_from_workbook


class QuoteEngine(ABC):
    name: str

    @abstractmethod
    def supports(self, request: QuoteRequest) -> bool:
        raise NotImplementedError

    @abstractmethod
    def quote(self, request: QuoteRequest) -> QuoteResponse:
        raise NotImplementedError


class RebuildWusEngine(QuoteEngine):
    name = "rebuild_wus_coverage"

    def supports(self, request: QuoteRequest) -> bool:
        return request.productCode == "WUS"

    def quote(self, request: QuoteRequest) -> QuoteResponse:
        return wus_quote.quote_wus(request, engine_name=self.name)


class LibreOfficeWorkbookEngine(QuoteEngine):
    name = "libreoffice_workbook"

    def supports(self, request: QuoteRequest) -> bool:
        return request.productCode == "WUS"

    def quote(self, request: QuoteRequest) -> QuoteResponse:
        return quote_wus_from_workbook(request, engine_name=self.name)


class AutoQuoteEngine(QuoteEngine):
    name = "auto"

    def __init__(self, *, has_soffice: bool):
        self.has_soffice = has_soffice
        self.rebuild_engine = RebuildWusEngine()
        self.workbook_engine = LibreOfficeWorkbookEngine() if has_soffice else None

    def supports(self, request: QuoteRequest) -> bool:
        return request.productCode == "WUS"

    def quote(self, request: QuoteRequest) -> QuoteResponse:
        # Workbook execution is preferred when the request is directly driven by
        # face amount. Premium-only reverse lookup still relies on the rebuild path.
        if self.workbook_engine is not None and request.faceAmount is not None:
            return self.workbook_engine.quote(request)
        return self.rebuild_engine.quote(request)


def get_engine() -> QuoteEngine:
    preferred = os.getenv("QUOTE_ENGINE", "auto").strip().lower()
    has_soffice = shutil.which("soffice") is not None or shutil.which("libreoffice") is not None

    if preferred == "rebuild":
        return RebuildWusEngine()
    if preferred == "libreoffice":
        if not has_soffice:
            raise RuntimeError("QUOTE_ENGINE=libreoffice but LibreOffice is not available.")
        return LibreOfficeWorkbookEngine()
    if preferred == "auto":
        return AutoQuoteEngine(has_soffice=has_soffice)
    return RebuildWusEngine()
