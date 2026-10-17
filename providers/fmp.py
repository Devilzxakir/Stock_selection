"""
Financial Modeling Prep (FMP) provider implementation.
Free tier: 250 requests/day
Starter: 1,000 requests/day
Professional: 5,000 requests/day
Enterprise: Custom

Key advantages:
- Direct historical P/E endpoint
- More generous free tier than Alpha Vantage
- Better quarterly coverage
- Direct earnings calendar with quarterly EPS
"""

import requests
import logging
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

from .base import (
    BaseProvider,
    ProviderName,
    ProviderResult,
    QuoteData,
    CompanyProfile,
    FinancialStatement,
    TimeSeriesPoint,
    HistoricalPEPoint,
)

logger = logging.getLogger(__name__)


class FMPProvider(BaseProvider):
    """Financial Modeling Prep API provider."""

    BASE_URL = "https://financialmodelingprep.com/api/v3"

    # Free tier limits
    DEFAULT_RPM = 30
    DEFAULT_RPD = 250

    def __init__(
        self,
        api_key: str,
        plan: str = "free",  # free, starter, professional, enterprise
        timeout: int = 15,
        max_retries: int = 3,
    ):
        rpm_limits = {"free": 30, "starter": 60, "professional": 120, "enterprise": 300}
        rpd_limits = {"free": 250, "starter": 1000, "professional": 5000, "enterprise": 50000}

        rpm = rpm_limits.get(plan.lower(), self.DEFAULT_RPM)
        rpd = rpd_limits.get(plan.lower(), self.DEFAULT_RPD)

        super().__init__(
            api_key=api_key,
            rate_limit_rpm=rpm,
            rate_limit_rpd=rpd,
            timeout=timeout,
            max_retries=max_retries,
        )
        self.plan = plan.lower()
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": "StockSense/1.0"})

    @property
    def name(self) -> ProviderName:
        return ProviderName.FINANCIAL_MODELING_PREP

    def _make_request(self, endpoint: str, params: Dict[str, Any]) -> ProviderResult:
        params["apikey"] = self.api_key
        url = f"{self.BASE_URL}/{endpoint}"
        try:
            response = self.session.get(url, params=params, timeout=self.timeout)
            response.raise_for_status()
            data = response.json()

            # FMP returns list or dict with error
            if isinstance(data, dict) and "Error Message" in data:
                return ProviderResult(success=False, error=data["Error Message"], provider=self.name)
            if isinstance(data, list) and data and isinstance(data[0], dict) and "Error Message" in data[0]:
                return ProviderResult(success=False, error=data[0]["Error Message"], provider=self.name)

            # Check for rate limit
            if isinstance(data, dict) and "Note" in data and "limit" in str(data["Note"]).lower():
                return ProviderResult(
                    success=False,
                    error=data["Note"],
                    rate_limited=True,
                    retry_after=60,
                    provider=self.name,
                )

            return ProviderResult(success=True, data=data, provider=self.name)

        except requests.exceptions.Timeout:
            return ProviderResult(success=False, error="Request timeout", provider=self.name)
        except requests.exceptions.RequestException as e:
            return ProviderResult(success=False, error=str(e), provider=self.name)
        except ValueError as e:
            return ProviderResult(success=False, error=f"Invalid JSON: {e}", provider=self.name)

    def get_quote(self, symbol: str) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)
        result = self._request_with_retry(
            "quote",
            {"symbol": symbol},
            cache_ttl=timedelta(minutes=2),
        )

        if not result.success:
            return result

        data = result.data
        if not data or not isinstance(data, list) or len(data) == 0:
            return ProviderResult(success=False, error="No quote data", provider=self.name)

        q = data[0]
        try:
            quote = QuoteData(
                symbol=symbol,
                price=self._safe_float(q.get("price")),
                change=self._safe_float(q.get("change")),
                change_percent=self._safe_float(q.get("changesPercentage")),
                volume=self._safe_int(q.get("volume")),
                market_cap=self._safe_float(q.get("marketCap")),
                timestamp=datetime.now(),
                provider=self.name,
                metadata={"raw": q},
            )
            return ProviderResult(success=True, data=quote, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Parse error: {e}", provider=self.name)

    def get_company_profile(self, symbol: str) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)
        result = self._request_with_retry(
            f"profile/{symbol}",
            {},
            cache_ttl=timedelta(hours=24),
        )

        if not result.success:
            return result

        data = result.data
        if not data or not isinstance(data, list) or len(data) == 0:
            return ProviderResult(success=False, error="No profile data", provider=self.name)

        p = data[0]
        try:
            profile = CompanyProfile(
                symbol=symbol,
                name=p.get("companyName", symbol),
                sector=p.get("sector"),
                industry=p.get("industry"),
                description=p.get("description"),
                market_cap=self._safe_float(p.get("mktCap")),
                shares_outstanding=self._safe_float(p.get("sharesOutstanding")),
                provider=self.name,
                metadata={"raw": p},
            )
            return ProviderResult(success=True, data=profile, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Parse error: {e}", provider=self.name)

    def get_income_statement(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)
        period = "quarter" if quarterly else "annual"
        endpoint = f"income-statement/{symbol}"
        params = {"period": period, "limit": 40 if quarterly else 20}

        result = self._request_with_retry(
            f"income_{period}",
            params,
            cache_ttl=timedelta(hours=6),
        )

        if not result.success:
            return result

        data = result.data
        if not data or not isinstance(data, list):
            return ProviderResult(success=False, error="No income statement data", provider=self.name)

        try:
            statement = self._parse_income_statement(symbol, data, quarterly)
            return ProviderResult(success=True, data=statement, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Parse error: {e}", provider=self.name)

    def _parse_income_statement(
        self, symbol: str, reports: List[Dict], quarterly: bool
    ) -> FinancialStatement:
        data = {}
        for report in reports:
            date_str = report.get("date") or report.get("calendarYear")
            if not date_str:
                continue
            try:
                if quarterly:
                    # FMP quarterly dates are like "2024-03-31"
                    date = datetime.strptime(date_str, "%Y-%m-%d")
                else:
                    date = datetime(int(date_str), 12, 31)
            except (ValueError, TypeError):
                continue

            period = f"Q{(date.month-1)//3+1} {date.year}" if quarterly else str(date.year)

            field_map = {
                "revenue": "Revenue",
                "grossProfit": "Gross Profit",
                "operatingIncome": "Operating Income",
                "netIncome": "Net Income",
                "ebitda": "EBITDA",
                "interestExpense": "Interest Expense",
                "depreciationAndAmortization": "Depreciation & Amortization",
                "costOfRevenue": "Cost of Revenue",
                "incomeTaxExpense": "Income Tax Expense",
            }

            for fmp_field, our_field in field_map.items():
                value = self._safe_float(report.get(fmp_field))
                if value is not None:
                    if our_field not in data:
                        data[our_field] = []
                    data[our_field].append(TimeSeriesPoint(
                        period=period,
                        date=date,
                        value=value,
                        provider=self.name,
                        metadata={"fiscal_date": date_str, "reported_currency": report.get("reportedCurrency", "USD")},
                    ))

        return FinancialStatement(
            symbol=symbol,
            statement_type="income",
            frequency="quarterly" if quarterly else "annual",
            data=data,
            provider=self.name,
        )

    def get_balance_sheet(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)
        period = "quarter" if quarterly else "annual"
        endpoint = f"balance-sheet-statement/{symbol}"
        params = {"period": period, "limit": 40 if quarterly else 20}

        result = self._request_with_retry(
            f"balance_{period}",
            params,
            cache_ttl=timedelta(hours=6),
        )

        if not result.success:
            return result

        data = result.data
        if not data or not isinstance(data, list):
            return ProviderResult(success=False, error="No balance sheet data", provider=self.name)

        try:
            statement = self._parse_balance_sheet(symbol, data, quarterly)
            return ProviderResult(success=True, data=statement, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Parse error: {e}", provider=self.name)

    def _parse_balance_sheet(
        self, symbol: str, reports: List[Dict], quarterly: bool
    ) -> FinancialStatement:
        data = {}
        for report in reports:
            date_str = report.get("date") or report.get("calendarYear")
            if not date_str:
                continue
            try:
                if quarterly:
                    date = datetime.strptime(date_str, "%Y-%m-%d")
                else:
                    date = datetime(int(date_str), 12, 31)
            except (ValueError, TypeError):
                continue

            period = f"Q{(date.month-1)//3+1} {date.year}" if quarterly else str(date.year)

            field_map = {
                "totalAssets": "Total Assets",
                "totalLiabilities": "Total Liabilities",
                "totalStockholdersEquity": "Total Equity",
                "cashAndCashEquivalents": "Cash & Equivalents",
                "shortTermInvestments": "Short-term Investments",
                "inventory": "Inventory",
                "netReceivables": "Receivables",
                "totalCurrentAssets": "Current Assets",
                "totalCurrentLiabilities": "Current Liabilities",
                "longTermDebt": "Long-term Debt",
                "shortTermDebt": "Short-term Debt",
                "commonStock": "Common Stock",
                "retainedEarnings": "Retained Earnings",
            }

            for fmp_field, our_field in field_map.items():
                value = self._safe_float(report.get(fmp_field))
                if value is not None:
                    if our_field not in data:
                        data[our_field] = []
                    data[our_field].append(TimeSeriesPoint(
                        period=period,
                        date=date,
                        value=value,
                        provider=self.name,
                        metadata={"fiscal_date": date_str},
                    ))

        return FinancialStatement(
            symbol=symbol,
            statement_type="balance_sheet",
            frequency="quarterly" if quarterly else "annual",
            data=data,
            provider=self.name,
        )

    def get_cash_flow(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)
        period = "quarter" if quarterly else "annual"
        params = {"period": period, "limit": 40 if quarterly else 20}

        result = self._request_with_retry(
            f"cashflow_{period}",
            params,
            cache_ttl=timedelta(hours=6),
        )

        if not result.success:
            return result

        data = result.data
        if not data or not isinstance(data, list):
            return ProviderResult(success=False, error="No cash flow data", provider=self.name)

        try:
            statement = self._parse_cash_flow(symbol, data, quarterly)
            return ProviderResult(success=True, data=statement, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Parse error: {e}", provider=self.name)

    def _parse_cash_flow(
        self, symbol: str, reports: List[Dict], quarterly: bool
    ) -> FinancialStatement:
        data = {}
        for report in reports:
            date_str = report.get("date") or report.get("calendarYear")
            if not date_str:
                continue
            try:
                if quarterly:
                    date = datetime.strptime(date_str, "%Y-%m-%d")
                else:
                    date = datetime(int(date_str), 12, 31)
            except (ValueError, TypeError):
                continue

            period = f"Q{(date.month-1)//3+1} {date.year}" if quarterly else str(date.year)

            field_map = {
                "operatingCashFlow": "Cash from Operating Activity",
                "capitalExpenditure": "Capital Expenditure",
                "freeCashFlow": "Free Cash Flow",
                "dividendsPaid": "Dividends Paid",
                "netChangeInCash": "Change in Cash",
            }

            for fmp_field, our_field in field_map.items():
                value = self._safe_float(report.get(fmp_field))
                if value is not None:
                    if our_field not in data:
                        data[our_field] = []
                    data[our_field].append(TimeSeriesPoint(
                        period=period,
                        date=date,
                        value=value,
                        provider=self.name,
                        metadata={"fiscal_date": date_str},
                    ))

        return FinancialStatement(
            symbol=symbol,
            statement_type="cash_flow",
            frequency="quarterly" if quarterly else "annual",
            data=data,
            provider=self.name,
        )

    def get_historical_pe(self, symbol: str, max_quarters: int = 20) -> ProviderResult:
        """FMP provides direct historical P/E endpoint!"""
        symbol = self.normalize_symbol(symbol)
        result = self._request_with_retry(
            f"historical_pe/{symbol}",
            {"limit": max_quarters},
            cache_ttl=timedelta(hours=6),
        )

        if not result.success:
            return result

        data = result.data
        if not data or not isinstance(data, list):
            return ProviderResult(success=False, error="No historical P/E data", provider=self.name)

        try:
            pe_points = []
            for item in data[:max_quarters]:
                date_str = item.get("date")
                if not date_str:
                    continue
                try:
                    date = datetime.strptime(date_str, "%Y-%m-%d")
                except ValueError:
                    continue

                pe = self._safe_float(item.get("pe"))
                if pe is None:
                    continue

                quarter = f"Q{(date.month-1)//3+1} {date.year}"
                pe_points.append(HistoricalPEPoint(
                    quarter=quarter,
                    date=date,
                    pe=pe,
                    eps=self._safe_float(item.get("eps")),
                    price=self._safe_float(item.get("price")),
                    provider=self.name,
                    ttm_eps=self._safe_float(item.get("eps")),
                    metadata={"source": "fmp_historical_pe", "raw": item},
                ))

            # Sort by date ascending
            pe_points.sort(key=lambda x: x.date)
            return ProviderResult(success=True, data=pe_points, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Parse error: {e}", provider=self.name)

    def get_key_metrics(self, symbol: str) -> ProviderResult:
        """Get key metrics from ratios and key-metrics endpoints."""
        symbol = self.normalize_symbol(symbol)

        # Get key metrics (TTM)
        km_result = self._request_with_retry(
            f"key-metrics/{symbol}",
            {"period": "annual", "limit": 1},
            cache_ttl=timedelta(hours=6),
        )

        # Get ratios (TTM)
        ratios_result = self._request_with_retry(
            f"ratios/{symbol}",
            {"period": "annual", "limit": 1},
            cache_ttl=timedelta(hours=6),
        )

        metrics = {}
        if km_result.success and km_result.data:
            km = km_result.data[0] if isinstance(km_result.data, list) else km_result.data
            metrics.update({
                "pe_ratio": self._safe_float(km.get("peRatio")),
                "peg_ratio": self._safe_float(km.get("pegRatio")),
                "price_to_book": self._safe_float(km.get("pbRatio")),
                "price_to_sales": self._safe_float(km.get("psRatio")),
                "ev_to_ebitda": self._safe_float(km.get("evToEBITDA")),
                "ev_to_revenue": self._safe_float(km.get("evToRevenue")),
                "profit_margin": self._safe_float(km.get("netProfitMargin")),
                "operating_margin": self._safe_float(km.get("operatingMargin")),
                "return_on_equity": self._safe_float(km.get("roe")),
                "return_on_assets": self._safe_float(km.get("roa")),
                "debt_to_equity": self._safe_float(km.get("debtToEquity")),
                "current_ratio": self._safe_float(km.get("currentRatio")),
                "quick_ratio": self._safe_float(km.get("quickRatio")),
                "beta": self._safe_float(km.get("beta")),
                "dividend_yield": self._safe_float(km.get("dividendYield")),
                "payout_ratio": self._safe_float(km.get("payoutRatio")),
                "earnings_growth": self._safe_float(km.get("earningsGrowth")),
                "revenue_growth": self._safe_float(km.get("revenueGrowth")),
                "book_value_per_share": self._safe_float(km.get("bookValuePerShare")),
                "eps_ttm": self._safe_float(km.get("epsDiluted")),
                "revenue_ttm": self._safe_float(km.get("revenuePerShare")),
                "free_cash_flow_per_share": self._safe_float(km.get("freeCashFlowPerShare")),
            })

        if ratios_result.success and ratios_result.data:
            ratios = ratios_result.data[0] if isinstance(ratios_result.data, list) else ratios_result.data
            metrics.update({
                "gross_margin": self._safe_float(ratios.get("grossProfitMargin")),
                "net_margin": self._safe_float(ratios.get("netProfitMargin")),
                "roic": self._safe_float(ratios.get("roic")),
                "interest_coverage": self._safe_float(ratios.get("interestCoverage")),
                "asset_turnover": self._safe_float(ratios.get("assetTurnover")),
                "inventory_turnover": self._safe_float(ratios.get("inventoryTurnover")),
                "receivables_turnover": self._safe_float(ratios.get("receivablesTurnover")),
            })

        return ProviderResult(success=True, data=metrics, provider=self.name)

    def get_earnings_calendar(self, symbol: str) -> ProviderResult:
        """Get historical earnings with quarterly EPS for P/E calculation."""
        symbol = self.normalize_symbol(symbol)
        result = self._request_with_retry(
            f"historical/earning_calendar/{symbol}",
            {},
            cache_ttl=timedelta(hours=6),
        )

        if not result.success:
            return result

        data = result.data
        if not data or not isinstance(data, list):
            return ProviderResult(success=False, error="No earnings calendar", provider=self.name)

        return ProviderResult(success=True, data=data, provider=self.name)

    @staticmethod
    def _safe_float(value: Any) -> Optional[float]:
        if value is None or value == "" or value == "None":
            return None
        try:
            return float(value)
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _safe_int(value: Any) -> Optional[int]:
        if value is None or value == "" or value == "None":
            return None
        try:
            return int(float(value))
        except (ValueError, TypeError):
            return None