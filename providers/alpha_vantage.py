"""
Alpha Vantage provider implementation.
Free tier: 25 requests/day, 5 requests/minute
Premium: Higher limits based on plan
"""

import requests
import logging
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional
import time

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


class AlphaVantageProvider(BaseProvider):
    """Alpha Vantage API provider."""

    BASE_URL = "https://www.alphavantage.co/query"

    # Free tier limits
    DEFAULT_RPM = 5
    DEFAULT_RPD = 25

    def __init__(
        self,
        api_key: str,
        premium: bool = False,
        timeout: int = 15,
        max_retries: int = 3,
    ):
        # Premium plans have higher limits
        rpm = 30 if premium else self.DEFAULT_RPM
        rpd = 500 if premium else self.DEFAULT_RPD

        super().__init__(
            api_key=api_key,
            rate_limit_rpm=rpm,
            rate_limit_rpd=rpd,
            timeout=timeout,
            max_retries=max_retries,
        )
        self.premium = premium
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": "StockSense/1.0"})

    @property
    def name(self) -> ProviderName:
        return ProviderName.ALPHA_VANTAGE

    def _make_request(self, endpoint: str, params: Dict[str, Any]) -> ProviderResult:
        params["apikey"] = self.api_key
        try:
            response = self.session.get(
                self.BASE_URL,
                params=params,
                timeout=self.timeout,
            )
            response.raise_for_status()
            data = response.json()

            # Check for API errors
            if "Error Message" in data:
                return ProviderResult(success=False, error=data["Error Message"], provider=self.name)
            if "Note" in data:
                # Rate limit message
                return ProviderResult(
                    success=False,
                    error=data["Note"],
                    rate_limited=True,
                    retry_after=60,
                    provider=self.name,
                )
            if "Information" in data:
                return ProviderResult(
                    success=False,
                    error=data["Information"],
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
            {"function": "GLOBAL_QUOTE", "symbol": symbol},
            cache_ttl=timedelta(minutes=2),
        )

        if not result.success:
            return result

        data = result.data.get("Global Quote", {})
        if not data:
            return ProviderResult(success=False, error="No quote data", provider=self.name)

        try:
            quote = QuoteData(
                symbol=symbol,
                price=self._safe_float(data.get("05. price")),
                change=self._safe_float(data.get("09. change")),
                change_percent=self._safe_float(data.get("10. change percent", "").replace("%", "")),
                volume=self._safe_int(data.get("06. volume")),
                market_cap=None,  # Not in GLOBAL_QUOTE
                timestamp=datetime.now(),
                provider=self.name,
                metadata={"raw": data},
            )
            return ProviderResult(success=True, data=quote, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Parse error: {e}", provider=self.name)

    def get_company_profile(self, symbol: str) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)
        result = self._request_with_retry(
            "overview",
            {"function": "OVERVIEW", "symbol": symbol},
            cache_ttl=timedelta(hours=24),
        )

        if not result.success:
            return result

        data = result.data
        if not data or "Symbol" not in data:
            return ProviderResult(success=False, error="No overview data", provider=self.name)

        try:
            profile = CompanyProfile(
                symbol=symbol,
                name=data.get("Name", symbol),
                sector=data.get("Sector"),
                industry=data.get("Industry"),
                description=data.get("Description"),
                market_cap=self._safe_float(data.get("MarketCapitalization")),
                shares_outstanding=self._safe_float(data.get("SharesOutstanding")),
                provider=self.name,
                metadata={"raw": data},
            )
            return ProviderResult(success=True, data=profile, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Parse error: {e}", provider=self.name)

    def get_income_statement(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)
        function = "INCOME_STATEMENT" + ("_QUARTERLY" if quarterly else "")
        result = self._request_with_retry(
            f"income_{'quarterly' if quarterly else 'annual'}",
            {"function": function, "symbol": symbol},
            cache_ttl=timedelta(hours=6),
        )

        if not result.success:
            return result

        reports = result.data.get("quarterlyReports" if quarterly else "annualReports", [])
        if not reports:
            return ProviderResult(success=False, error="No income statement data", provider=self.name)

        try:
            statement = self._parse_income_statement(symbol, reports, quarterly)
            return ProviderResult(success=True, data=statement, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Parse error: {e}", provider=self.name)

    def _parse_income_statement(
        self, symbol: str, reports: List[Dict], quarterly: bool
    ) -> FinancialStatement:
        data = {}
        for report in reports:
            fiscal_date = report.get("fiscalDateEnding")
            if not fiscal_date:
                continue
            try:
                date = datetime.strptime(fiscal_date, "%Y-%m-%d")
            except ValueError:
                continue

            period = f"Q{(date.month-1)//3+1} {date.year}" if quarterly else str(date.year)

            # Map Alpha Vantage fields to our standard names
            field_map = {
                "totalRevenue": "Revenue",
                "grossProfit": "Gross Profit",
                "operatingIncome": "Operating Income",
                "netIncome": "Net Income",
                "ebitda": "EBITDA",
                "interestExpense": "Interest Expense",
                "depreciationAndAmortization": "Depreciation & Amortization",
                "costOfRevenue": "Cost of Revenue",
            }

            for av_field, our_field in field_map.items():
                value = self._safe_float(report.get(av_field))
                if value is not None:
                    if our_field not in data:
                        data[our_field] = []
                    data[our_field].append(TimeSeriesPoint(
                        period=period,
                        date=date,
                        value=value,
                        provider=self.name,
                        metadata={"fiscal_date": fiscal_date, "reported_currency": report.get("reportedCurrency", "USD")},
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
        function = "BALANCE_SHEET" + ("_QUARTERLY" if quarterly else "")
        result = self._request_with_retry(
            f"balance_{'quarterly' if quarterly else 'annual'}",
            {"function": function, "symbol": symbol},
            cache_ttl=timedelta(hours=6),
        )

        if not result.success:
            return result

        reports = result.data.get("quarterlyReports" if quarterly else "annualReports", [])
        if not reports:
            return ProviderResult(success=False, error="No balance sheet data", provider=self.name)

        try:
            statement = self._parse_balance_sheet(symbol, reports, quarterly)
            return ProviderResult(success=True, data=statement, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Parse error: {e}", provider=self.name)

    def _parse_balance_sheet(
        self, symbol: str, reports: List[Dict], quarterly: bool
    ) -> FinancialStatement:
        data = {}
        for report in reports:
            fiscal_date = report.get("fiscalDateEnding")
            if not fiscal_date:
                continue
            try:
                date = datetime.strptime(fiscal_date, "%Y-%m-%d")
            except ValueError:
                continue

            period = f"Q{(date.month-1)//3+1} {date.year}" if quarterly else str(date.year)

            field_map = {
                "totalAssets": "Total Assets",
                "totalLiabilities": "Total Liabilities",
                "totalShareholderEquity": "Total Equity",
                "cashAndCashEquivalentsAtCarryingValue": "Cash & Equivalents",
                "shortTermInvestments": "Short-term Investments",
                "inventory": "Inventory",
                "currentNetReceivables": "Receivables",
                "currentAssets": "Current Assets",
                "currentLiabilities": "Current Liabilities",
                "longTermDebt": "Long-term Debt",
                "shortTermDebt": "Short-term Debt",
                "commonStock": "Common Stock",
                "retainedEarnings": "Retained Earnings",
            }

            for av_field, our_field in field_map.items():
                value = self._safe_float(report.get(av_field))
                if value is not None:
                    if our_field not in data:
                        data[our_field] = []
                    data[our_field].append(TimeSeriesPoint(
                        period=period,
                        date=date,
                        value=value,
                        provider=self.name,
                        metadata={"fiscal_date": fiscal_date},
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
        function = "CASH_FLOW" + ("_QUARTERLY" if quarterly else "")
        result = self._request_with_retry(
            f"cashflow_{'quarterly' if quarterly else 'annual'}",
            {"function": function, "symbol": symbol},
            cache_ttl=timedelta(hours=6),
        )

        if not result.success:
            return result

        reports = result.data.get("quarterlyReports" if quarterly else "annualReports", [])
        if not reports:
            return ProviderResult(success=False, error="No cash flow data", provider=self.name)

        try:
            statement = self._parse_cash_flow(symbol, reports, quarterly)
            return ProviderResult(success=True, data=statement, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Parse error: {e}", provider=self.name)

    def _parse_cash_flow(
        self, symbol: str, reports: List[Dict], quarterly: bool
    ) -> FinancialStatement:
        data = {}
        for report in reports:
            fiscal_date = report.get("fiscalDateEnding")
            if not fiscal_date:
                continue
            try:
                date = datetime.strptime(fiscal_date, "%Y-%m-%d")
            except ValueError:
                continue

            period = f"Q{(date.month-1)//3+1} {date.year}" if quarterly else str(date.year)

            field_map = {
                "operatingCashflow": "Cash from Operating Activity",
                "capitalExpenditures": "Capital Expenditure",
                "freeCashflow": "Free Cash Flow",
                "dividendPayout": "Dividends Paid",
                "changeInCashAndCashEquivalents": "Change in Cash",
            }

            for av_field, our_field in field_map.items():
                value = self._safe_float(report.get(av_field))
                if value is not None:
                    if our_field not in data:
                        data[our_field] = []
                    data[our_field].append(TimeSeriesPoint(
                        period=period,
                        date=date,
                        value=value,
                        provider=self.name,
                        metadata={"fiscal_date": fiscal_date},
                    ))

        return FinancialStatement(
            symbol=symbol,
            statement_type="cash_flow",
            frequency="quarterly" if quarterly else "annual",
            data=data,
            provider=self.name,
        )

    def get_historical_pe(self, symbol: str, max_quarters: int = 20) -> ProviderResult:
        """Alpha Vantage doesn't provide direct historical P/E.
        We need to compute it from price history + earnings."""
        symbol = self.normalize_symbol(symbol)

        # Get quarterly earnings
        earnings_result = self._request_with_retry(
            "earnings",
            {"function": "EARNINGS", "symbol": symbol},
            cache_ttl=timedelta(hours=6),
        )

        if not earnings_result.success:
            return earnings_result

        quarterly_earnings = earnings_result.data.get("quarterlyEarnings", [])
        if not quarterly_earnings:
            return ProviderResult(success=False, error="No quarterly earnings", provider=self.name)

        # Get daily price history (compact = last 100 points)
        price_result = self._request_with_retry(
            "price_daily",
            {"function": "TIME_SERIES_DAILY_ADJUSTED", "symbol": symbol, "outputsize": "full"},
            cache_ttl=timedelta(hours=1),
        )

        if not price_result.success:
            return price_result

        price_series = price_result.data.get("Time Series (Daily)", {})
        if not price_series:
            return ProviderResult(success=False, error="No price data", provider=self.name)

        try:
            pe_points = self._compute_historical_pe(quarterly_earnings, price_series, max_quarters)
            return ProviderResult(success=True, data=pe_points, provider=self.name)
        except Exception as e:
            return ProviderResult(success=False, error=f"Compute error: {e}", provider=self.name)

    def _compute_historical_pe(
        self,
        quarterly_earnings: List[Dict],
        price_series: Dict,
        max_quarters: int,
    ) -> List[HistoricalPEPoint]:
        """Compute historical P/E from earnings and price data."""
        # Parse earnings
        earnings_by_quarter = {}
        for eq in quarterly_earnings:
            reported_date = eq.get("reportedDate")
            if not reported_date:
                continue
            try:
                eps = self._safe_float(eq.get("reportedEPS"))
                if eps is None or eps <= 0:
                    continue
                date = datetime.strptime(reported_date, "%Y-%m-%d")
                quarter = f"Q{(date.month-1)//3+1} {date.year}"
                earnings_by_quarter[quarter] = {
                    "date": date,
                    "eps": eps,
                    "reported_date": date,
                }
            except (ValueError, TypeError):
                continue

        # Sort quarters chronologically
        sorted_quarters = sorted(earnings_by_quarter.items(), key=lambda x: x[1]["date"])

        # Compute TTM EPS for each quarter (need 4 quarters)
        pe_points = []
        for i in range(3, len(sorted_quarters)):
            if len(pe_points) >= max_quarters:
                break

            current_quarter, current_data = sorted_quarters[i]
            ttm_quarters = sorted_quarters[i-3:i+1]
            ttm_eps = sum(q[1]["eps"] for q in ttm_quarters)

            # Find price near quarter end
            quarter_end = self._estimate_quarter_end(current_data["date"])
            price = self._find_price_near_date(price_series, quarter_end)

            if price and ttm_eps > 0:
                pe = price / ttm_eps
                pe_points.append(HistoricalPEPoint(
                    quarter=current_quarter,
                    date=current_data["date"],
                    pe=round(pe, 2),
                    eps=round(ttm_eps, 2),
                    price=round(price, 2),
                    provider=self.name,
                    ttm_eps=round(ttm_eps, 2),
                    metadata={"computed_from": "earnings+price"},
                ))

        return pe_points

    def _estimate_quarter_end(self, reported_date: datetime) -> datetime:
        """Estimate fiscal quarter end from earnings report date."""
        # Earnings typically reported 30-45 days after quarter end
        approx = reported_date - timedelta(days=45)
        # Snap to nearest quarter end
        for year in (approx.year - 1, approx.year, approx.year + 1):
            for month in (3, 6, 9, 12):
                quarter_end = datetime(year, month, 1) + timedelta(days=31)
                quarter_end = quarter_end.replace(day=1) - timedelta(days=1)
                if abs((quarter_end - approx).days) <= 20:
                    return quarter_end
        return approx

    def _find_price_near_date(self, price_series: Dict, target_date: datetime, max_days: int = 10) -> Optional[float]:
        """Find adjusted close price near target date."""
        target = target_date.date()
        best_price = None
        best_diff = max_days + 1

        for date_str, data in price_series.items():
            try:
                date = datetime.strptime(date_str, "%Y-%m-%d").date()
                diff = abs((date - target).days)
                if diff <= max_days and diff < best_diff:
                    best_diff = diff
                    best_price = self._safe_float(data.get("5. adjusted close"))
            except ValueError:
                continue

        return best_price

    def get_key_metrics(self, symbol: str) -> ProviderResult:
        """Get key metrics from OVERVIEW endpoint."""
        symbol = self.normalize_symbol(symbol)
        result = self._request_with_retry(
            "overview",
            {"function": "OVERVIEW", "symbol": symbol},
            cache_ttl=timedelta(hours=6),
        )

        if not result.success:
            return result

        data = result.data
        metrics = {
            "pe_ratio": self._safe_float(data.get("PERatio")),
            "peg_ratio": self._safe_float(data.get("PEGRatio")),
            "price_to_book": self._safe_float(data.get("PriceToBookRatio")),
            "price_to_sales": self._safe_float(data.get("PriceToSalesRatioTTM")),
            "ev_to_ebitda": self._safe_float(data.get("EVToEBITDA")),
            "ev_to_revenue": self._safe_float(data.get("EVToRevenue")),
            "profit_margin": self._safe_float(data.get("ProfitMargin")),
            "operating_margin": self._safe_float(data.get("OperatingMarginTTM")),
            "return_on_equity": self._safe_float(data.get("ReturnOnEquityTTM")),
            "return_on_assets": self._safe_float(data.get("ReturnOnAssetsTTM")),
            "debt_to_equity": self._safe_float(data.get("DebtToEquity")),
            "current_ratio": self._safe_float(data.get("CurrentRatio")),
            "quick_ratio": self._safe_float(data.get("QuickRatio")),
            "beta": self._safe_float(data.get("Beta")),
            "dividend_yield": self._safe_float(data.get("DividendYield")),
            "payout_ratio": self._safe_float(data.get("PayoutRatio")),
            "earnings_growth": self._safe_float(data.get("QuarterlyEarningsGrowthYOY")),
            "revenue_growth": self._safe_float(data.get("QuarterlyRevenueGrowthYOY")),
            "analyst_target_price": self._safe_float(data.get("AnalystTargetPrice")),
            "trailing_pe": self._safe_float(data.get("PERatio")),
            "forward_pe": self._safe_float(data.get("ForwardPE")),
            "book_value_per_share": self._safe_float(data.get("BookValue")),
            "eps_ttm": self._safe_float(data.get("EPS")),
            "revenue_ttm": self._safe_float(data.get("RevenueTTM")),
            "gross_profit_ttm": self._safe_float(data.get("GrossProfitTTM")),
            "ebitda_ttm": self._safe_float(data.get("EBITDA")),
            "diluted_eps_ttm": self._safe_float(data.get("DilutedEPSTTM")),
        }

        return ProviderResult(success=True, data=metrics, provider=self.name)

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