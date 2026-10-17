"""
Yahoo Finance (yfinance) provider implementation.
Wraps existing yfinance usage into the new provider interface.
"""

import yfinance as yf
import pandas as pd
import logging
import time
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


class YahooFinanceProvider(BaseProvider):
    """Yahoo Finance provider using yfinance library."""

    # yfinance has no official rate limits but we should be respectful
    DEFAULT_RPM = 60
    DEFAULT_RPD = 10000

    def __init__(
        self,
        api_key: str = "",  # Not used for yfinance
        timeout: int = 15,
        max_retries: int = 3,
    ):
        super().__init__(
            api_key=api_key,
            rate_limit_rpm=self.DEFAULT_RPM,
            rate_limit_rpd=self.DEFAULT_RPD,
            timeout=timeout,
            max_retries=max_retries,
        )

    @property
    def name(self) -> ProviderName:
        return ProviderName.YAHOO_FINANCE

    def _make_request(self, endpoint: str, params: Dict[str, Any]) -> ProviderResult:
        # Not used directly - we use yfinance methods
        return ProviderResult(success=False, error="Use specific methods", provider=self.name)

    def get_quote(self, symbol: str) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)

        def _fetch():
            tk = yf.Ticker(symbol)
            info = tk.info or {}
            if not info or info.get("trailingPegRatio") is None and info.get("shortName") is None:
                return ProviderResult(success=False, error="No data from Yahoo Finance", provider=self.name)

            quote = QuoteData(
                symbol=symbol,
                price=info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose"),
                change=info.get("regularMarketChange"),
                change_percent=info.get("regularMarketChangePercent"),
                volume=info.get("volume") or info.get("regularMarketVolume"),
                market_cap=info.get("marketCap"),
                timestamp=datetime.now(),
                provider=self.name,
                metadata={"raw_info": info},
            )
            return ProviderResult(success=True, data=quote, provider=self.name)

        return self._request_with_retry(f"quote_{symbol}", {"symbol": symbol}, cache_ttl=timedelta(minutes=2), use_cache=True, _custom_fetch=_fetch)

    def get_company_profile(self, symbol: str) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)

        def _fetch():
            tk = yf.Ticker(symbol)
            info = tk.info or {}
            if not info:
                return ProviderResult(success=False, error="No profile data", provider=self.name)

            profile = CompanyProfile(
                symbol=symbol,
                name=info.get("longName") or info.get("shortName") or symbol,
                sector=info.get("sector"),
                industry=info.get("industry"),
                description=info.get("longBusinessSummary"),
                market_cap=info.get("marketCap"),
                shares_outstanding=info.get("sharesOutstanding"),
                provider=self.name,
                metadata={"raw_info": info},
            )
            return ProviderResult(success=True, data=profile, provider=self.name)

        return self._request_with_retry(f"profile_{symbol}", {"symbol": symbol}, cache_ttl=timedelta(hours=24), use_cache=True, _custom_fetch=_fetch)

    def get_income_statement(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)
        period = "quarterly" if quarterly else "annual"

        def _fetch():
            tk = yf.Ticker(symbol)
            df = tk.quarterly_financials if quarterly else tk.financials
            if df is None or df.empty:
                return ProviderResult(success=False, error="No income statement data", provider=self.name)

            return ProviderResult(success=True, data=self._parse_yf_statement(symbol, df, "income", quarterly), provider=self.name)

        return self._request_with_retry(f"income_{symbol}_{period}", {"symbol": symbol, "quarterly": quarterly}, cache_ttl=timedelta(hours=6), use_cache=True, _custom_fetch=_fetch)

    def get_balance_sheet(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)
        period = "quarterly" if quarterly else "annual"

        def _fetch():
            tk = yf.Ticker(symbol)
            df = tk.quarterly_balance_sheet if quarterly else tk.balance_sheet
            if df is None or df.empty:
                return ProviderResult(success=False, error="No balance sheet data", provider=self.name)

            return ProviderResult(success=True, data=self._parse_yf_statement(symbol, df, "balance_sheet", quarterly), provider=self.name)

        return self._request_with_retry(f"balance_{symbol}_{period}", {"symbol": symbol, "quarterly": quarterly}, cache_ttl=timedelta(hours=6), use_cache=True, _custom_fetch=_fetch)

    def get_cash_flow(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)
        period = "quarterly" if quarterly else "annual"

        def _fetch():
            tk = yf.Ticker(symbol)
            df = tk.quarterly_cashflow if quarterly else tk.cashflow
            if df is None or df.empty:
                return ProviderResult(success=False, error="No cash flow data", provider=self.name)

            return ProviderResult(success=True, data=self._parse_yf_statement(symbol, df, "cash_flow", quarterly), provider=self.name)

        return self._request_with_retry(f"cashflow_{symbol}_{period}", {"symbol": symbol, "quarterly": quarterly}, cache_ttl=timedelta(hours=6), use_cache=True, _custom_fetch=_fetch)

    def get_historical_pe(self, symbol: str, max_quarters: int = 20) -> ProviderResult:
        """Compute historical P/E from yfinance earnings history and price data."""
        symbol = self.normalize_symbol(symbol)

        def _fetch():
            tk = yf.Ticker(symbol)

            # Get earnings dates with reported EPS
            try:
                earnings = tk.get_earnings_dates(limit=100)
            except Exception:
                return ProviderResult(success=False, error="No earnings data", provider=self.name)

            if earnings is None or earnings.empty:
                return ProviderResult(success=False, error="No earnings data", provider=self.name)

            # Get price history
            hist = tk.history(period="5y", interval="1d")
            if hist is None or hist.empty:
                return ProviderResult(success=False, error="No price data", provider=self.name)

            pe_points = self._compute_pe_from_earnings(earnings, hist, max_quarters)
            if not pe_points:
                return ProviderResult(success=False, error="Could not compute historical P/E", provider=self.name)

            return ProviderResult(success=True, data=pe_points, provider=self.name)

        return self._request_with_retry(
            f"historical_pe_{symbol}",
            {"symbol": symbol, "max_quarters": max_quarters},
            cache_ttl=timedelta(hours=6),
            use_cache=True,
            _custom_fetch=_fetch,
        )

    def _compute_pe_from_earnings(
        self,
        earnings: pd.DataFrame,
        hist: pd.DataFrame,
        max_quarters: int,
    ) -> List[HistoricalPEPoint]:
        """Compute P/E from earnings history and price data."""
        eps_col = None
        for c in earnings.columns:
            if str(c).strip().lower() == "reported eps":
                eps_col = c
                break
        if eps_col is None:
            return []

        past = earnings.dropna(subset=[eps_col])
        if past.empty:
            return []
        past = past[past.index < pd.Timestamp.now(tz="UTC")]
        past = past[~past.index.duplicated(keep="last")].sort_index()
        past = past.tail(40)
        if len(past) < 4:
            return []

        eps_series = past[eps_col].astype(float)
        dates = list(eps_series.index)

        pe_data = []
        for i in range(3, len(dates)):
            if len(pe_data) >= max_quarters:
                break
            try:
                ttm_eps = float(eps_series.iloc[i - 3:i + 1].sum())
            except Exception:
                continue
            if ttm_eps <= 0:
                continue

            q_end = self._pe_quarter_end(dates[i])
            price = self._pe_price_near(hist, q_end)
            if price is None:
                continue

            pe_data.append(HistoricalPEPoint(
                quarter=f"Q{(q_end.month - 1) // 3 + 1} {q_end.year}",
                date=q_end,
                pe=round(price / ttm_eps, 2),
                eps=round(ttm_eps, 2),
                price=round(price, 2),
                provider=self.name,
                ttm_eps=round(ttm_eps, 2),
                metadata={"computed_from": "yfinance_earnings+price"},
            ))

        pe_data.sort(key=lambda x: x.date)
        return pe_data

    def _pe_quarter_end(self, announce_ts: pd.Timestamp) -> datetime:
        approx = announce_ts.tz_localize(None) - pd.Timedelta(days=45)
        candidates = []
        for year in (approx.year - 1, approx.year, approx.year + 1):
            for month in (3, 6, 9, 12):
                candidates.append(pd.Timestamp(year, month, 1) + pd.offsets.MonthEnd(0))
        nearest = min(candidates, key=lambda d: abs((d - approx).days))
        if abs((nearest - approx).days) <= 20:
            return nearest.normalize().to_pydatetime()
        return approx.normalize().to_pydatetime()

    def _pe_price_near(self, hist: pd.DataFrame, q_end: datetime) -> Optional[float]:
        idx = hist.index
        naive = idx.tz_localize(None) if getattr(idx, "tz", None) is not None else idx
        target = pd.Timestamp(q_end)
        pos = naive.get_indexer([target], method="nearest")[0]
        if pos < 0 or abs((naive[pos] - target).days) > 10:
            return None
        try:
            val = hist["Close"].iloc[pos]
            return float(val) if pd.notna(val) else None
        except Exception:
            return None

    def get_key_metrics(self, symbol: str) -> ProviderResult:
        symbol = self.normalize_symbol(symbol)

        def _fetch():
            tk = yf.Ticker(symbol)
            info = tk.info or {}
            if not info:
                return ProviderResult(success=False, error="No key metrics", provider=self.name)

            metrics = {
                "pe_ratio": info.get("trailingPE"),
                "forward_pe": info.get("forwardPE"),
                "peg_ratio": info.get("trailingPegRatio"),
                "price_to_book": info.get("priceToBook"),
                "price_to_sales": info.get("priceToSalesTrailing12Months"),
                "ev_to_ebitda": None,  # Not directly available
                "ev_to_revenue": None,
                "profit_margin": info.get("profitMargins"),
                "operating_margin": info.get("operatingMargins"),
                "return_on_equity": info.get("returnOnEquity"),
                "return_on_assets": None,
                "debt_to_equity": info.get("debtToEquity"),
                "current_ratio": info.get("currentRatio"),
                "quick_ratio": info.get("quickRatio"),
                "beta": info.get("beta"),
                "dividend_yield": info.get("dividendYield"),
                "payout_ratio": info.get("payoutRatio"),
                "earnings_growth": info.get("earningsGrowth"),
                "revenue_growth": info.get("revenueGrowth"),
                "book_value_per_share": info.get("bookValue"),
                "eps_ttm": info.get("trailingEps"),
                "forward_eps": info.get("forwardEps"),
                "revenue_ttm": info.get("totalRevenue"),
                "gross_profit_ttm": info.get("grossProfits"),
                "ebitda_ttm": info.get("ebitda"),
                "free_cash_flow_ttm": info.get("freeCashflow"),
                "analyst_target_price": info.get("targetMeanPrice"),
                "trailing_pe": info.get("trailingPE"),
                "held_percent_institutions": info.get("heldPercentInstitutions"),
                "held_percent_insiders": info.get("heldPercentInsiders"),
                "fifty_two_week_high": info.get("fiftyTwoWeekHigh"),
                "fifty_two_week_low": info.get("fiftyTwoWeekLow"),
                "fifty_day_avg": info.get("fiftyDayAverage"),
                "two_hundred_day_avg": info.get("twoHundredDayAverage"),
                "total_cash": info.get("totalCash"),
                "total_debt": info.get("totalDebt"),
                "operating_cashflow": info.get("operatingCashflow"),
                "return_on_assets": info.get("returnOnAssets"),
            }
            return ProviderResult(success=True, data=metrics, provider=self.name)

        return self._request_with_retry(
            f"keymetrics_{symbol}",
            {"symbol": symbol},
            cache_ttl=timedelta(hours=6),
            use_cache=True,
            _custom_fetch=_fetch,
        )

    def _parse_yf_statement(self, symbol: str, df: pd.DataFrame, stmt_type: str, quarterly: bool) -> FinancialStatement:
        data = {}
        if df is None or df.empty:
            return FinancialStatement(
                symbol=symbol,
                statement_type=stmt_type,
                frequency="quarterly" if quarterly else "annual",
                data=data,
                provider=self.name,
            )

        for idx in df.index:
            # Map yfinance index names to our standard names
            our_name = self._map_yf_metric(str(idx), stmt_type)
            if not our_name:
                continue

            series = df.loc[idx]
            if isinstance(series, pd.DataFrame):
                # Duplicate index label — take the first row
                series = series.iloc[0]

            for col, val in series.items():
                if val is None or not pd.notna(val):
                    continue
                try:
                    num = float(val)
                except (TypeError, ValueError):
                    continue

                # Each column is its own reporting period — keep them all
                if hasattr(col, "year") and hasattr(col, "month"):
                    date = datetime(int(col.year), int(col.month), int(col.day)) if hasattr(col, "day") else datetime(int(col.year), 12, 31)
                else:
                    try:
                        date = datetime.strptime(str(col), "%Y-%m-%d")
                    except ValueError:
                        try:
                            date = datetime(int(str(col)[:4]), 12, 31)
                        except ValueError:
                            date = datetime.now()

                period = f"Q{(date.month-1)//3+1} {date.year}" if quarterly else str(date.year)
                data.setdefault(our_name, []).append(TimeSeriesPoint(
                    period=period,
                    date=date,
                    value=num,
                    provider=self.name,
                    metadata={"yf_metric": str(idx)},
                ))

        # Oldest → newest so downstream [-1] lookups see the latest value
        for points in data.values():
            points.sort(key=lambda p: p.date)

        return FinancialStatement(
            symbol=symbol,
            statement_type=stmt_type,
            frequency="quarterly" if quarterly else "annual",
            data=data,
            provider=self.name,
        )

    def _map_yf_metric(self, yf_name: str, stmt_type: str) -> Optional[str]:
        name_lower = yf_name.lower()

        if stmt_type == "income":
            if any(w in name_lower for w in ["total revenue", "revenue", "sales", "operating revenue"]):
                return "Revenue"
            if any(w in name_lower for w in ["gross profit"]):
                return "Gross Profit"
            if any(w in name_lower for w in ["operating income", "ebit"]):
                return "Operating Income"
            if any(w in name_lower for w in ["net income", "net profit", "net earnings"]):
                if "discontinued" not in name_lower and "comprehensive" not in name_lower:
                    return "Net Income"
            if "ebitda" in name_lower:
                return "EBITDA"
            if "interest expense" in name_lower:
                return "Interest Expense"
            if "depreciation" in name_lower:
                return "Depreciation & Amortization"
            if "cost of revenue" in name_lower:
                return "Cost of Revenue"

        elif stmt_type == "balance_sheet":
            if "total assets" in name_lower:
                return "Total Assets"
            if "total liabilities" in name_lower:
                return "Total Liabilities"
            if "stockholders equity" in name_lower or "shareholders equity" in name_lower:
                return "Total Equity"
            if "cash and cash equivalents" in name_lower:
                return "Cash & Equivalents"
            if "short term investments" in name_lower:
                return "Short-term Investments"
            if "inventory" in name_lower:
                return "Inventory"
            if "receivables" in name_lower or "accounts receivable" in name_lower:
                return "Receivables"
            if "current assets" in name_lower and "other current assets" not in name_lower:
                return "Current Assets"
            if "current liabilities" in name_lower and "other current liabilities" not in name_lower:
                return "Current Liabilities"
            if "long term debt" in name_lower:
                return "Long-term Debt"
            if "total debt" in name_lower:
                return "Total Debt"
            if "short term debt" in name_lower or "current debt" in name_lower:
                return "Short-term Debt"
            if "ordinary shares number" in name_lower or "share issued" in name_lower:
                return "No. of Equity Shares"
            if "common stock" in name_lower:
                return "Common Stock"
            if "retained earnings" in name_lower:
                return "Retained Earnings"

        elif stmt_type == "cash_flow":
            if "operating cash flow" in name_lower or "cash from operations" in name_lower:
                return "Cash from Operating Activity"
            if "capital expenditure" in name_lower or "capex" in name_lower:
                return "Capital Expenditure"
            if "free cash flow" in name_lower:
                return "Free Cash Flow"
            if "dividends paid" in name_lower:
                return "Dividends Paid"
            if "change in cash" in name_lower:
                return "Change in Cash"

        return None

    def _request_with_retry(
        self,
        cache_key: str,
        params: Dict[str, Any],
        cache_ttl: timedelta = None,
        use_cache: bool = True,
        _custom_fetch=None,
    ) -> ProviderResult:
        """Override to support custom fetch functions."""
        if use_cache:
            cached = self._get_cached(cache_key)
            if cached is not None:
                return ProviderResult(success=True, data=cached, provider=self.name)

        if not self.rate_limiter.acquire():
            return ProviderResult(
                success=False,
                error="Rate limit exceeded",
                rate_limited=True,
                retry_after=int(self.rate_limiter.get_wait_time()) + 1,
                provider=self.name,
            )

        last_error = None
        for attempt in range(self.max_retries):
            try:
                if _custom_fetch:
                    result = _custom_fetch()
                else:
                    result = ProviderResult(success=False, error="No fetch function", provider=self.name)

                if result.success:
                    if use_cache and cache_ttl:
                        self._set_cache(cache_key, result.data, cache_ttl)
                    return result

                last_error = result.error
                if result.rate_limited:
                    wait = result.retry_after or (self.backoff_factor ** attempt)
                    logger.warning(f"{self.name.value}: Rate limited, waiting {wait}s")
                    time.sleep(wait)
                    continue

                if attempt < self.max_retries - 1:
                    wait = self.backoff_factor ** attempt
                    logger.warning(f"{self.name.value}: Request failed (attempt {attempt+1}), retrying in {wait}s: {last_error}")
                    time.sleep(wait)

            except Exception as e:
                last_error = str(e)
                if attempt < self.max_retries - 1:
                    wait = self.backoff_factor ** attempt
                    logger.warning(f"{self.name.value}: Exception (attempt {attempt+1}), retrying in {wait}s: {e}")
                    time.sleep(wait)

        return ProviderResult(
            success=False,
            error=f"Max retries exceeded: {last_error}",
            provider=self.name,
        )

    def normalize_symbol(self, symbol: str) -> str:
        return symbol.upper().strip()