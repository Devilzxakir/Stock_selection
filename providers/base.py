"""
Base classes and interfaces for financial data providers.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple
import logging
import time
import threading

logger = logging.getLogger(__name__)


class ProviderName(Enum):
    YAHOO_FINANCE = "yahoo_finance"
    ALPHA_VANTAGE = "alpha_vantage"
    FINANCIAL_MODELING_PREP = "financial_modeling_prep"


class DataMode(Enum):
    SMART = "smart"
    PRIMARY_ONLY = "primary_only"
    BOTH = "both"


@dataclass
class DataPoint:
    """A single data point with provenance tracking."""
    value: Any
    provider: ProviderName
    retrieved_at: datetime = field(default_factory=datetime.now)
    reporting_date: Optional[datetime] = None
    currency: str = "USD"
    unit: str = "raw"
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "value": self.value,
            "provider": self.provider.value,
            "retrieved_at": self.retrieved_at.isoformat(),
            "reporting_date": self.reporting_date.isoformat() if self.reporting_date else None,
            "currency": self.currency,
            "unit": self.unit,
            "metadata": self.metadata,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DataPoint":
        return cls(
            value=data["value"],
            provider=ProviderName(data["provider"]),
            retrieved_at=datetime.fromisoformat(data["retrieved_at"]),
            reporting_date=datetime.fromisoformat(data["reporting_date"]) if data.get("reporting_date") else None,
            currency=data.get("currency", "USD"),
            unit=data.get("unit", "raw"),
            metadata=data.get("metadata", {}),
        )


@dataclass
class TimeSeriesPoint:
    """A time series data point with period info."""
    period: str
    date: datetime
    value: Optional[float]
    provider: ProviderName
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "period": self.period,
            "date": self.date.isoformat(),
            "value": self.value,
            "provider": self.provider.value,
            "metadata": self.metadata,
        }


@dataclass
class FinancialStatement:
    """Normalized financial statement data."""
    symbol: str
    statement_type: str
    frequency: str
    data: Dict[str, List[TimeSeriesPoint]]
    provider: ProviderName
    retrieved_at: datetime = field(default_factory=datetime.now)
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class QuoteData:
    """Normalized quote/price data."""
    symbol: str
    price: Optional[float]
    change: Optional[float]
    change_percent: Optional[float]
    volume: Optional[int]
    market_cap: Optional[float]
    timestamp: datetime
    provider: ProviderName
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class CompanyProfile:
    """Normalized company profile data."""
    symbol: str
    name: str
    sector: Optional[str]
    industry: Optional[str]
    description: Optional[str]
    market_cap: Optional[float]
    shares_outstanding: Optional[float]
    provider: ProviderName
    retrieved_at: datetime = field(default_factory=datetime.now)
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class HistoricalPEPoint:
    """Historical P/E ratio point with full provenance."""
    quarter: str
    date: datetime
    pe: Optional[float]
    eps: Optional[float]
    price: Optional[float]
    provider: ProviderName
    ttm_eps: Optional[float] = None
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "quarter": self.quarter,
            "date": self.date.isoformat(),
            "pe": self.pe,
            "eps": self.eps,
            "price": self.price,
            "provider": self.provider.value,
            "ttm_eps": self.ttm_eps,
            "metadata": self.metadata,
        }


@dataclass
class ProviderResult:
    """Result from a provider request."""
    success: bool
    data: Any = None
    error: Optional[str] = None
    rate_limited: bool = False
    retry_after: Optional[int] = None
    provider: Optional[ProviderName] = None
    request_duration_ms: int = 0


class RateLimiter:
    """Thread-safe rate limiter with token bucket algorithm."""

    def __init__(self, requests_per_minute: int, requests_per_day: Optional[int] = None):
        self.requests_per_minute = requests_per_minute
        self.requests_per_day = requests_per_day
        self._minute_bucket: List[float] = []
        self._day_bucket: List[float] = []
        self._lock = threading.Lock()

    def acquire(self, timeout: float = 30.0) -> bool:
        """Wait until a request slot is available. Returns True if acquired, False if timeout."""
        start = time.time()
        while time.time() - start < timeout:
            with self._lock:
                now = time.time()
                # Clean old entries
                self._minute_bucket = [t for t in self._minute_bucket if now - t < 60]
                if self.requests_per_day:
                    self._day_bucket = [t for t in self._day_bucket if now - t < 86400]

                minute_ok = len(self._minute_bucket) < self.requests_per_minute
                day_ok = not self.requests_per_day or len(self._day_bucket) < self.requests_per_day

                if minute_ok and day_ok:
                    self._minute_bucket.append(now)
                    if self.requests_per_day:
                        self._day_bucket.append(now)
                    return True

            time.sleep(0.5)
        return False

    def get_wait_time(self) -> float:
        """Get estimated wait time until next slot."""
        with self._lock:
            now = time.time()
            self._minute_bucket = [t for t in self._minute_bucket if now - t < 60]
            if not self._minute_bucket:
                return 0
            oldest = min(self._minute_bucket)
            return max(0, 60 - (now - oldest))


class BaseProvider(ABC):
    """Abstract base class for financial data providers."""

    def __init__(
        self,
        api_key: str,
        rate_limit_rpm: int = 60,
        rate_limit_rpd: Optional[int] = None,
        timeout: int = 10,
        max_retries: int = 3,
        backoff_factor: float = 1.5,
    ):
        self.api_key = api_key
        self.timeout = timeout
        self.max_retries = max_retries
        self.backoff_factor = backoff_factor
        self.rate_limiter = RateLimiter(rate_limit_rpm, rate_limit_rpd)
        self._cache: Dict[str, Tuple[Any, datetime]] = {}
        self._cache_ttl: Dict[str, timedelta] = {}
        self._lock = threading.Lock()

    @property
    @abstractmethod
    def name(self) -> ProviderName:
        pass

    @abstractmethod
    def _make_request(self, endpoint: str, params: Dict[str, Any]) -> ProviderResult:
        pass

    def _get_cache_key(self, endpoint: str, params: Dict[str, Any]) -> str:
        import json
        return f"{endpoint}:{json.dumps(params, sort_keys=True)}"

    def _get_cached(self, key: str) -> Optional[Any]:
        with self._lock:
            if key in self._cache:
                data, ts = self._cache[key]
                ttl = self._cache_ttl.get(key, timedelta(minutes=15))
                if datetime.now() - ts < ttl:
                    return data
                else:
                    del self._cache[key]
                    if key in self._cache_ttl:
                        del self._cache_ttl[key]
        return None

    def _set_cache(self, key: str, data: Any, ttl: timedelta = None):
        with self._lock:
            self._cache[key] = (data, datetime.now())
            self._cache_ttl[key] = ttl or timedelta(minutes=15)

    def _request_with_retry(
        self,
        endpoint: str,
        params: Dict[str, Any],
        cache_ttl: timedelta = None,
        use_cache: bool = True,
    ) -> ProviderResult:
        cache_key = self._get_cache_key(endpoint, params)

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
            start = time.time()
            result = self._make_request(endpoint, params)
            result.request_duration_ms = int((time.time() - start) * 1000)
            result.provider = self.name

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

        return ProviderResult(
            success=False,
            error=f"Max retries exceeded: {last_error}",
            provider=self.name,
        )

    @abstractmethod
    def get_quote(self, symbol: str) -> ProviderResult:
        pass

    @abstractmethod
    def get_company_profile(self, symbol: str) -> ProviderResult:
        pass

    @abstractmethod
    def get_income_statement(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        pass

    @abstractmethod
    def get_balance_sheet(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        pass

    @abstractmethod
    def get_cash_flow(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        pass

    @abstractmethod
    def get_historical_pe(self, symbol: str, max_quarters: int = 20) -> ProviderResult:
        pass

    @abstractmethod
    def get_key_metrics(self, symbol: str) -> ProviderResult:
        pass

    def normalize_symbol(self, symbol: str) -> str:
        return symbol.upper().strip()

    def clear_cache(self):
        with self._lock:
            self._cache.clear()
            self._cache_ttl.clear()


class ProviderManager:
    """Manages multiple providers with different modes."""

    def __init__(
        self,
        primary: BaseProvider,
        secondary: Optional[BaseProvider] = None,
        mode: DataMode = DataMode.SMART,
    ):
        self.primary = primary
        self.secondary = secondary
        self.mode = mode

    def set_mode(self, mode: DataMode):
        self.mode = mode

    def _is_data_sufficient(self, data: Any, required_fields: List[str]) -> Tuple[bool, List[str]]:
        """Check if data has all required fields."""
        if data is None:
            return False, required_fields
        missing = []
        for field in required_fields:
            if not self._has_field(data, field):
                missing.append(field)
        return len(missing) == 0, missing

    def _has_field(self, data: Any, field: str) -> bool:
        """Check if nested field exists and is not None."""
        parts = field.split(".")
        current = data
        for part in parts:
            if isinstance(current, dict):
                current = current.get(part)
            elif hasattr(current, part):
                current = getattr(current, part)
            else:
                return False
            if current is None:
                return False
        return True

    def _merge_data(self, primary_data: Any, secondary_data: Any, field_priority: Dict[str, ProviderName]) -> Any:
        """Merge data from two providers with conflict resolution."""
        if primary_data is None:
            return secondary_data
        if secondary_data is None:
            return primary_data

        # For dict-like data, merge recursively
        if isinstance(primary_data, dict) and isinstance(secondary_data, dict):
            merged = primary_data.copy()
            for key, value in secondary_data.items():
                if key not in merged or merged[key] is None:
                    merged[key] = value
                elif key in field_priority:
                    # Use priority provider's value
                    if field_priority[key] == self.secondary.name and self.secondary:
                        merged[key] = value
                # Otherwise keep primary
            return merged

        return primary_data

    def get_quote(self, symbol: str) -> ProviderResult:
        symbol = self.primary.normalize_symbol(symbol)

        if self.mode == DataMode.PRIMARY_ONLY:
            return self.primary.get_quote(symbol)

        primary_result = self.primary.get_quote(symbol)

        if self.mode == DataMode.BOTH and self.secondary:
            secondary_result = self.secondary.get_quote(symbol)
            if primary_result.success and secondary_result.success:
                # Compare and merge
                merged = self._merge_quotes(primary_result.data, secondary_result.data)
                return ProviderResult(success=True, data=merged, provider=self.primary.name)
            elif not primary_result.success and secondary_result.success:
                return secondary_result
            return primary_result

        # Smart mode: use primary, fallback to secondary if primary fails or data insufficient
        if primary_result.success:
            sufficient, missing = self._is_data_sufficient(
                primary_result.data, ["price", "market_cap"]
            )
            if sufficient:
                return primary_result

        if self.secondary:
            logger.info(f"Primary provider insufficient for {symbol}, trying secondary")
            return self.secondary.get_quote(symbol)

        return primary_result

    def _merge_quotes(self, primary: QuoteData, secondary: QuoteData) -> QuoteData:
        """Merge two quote data objects."""
        return QuoteData(
            symbol=primary.symbol,
            price=primary.price if primary.price is not None else secondary.price,
            change=primary.change if primary.change is not None else secondary.change,
            change_percent=primary.change_percent if primary.change_percent is not None else secondary.change_percent,
            volume=primary.volume if primary.volume is not None else secondary.volume,
            market_cap=primary.market_cap if primary.market_cap is not None else secondary.market_cap,
            timestamp=primary.timestamp,
            provider=primary.provider,
            metadata={
                **primary.metadata,
                "secondary_provider": secondary.provider.value,
                "secondary_price": secondary.price,
                "price_diff_pct": abs(primary.price - secondary.price) / primary.price * 100 if primary.price and secondary.price else None,
            },
        )

    def get_company_profile(self, symbol: str) -> ProviderResult:
        symbol = self.primary.normalize_symbol(symbol)

        if self.mode == DataMode.PRIMARY_ONLY:
            return self.primary.get_company_profile(symbol)

        primary_result = self.primary.get_company_profile(symbol)

        if self.mode == DataMode.BOTH and self.secondary:
            secondary_result = self.secondary.get_company_profile(symbol)
            if primary_result.success and secondary_result.success:
                merged = self._merge_profiles(primary_result.data, secondary_result.data)
                return ProviderResult(success=True, data=merged, provider=self.primary.name)
            elif not primary_result.success and secondary_result.success:
                return secondary_result
            return primary_result

        if primary_result.success:
            sufficient, _ = self._is_data_sufficient(
                primary_result.data, ["name", "sector", "industry"]
            )
            if sufficient:
                return primary_result

        if self.secondary:
            return self.secondary.get_company_profile(symbol)

        return primary_result

    def _merge_profiles(self, primary: CompanyProfile, secondary: CompanyProfile) -> CompanyProfile:
        return CompanyProfile(
            symbol=primary.symbol,
            name=primary.name or secondary.name,
            sector=primary.sector or secondary.sector,
            industry=primary.industry or secondary.industry,
            description=primary.description or secondary.description,
            market_cap=primary.market_cap if primary.market_cap is not None else secondary.market_cap,
            shares_outstanding=primary.shares_outstanding if primary.shares_outstanding is not None else secondary.shares_outstanding,
            provider=primary.provider,
            metadata={
                **primary.metadata,
                "secondary_provider": secondary.provider.value,
            },
        )

    def get_income_statement(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        symbol = self.primary.normalize_symbol(symbol)

        if self.mode == DataMode.PRIMARY_ONLY:
            return self.primary.get_income_statement(symbol, quarterly)

        primary_result = self.primary.get_income_statement(symbol, quarterly)

        if self.mode == DataMode.BOTH and self.secondary:
            secondary_result = self.secondary.get_income_statement(symbol, quarterly)
            if primary_result.success and secondary_result.success:
                merged = self._merge_statements(primary_result.data, secondary_result.data)
                return ProviderResult(success=True, data=merged, provider=self.primary.name)
            elif not primary_result.success and secondary_result.success:
                return secondary_result
            return primary_result

        if primary_result.success:
            sufficient, missing = self._is_data_sufficient(
                primary_result.data, ["data.Revenue", "data.Net profit"]
            )
            if sufficient:
                return primary_result

        if self.secondary:
            return self.secondary.get_income_statement(symbol, quarterly)

        return primary_result

    def _merge_statements(self, primary: FinancialStatement, secondary: FinancialStatement) -> FinancialStatement:
        merged_data = primary.data.copy()
        for metric, points in secondary.data.items():
            if metric not in merged_data or not merged_data[metric]:
                merged_data[metric] = points
            else:
                # Merge time series points, deduplicate by date
                existing_dates = {p.date for p in merged_data[metric]}
                for pt in points:
                    if pt.date not in existing_dates:
                        merged_data[metric].append(pt)

        return FinancialStatement(
            symbol=primary.symbol,
            statement_type=primary.statement_type,
            frequency=primary.frequency,
            data=merged_data,
            provider=primary.provider,
            metadata={**primary.metadata, "secondary_provider": secondary.provider.value},
        )

    def get_balance_sheet(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        symbol = self.primary.normalize_symbol(symbol)

        if self.mode == DataMode.PRIMARY_ONLY:
            return self.primary.get_balance_sheet(symbol, quarterly)

        primary_result = self.primary.get_balance_sheet(symbol, quarterly)

        if self.mode == DataMode.BOTH and self.secondary:
            secondary_result = self.secondary.get_balance_sheet(symbol, quarterly)
            if primary_result.success and secondary_result.success:
                merged = self._merge_statements(primary_result.data, secondary_result.data)
                return ProviderResult(success=True, data=merged, provider=self.primary.name)
            elif not primary_result.success and secondary_result.success:
                return secondary_result
            return primary_result

        if primary_result.success:
            sufficient, _ = self._is_data_sufficient(
                primary_result.data, ["data.Total Assets", "data.Total Debt"]
            )
            if sufficient:
                return primary_result

        if self.secondary:
            return self.secondary.get_balance_sheet(symbol, quarterly)

        return primary_result

    def get_cash_flow(self, symbol: str, quarterly: bool = True) -> ProviderResult:
        symbol = self.primary.normalize_symbol(symbol)

        if self.mode == DataMode.PRIMARY_ONLY:
            return self.primary.get_cash_flow(symbol, quarterly)

        primary_result = self.primary.get_cash_flow(symbol, quarterly)

        if self.mode == DataMode.BOTH and self.secondary:
            secondary_result = self.secondary.get_cash_flow(symbol, quarterly)
            if primary_result.success and secondary_result.success:
                merged = self._merge_statements(primary_result.data, secondary_result.data)
                return ProviderResult(success=True, data=merged, provider=self.primary.name)
            elif not primary_result.success and secondary_result.success:
                return secondary_result
            return primary_result

        if primary_result.success:
            sufficient, _ = self._is_data_sufficient(
                primary_result.data, ["data.Cash from Operating Activity"]
            )
            if sufficient:
                return primary_result

        if self.secondary:
            return self.secondary.get_cash_flow(symbol, quarterly)

        return primary_result

    def get_historical_pe(self, symbol: str, max_quarters: int = 20) -> ProviderResult:
        symbol = self.primary.normalize_symbol(symbol)

        if self.mode == DataMode.PRIMARY_ONLY:
            return self.primary.get_historical_pe(symbol, max_quarters)

        primary_result = self.primary.get_historical_pe(symbol, max_quarters)

        if self.mode == DataMode.BOTH and self.secondary:
            secondary_result = self.secondary.get_historical_pe(symbol, max_quarters)
            if primary_result.success and secondary_result.success:
                merged = self._merge_historical_pe(primary_result.data, secondary_result.data)
                return ProviderResult(success=True, data=merged, provider=self.primary.name)
            elif not primary_result.success and secondary_result.success:
                return secondary_result
            return primary_result

        # Smart mode: check if primary has enough quarters (target 16+)
        if primary_result.success and primary_result.data:
            if len(primary_result.data) >= 16:
                return primary_result

        if self.secondary:
            logger.info(f"Primary provider has insufficient historical P/E for {symbol}, trying secondary")
            secondary_result = self.secondary.get_historical_pe(symbol, max_quarters)
            if secondary_result.success:
                # Merge if primary has some data
                if primary_result.success and primary_result.data:
                    return ProviderResult(
                        success=True,
                        data=self._merge_historical_pe(primary_result.data, secondary_result.data),
                        provider=self.primary.name,
                    )
                return secondary_result

        return primary_result

    def _merge_historical_pe(
        self,
        primary: List[HistoricalPEPoint],
        secondary: List[HistoricalPEPoint],
    ) -> List[HistoricalPEPoint]:
        """Merge historical P/E data, deduplicate by quarter."""
        by_quarter = {}
        for pt in primary:
            by_quarter[pt.quarter] = pt
        for pt in secondary:
            if pt.quarter not in by_quarter:
                by_quarter[pt.quarter] = pt
            else:
                # Flag discrepancy if both have values
                existing = by_quarter[pt.quarter]
                if existing.pe is not None and pt.pe is not None:
                    diff_pct = abs(existing.pe - pt.pe) / existing.pe * 100
                    if diff_pct > 10:
                        existing.metadata["discrepancy"] = {
                            "secondary_provider": pt.provider.value,
                            "secondary_pe": pt.pe,
                            "diff_pct": round(diff_pct, 2),
                        }
        return sorted(by_quarter.values(), key=lambda x: x.date)

    def get_key_metrics(self, symbol: str) -> ProviderResult:
        symbol = self.primary.normalize_symbol(symbol)

        if self.mode == DataMode.PRIMARY_ONLY:
            return self.primary.get_key_metrics(symbol)

        primary_result = self.primary.get_key_metrics(symbol)

        if self.mode == DataMode.BOTH and self.secondary:
            secondary_result = self.secondary.get_key_metrics(symbol)
            if primary_result.success and secondary_result.success:
                merged = self._merge_key_metrics(primary_result.data, secondary_result.data)
                return ProviderResult(success=True, data=merged, provider=self.primary.name)
            elif not primary_result.success and secondary_result.success:
                return secondary_result
            return primary_result

        if primary_result.success:
            return primary_result

        if self.secondary:
            return self.secondary.get_key_metrics(symbol)

        return primary_result

    def _merge_key_metrics(self, primary: Dict, secondary: Dict) -> Dict:
        merged = primary.copy()
        for key, value in secondary.items():
            if key not in merged or merged[key] is None:
                merged[key] = value
            elif isinstance(value, (int, float)) and isinstance(merged[key], (int, float)):
                # Flag discrepancy
                if merged[key] != 0:
                    diff_pct = abs(merged[key] - value) / abs(merged[key]) * 100
                    if diff_pct > 10:
                        merged[f"{key}_discrepancy"] = {
                            "primary": merged[key],
                            "secondary": value,
                            "secondary_provider": self.secondary.name.value,
                            "diff_pct": round(diff_pct, 2),
                        }
        return merged