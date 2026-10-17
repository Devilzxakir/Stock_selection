"""
Financial Data Providers Package.

This package provides a unified interface for multiple financial data providers
with support for smart fallback, merging, and conflict resolution.
"""

from .base import (
    ProviderName,
    DataMode,
    DataPoint,
    TimeSeriesPoint,
    FinancialStatement,
    QuoteData,
    CompanyProfile,
    HistoricalPEPoint,
    ProviderResult,
    RateLimiter,
    BaseProvider,
    ProviderManager,
)

from .yahoo_finance import YahooFinanceProvider
from .alpha_vantage import AlphaVantageProvider
from .fmp import FMPProvider

__all__ = [
    "ProviderName",
    "DataMode",
    "DataPoint",
    "TimeSeriesPoint",
    "FinancialStatement",
    "QuoteData",
    "CompanyProfile",
    "HistoricalPEPoint",
    "ProviderResult",
    "RateLimiter",
    "BaseProvider",
    "ProviderManager",
    "YahooFinanceProvider",
    "AlphaVantageProvider",
    "FMPProvider",
]


def create_provider_manager(
    yf_api_key: str = "",
    av_api_key: str = "",
    fmp_api_key: str = "",
    mode: DataMode = DataMode.SMART,
    av_premium: bool = False,
    fmp_plan: str = "free",
) -> ProviderManager:
    """
    Factory function to create a ProviderManager with configured providers.

    Args:
        yf_api_key: Not used for Yahoo Finance (kept for interface consistency)
        av_api_key: Alpha Vantage API key
        fmp_api_key: Financial Modeling Prep API key
        mode: Data fetching mode (SMART, PRIMARY_ONLY, BOTH)
        av_premium: Whether Alpha Vantage premium plan is used
        fmp_plan: FMP plan tier (free, starter, professional, enterprise)

    Returns:
        Configured ProviderManager instance
    """
    providers = []

    # Yahoo Finance is always available as primary
    yf_provider = YahooFinanceProvider(api_key=yf_api_key)
    providers.append(yf_provider)

    # Add Alpha Vantage if key provided
    av_provider = None
    if av_api_key:
        av_provider = AlphaVantageProvider(api_key=av_api_key, premium=av_premium)
        providers.append(av_provider)

    # Add FMP if key provided
    fmp_provider = None
    if fmp_api_key:
        fmp_provider = FMPProvider(api_key=fmp_api_key, plan=fmp_plan)
        providers.append(fmp_provider)

    if not providers:
        raise ValueError("At least one provider must be configured")

    # Primary is Yahoo Finance, secondary is first available additional provider
    primary = providers[0]
    secondary = providers[1] if len(providers) > 1 else None

    return ProviderManager(primary=primary, secondary=secondary, mode=mode)