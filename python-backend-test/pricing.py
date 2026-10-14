"""
pricing.py — Pricing calculation and Cosmos DB settings loader.

Reads the active pricing and exchange rate configuration from Cosmos DB,
and calculates a locked cost snapshot at generation/failure time.
"""

import datetime
from typing import Optional, Dict, Any

from config import cosmos_settings_container

DEFAULT_PRICING_SETTINGS = {
    "exchangeRate": 4.45,
    "visionInput": 2.50,
    "visionOutput": 10.00,
    "llmInput": 2.00,
    "llmCompletion": 8.00,
    "contentUnderstanding": 5.00
}


def load_pricing_settings() -> Optional[Dict[str, Any]]:
    """Fetch the active app settings document from Cosmos DB."""
    if cosmos_settings_container is None:
        return None

    try:
        settings = cosmos_settings_container.read_item(
            item="app-settings",
            partition_key="app-settings"
        )
        return settings
    except Exception as e:
        print(f"[pricing] Warning: Could not read app-settings from Cosmos DB: {e}")
        return None


def compute_cost_summary(
    metrics: Optional[Dict[str, Any]],
    settings: Optional[Dict[str, Any]] = None
) -> Optional[Dict[str, Any]]:
    """
    Compute a locked cost summary from token/page metrics and pricing settings.
    Returns None if neither settings nor metrics are available.
    """
    if not metrics:
        return None

    if settings is None:
        settings = load_pricing_settings()

    # If settings could not be fetched from DB, fall back to default pricing
    if not settings:
        settings = DEFAULT_PRICING_SETTINGS

    try:
        # Settings rates are defined per 1M tokens and per 1K pages
        vision_input_rate = float(settings.get("visionInput", DEFAULT_PRICING_SETTINGS["visionInput"])) / 1_000_000
        vision_output_rate = float(settings.get("visionOutput", DEFAULT_PRICING_SETTINGS["visionOutput"])) / 1_000_000
        llm_input_rate = float(settings.get("llmInput", DEFAULT_PRICING_SETTINGS["llmInput"])) / 1_000_000
        llm_completion_rate = float(settings.get("llmCompletion", DEFAULT_PRICING_SETTINGS["llmCompletion"])) / 1_000_000
        cu_page_rate = float(settings.get("contentUnderstanding", DEFAULT_PRICING_SETTINGS["contentUnderstanding"])) / 1_000
        exchange_rate = float(settings.get("exchangeRate", DEFAULT_PRICING_SETTINGS["exchangeRate"]))

        v_in = float(metrics.get("vision_tokens_prompt") or 0) * vision_input_rate
        v_out = float(metrics.get("vision_tokens_completion") or 0) * vision_output_rate
        l_in = float(metrics.get("llm_tokens_prompt") or 0) * llm_input_rate
        l_out = float(metrics.get("llm_tokens_completion") or 0) * llm_completion_rate
        cu = float(metrics.get("content_understanding_pages") or 0) * cu_page_rate

        cost_usd = v_in + v_out + l_in + l_out + cu
        cost_myr = cost_usd * exchange_rate

        return {
            "total_usd": round(cost_usd, 6),
            "total_myr": round(cost_myr, 2),
            "exchange_rate": exchange_rate,
            "breakdown_usd": {
                "vision": round(v_in + v_out, 6),
                "llm": round(l_in + l_out, 6),
                "content_understanding": round(cu, 6)
            },
            "rates_used": {
                "visionInput": float(settings.get("visionInput", DEFAULT_PRICING_SETTINGS["visionInput"])),
                "visionOutput": float(settings.get("visionOutput", DEFAULT_PRICING_SETTINGS["visionOutput"])),
                "llmInput": float(settings.get("llmInput", DEFAULT_PRICING_SETTINGS["llmInput"])),
                "llmCompletion": float(settings.get("llmCompletion", DEFAULT_PRICING_SETTINGS["llmCompletion"])),
                "contentUnderstanding": float(settings.get("contentUnderstanding", DEFAULT_PRICING_SETTINGS["contentUnderstanding"]))
            },
            "calculated_at": datetime.datetime.utcnow().isoformat() + "Z"
        }
    except Exception as e:
        print(f"[pricing] Error computing cost summary: {e}")
        return None
