/**
 * cost.ts — Cost resolution and formatting helper.
 * 
 * Provides unified resolution for task generation costs:
 * - Prefers locked cost_summary snapshots saved at generation/failure time.
 * - Falls back to dynamic calculation using cost_metrics + current settings for legacy items.
 */

export interface CostResult {
  costUsd: number;
  costMyr: number;
  exchangeRate: number;
  isLocked: boolean;
  calculatedAt?: string;
  ratesUsed?: Record<string, number>;
}

export function resolveCost(record: any, currentSettings?: any): CostResult | null {
  if (!record) return null;

  // 1. Preferred: Return the locked cost snapshot stored at generation time
  if (record.cost_summary && typeof record.cost_summary.total_myr === "number") {
    return {
      costUsd: Number(record.cost_summary.total_usd ?? 0),
      costMyr: Number(record.cost_summary.total_myr ?? 0),
      exchangeRate: Number(record.cost_summary.exchange_rate ?? 4.45),
      isLocked: true,
      calculatedAt: record.cost_summary.calculated_at,
      ratesUsed: record.cost_summary.rates_used,
    };
  }

  // 2. Legacy fallback: Dynamically recalculate if raw cost_metrics exist
  if (!record.cost_metrics) return null;

  const metrics = record.cost_metrics;
  const settings = currentSettings || {};

  const visionInputRate = (settings.visionInput ?? 2.50) / 1_000_000;
  const visionOutputRate = (settings.visionOutput ?? 10.00) / 1_000_000;
  const llmInputRate = (settings.llmInput ?? 2.00) / 1_000_000;
  const llmCompletionRate = (settings.llmCompletion ?? 8.00) / 1_000_000;
  const cuPageRate = (settings.contentUnderstanding ?? 5.00) / 1_000;
  const exchangeRate = Number(settings.exchangeRate ?? 4.45);

  let costUsd = 0;
  costUsd += (metrics.vision_tokens_prompt || 0) * visionInputRate;
  costUsd += (metrics.vision_tokens_completion || 0) * visionOutputRate;
  costUsd += (metrics.llm_tokens_prompt || 0) * llmInputRate;
  costUsd += (metrics.llm_tokens_completion || 0) * llmCompletionRate;
  costUsd += (metrics.content_understanding_pages || 0) * cuPageRate;

  const costMyr = costUsd * exchangeRate;

  return {
    costUsd,
    costMyr,
    exchangeRate,
    isLocked: false,
  };
}
