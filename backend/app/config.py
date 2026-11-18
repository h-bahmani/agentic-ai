"""Environment configuration. No automatic downloads, migrations or cloud fallback."""
import os
from dataclasses import dataclass, field
from urllib.parse import urlsplit

READ_ONLY = frozenset(['database_schemas', 'database_tables', 'database_describe', 'database_relationships', 'database_row_count', 'database_distinct', 'quality_null_count', 'quality_duplicates', 'database_column_stats', 'quality_orphan_fks', 'sales_order_count', 'sales_revenue', 'sales_average_order', 'sales_cancellation_rate', 'sales_breakdown', 'sales_daily_trend', 'sales_compare_periods', 'customer_profile', 'customer_orders', 'customer_spend', 'customer_last_purchase', 'customer_top_spenders', 'customer_inactive', 'finance_payment_statuses', 'finance_failed_payments', 'finance_payment_mismatches', 'finance_unpaid_sales', 'subscriptions_status_counts', 'subscriptions_for_customer', 'subscriptions_no_autorenew', 'subscriptions_ending_before', 'support_open_tickets', 'support_categories', 'support_average_resolution', 'support_customer_tickets', 'marketing_campaign_profile', 'marketing_conversion_rate', 'marketing_response_breakdown', 'marketing_budget_per_conversion', 'database_catalog', 'database_relationship_graph', 'database_relationship_paths', 'database_safe_join', 'business_semantics', 'customer_overview', 'database_related_records'])

def default_tools():
    value=os.getenv("MCP_ALLOWLIST", "all_readonly")
    return READ_ONLY if value == "all_readonly" else frozenset(x.strip() for x in value.split(",") if x.strip())

@dataclass(frozen=True)
class Settings:
    provider: str = field(default_factory=lambda: os.getenv('LLM_PROVIDER', 'ollama'))
    llm_url: str = field(default_factory=lambda: os.getenv('OLLAMA_BASE_URL', 'http://127.0.0.1:11434'))
    model: str = field(default_factory=lambda: os.getenv('OLLAMA_MODEL', 'qwen3:4b'))
    mcp_url: str = field(default_factory=lambda: os.getenv('MCP_URL', 'http://127.0.0.1:8560/mcp'))
    allowlist: frozenset[str] = field(default_factory=default_tools)
    timeout: float = field(default_factory=lambda: float(os.getenv('LLM_TIMEOUT', '120')))
    run_timeout: float = field(default_factory=lambda: float(os.getenv('RUN_TIMEOUT', '240')))
    context: int = field(default_factory=lambda: int(os.getenv('LLM_CONTEXT', '4096')))
    temperature: float = field(default_factory=lambda: float(os.getenv('LLM_TEMPERATURE', '0')))
    max_steps: int = field(default_factory=lambda: int(os.getenv('MAX_TOOL_STEPS', '6')))
    report_max_pages: int = field(default_factory=lambda: int(os.getenv('REPORT_MAX_PAGES', '30')))
    token: str = field(default_factory=lambda: os.getenv('AGENT_API_TOKEN', ''))

    def __post_init__(self):
        for url in (self.llm_url, self.mcp_url):
            parsed = urlsplit(url)
            if parsed.scheme != 'http' or parsed.hostname not in {'127.0.0.1', 'localhost', '::1'} or parsed.username or parsed.password:
                raise ValueError('Only trusted loopback HTTP services are enabled in development.')
        if self.provider != 'ollama':
            raise ValueError('Only the implemented Ollama adapter is enabled.')
        if not self.allowlist <= READ_ONLY:
            raise ValueError('Unreviewed or write-capable tools cannot be enabled in Phase 1.')
        if not 1 <= self.max_steps <= 12 or not 512 <= self.context <= 16384 or self.timeout <= 0 or self.run_timeout <= 0:
            raise ValueError('Invalid execution limits.')
        if not 1<=self.report_max_pages<=50:raise ValueError('Invalid report pagination limit.')
