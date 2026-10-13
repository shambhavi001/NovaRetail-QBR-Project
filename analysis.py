import pandas as pd

file = "novaretail_dataset.xlsx"

performance = pd.read_excel(
    file,
    sheet_name="Regional_Performance"
)

customer = pd.read_excel(
    file,
    sheet_name="Customer_Experience"
)

risk = pd.read_excel(
    file,
    sheet_name="Operational_Risk"
)

def get_kpis():

    current = performance[
        performance["Quarter"]=="2026-Q2"
    ]

    previous = performance[
        performance["Quarter"]=="2026-Q1"
    ]

    total_revenue = current["Revenue_USD_M"].sum()

    previous_revenue = previous["Revenue_USD_M"].sum()

    revenue_growth = (
        (total_revenue-previous_revenue)
        / previous_revenue
    ) * 100

    avg_nps = customer[
        customer["Quarter"]=="2026-Q2"
    ]["NPS"].mean()

    return {
        "revenue": f"${total_revenue/1000:.2f}B",
        "growth": f"{revenue_growth:+.1f}%",
        "nps": f"{avg_nps:.1f}"
    }
    
def get_top_region():

    current = performance[
        performance["Quarter"]=="2026-Q2"
    ]

    top_region = current.loc[
        current["Revenue_USD_M"].idxmax(),
        "Region"
    ]

    return top_region


def get_best_nps_region():

    current = customer[
        customer["Quarter"]=="2026-Q2"
    ]

    row = current.loc[
        current["NPS"].idxmax()
    ]

    return row["Region"]

def get_highest_risk_region():

    current = risk[
        risk["Quarter"]=="2026-Q2"
    ]

    row = current.loc[
        current["Stockout_Rate_pct"].idxmax()
    ]

    return row["Region"]

def get_business_facts():

    kpis = get_kpis()

    current_perf = performance[
        performance["Quarter"] == "2026-Q2"
    ]

    previous_perf = performance[
        performance["Quarter"] == "2026-Q1"
    ]

    current_customer = customer[
        customer["Quarter"] == "2026-Q2"
    ]

    current_risk = risk[
        risk["Quarter"] == "2026-Q2"
    ]

    # ==================================================
    # DYNAMIC SUMMARY FACTS
    # ==================================================

    top_revenue_region = current_perf.loc[
        current_perf["Revenue_USD_M"].idxmax(),
        "Region"
    ]

    best_customer_region = current_customer.loc[
        current_customer["NPS"].idxmax(),
        "Region"
    ]

    highest_risk_region = current_risk.loc[
        current_risk["Stockout_Rate_pct"].idxmax(),
        "Region"
    ]

    highest_returns_region = current_perf.loc[
        current_perf["Returns_Rate_pct"].idxmax(),
        "Region"
    ]

    growth_df = current_perf[
        ["Region", "Revenue_USD_M"]
    ].merge(
        previous_perf[
            ["Region", "Revenue_USD_M"]
        ],
        on="Region",
        suffixes=("_current", "_previous")
    )

    growth_df["growth_pct"] = (
        (
            growth_df["Revenue_USD_M_current"]
            -
            growth_df["Revenue_USD_M_previous"]
        )
        /
        growth_df["Revenue_USD_M_previous"]
    ) * 100

    fastest_growth_region = growth_df.loc[
        growth_df["growth_pct"].idxmax(),
        "Region"
    ]

    return {

        # ==================================================
        # EXECUTIVE KPIs
        # ==================================================

        "quarter": "2026-Q2",

        "revenue": kpis["revenue"],

        "growth": kpis["growth"],

        "average_nps": kpis["nps"],

        # ==================================================
        # REGIONAL PERFORMANCE
        # ==================================================

        "regional_revenue": {
            row["Region"]: float(row["Revenue_USD_M"])
            for _, row in current_perf.iterrows()
        },

        "units_sold": {
            row["Region"]: float(row["Units_Sold_K"])
            for _, row in current_perf.iterrows()
        },

        "returns_rate": {
            row["Region"]: float(row["Returns_Rate_pct"])
            for _, row in current_perf.iterrows()
        },

        "marketing_spend": {
            row["Region"]: float(row["Marketing_Spend_USD_M"])
            for _, row in current_perf.iterrows()
        },

        "revenue_history": performance[
            [
                "Quarter",
                "Region",
                "Revenue_USD_M"
            ]
        ].to_dict("records"),

        # ==================================================
        # CUSTOMER EXPERIENCE
        # ==================================================

        "regional_nps": {
            row["Region"]: float(row["NPS"])
            for _, row in current_customer.iterrows()
        },

        "regional_csat": {
            row["Region"]: float(row["CSAT_pct"])
            for _, row in current_customer.iterrows()
        },

        "delivery_days": {
            row["Region"]: float(row["Avg_Delivery_Days"])
            for _, row in current_customer.iterrows()
        },

        "support_tickets": {
            row["Region"]: float(row["Support_Tickets_K"])
            for _, row in current_customer.iterrows()
        },

        "nps_history": customer[
            [
                "Quarter",
                "Region",
                "NPS"
            ]
        ].to_dict("records"),

        # ==================================================
        # OPERATIONAL RISK
        # ==================================================

        "stockout_rate": {
            row["Region"]: float(row["Stockout_Rate_pct"])
            for _, row in current_risk.iterrows()
        },

        "return_processing_days": {
            row["Region"]: float(
                row["Return_Processing_Days"]
            )
            for _, row in current_risk.iterrows()
        },

        "warehouse_utilization": {
            row["Region"]: float(
                row["Warehouse_Capacity_Util_pct"]
            )
            for _, row in current_risk.iterrows()
        },

        "fraud_incidents": {
            row["Region"]: int(
                row["Fraud_Incidents"]
            )
            for _, row in current_risk.iterrows()
        },

        "risk_history": risk[
            [
                "Quarter",
                "Region",
                "Stockout_Rate_pct",
                "Fraud_Incidents"
            ]
        ].to_dict("records"),

        # ==================================================
        # AI HELPER FACTS
        # ==================================================

        "top_revenue_region":
            top_revenue_region,

        "fastest_growth_region":
            fastest_growth_region,

        "best_customer_region":
            best_customer_region,

        "highest_risk_region":
            highest_risk_region,

        "highest_returns_region":
            highest_returns_region
    }
print(get_business_facts())