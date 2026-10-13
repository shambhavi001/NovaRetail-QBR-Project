import pandas as pd
import matplotlib.pyplot as plt


import pandas as pd
import matplotlib.pyplot as plt

df = pd.read_excel(
    "novaretail_dataset.xlsx",
    sheet_name="Regional_Performance"
)

pivot = df.pivot(
    index="Quarter",
    columns="Region",
    values="Revenue_USD_M"
)

fig, ax = plt.subplots(figsize=(8, 5))

pivot.plot(
    kind="bar",
    ax=ax
)

plt.title(
    "Revenue by region across quarters",
    fontsize=12,
    fontweight="bold"
)

plt.ylabel("Revenue (USD M)")
plt.xlabel("")

plt.xticks(rotation=0)

plt.legend(
    title=None,
    ncol=2,
    fontsize=8,
    frameon=False
)

plt.tight_layout()

plt.savefig(
    "charts/revenue.png",
    dpi=200,
    bbox_inches="tight"
)

plt.close()


df = pd.read_excel(
    "novaretail_dataset.xlsx",
    sheet_name="Customer_Experience"
)

pivot = df.pivot(
    index="Quarter",
    columns="Region",
    values="NPS"
)

plt.figure(figsize=(9, 5))

pivot.plot(
    marker="o",
    linewidth=2
)

plt.title(
    "NPS by Region Across Quarters"
)

plt.ylabel("NPS")
plt.xlabel("Quarter")

plt.legend(title="Region")

plt.tight_layout()

plt.savefig(
    "charts/customer.png"
)

plt.close()



df = pd.read_excel(
    "novaretail_dataset.xlsx",
    sheet_name="Operational_Risk"
)

pivot = df.pivot(
    index="Quarter",
    columns="Region",
    values="Stockout_Rate_pct"
)

plt.figure(figsize=(9, 5))

pivot.plot(
    marker="o",
    linewidth=2
)

plt.title(
    "Stockout Rate by Region"
)

plt.ylabel("Stockout Rate (%)")
plt.xlabel("Quarter")

plt.legend(title="Region")

plt.tight_layout()

plt.savefig(
    "charts/stockout_rate_pct.png"
)

plt.close()