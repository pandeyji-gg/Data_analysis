import pandas as pd

# Load data
frame = pd.read_excel("transportation-to-work-1.xlsx")

# Clean columns
frame["mode_name"] = frame["mode_name"].str.strip().str.capitalize()
frame["race_eth_name"] = frame["race_eth_name"].str.strip().str.capitalize()

# Most used transport
most_used = frame.groupby("mode_name")["pop_mode"].sum().sort_values(ascending=False)

# Public transport by race (raw)
public_transport = frame.loc[
    frame["mode"] == "PUBLICTR", ["race_eth_name", "mode", "pop_mode", "pop_total"]
]
raw_users = (
    public_transport.groupby("race_eth_name")["pop_mode"]
    .sum()
    .sort_values(ascending=False)
)

# Public transport by race (adjusted for population)
total_pop = frame.groupby("race_eth_name")["pop_total"].sum()
percent = (raw_users / total_pop).sort_values(ascending=False).drop("Total")

# Export
with pd.ExcelWriter("transport_analysis.xlsx") as writer:
    most_used.to_excel(writer, sheet_name="Most Used Transport")
    percent.to_excel(writer, sheet_name="Public Transport by Race")
