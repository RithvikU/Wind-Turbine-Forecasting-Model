import pandas as pd
import plotly.express as px
import os

# ==============================
# Load Dataset
# ==============================

file_path = r"C:\Users\rithv\OneDrive\MyData\College\Projects\Wind Turbine Research\Final Databse1.xlsx"
df = pd.read_excel(file_path)

df.columns = df.columns.str.strip().str.lower()

# ==============================
# Required Columns
# ==============================

commission_col = "p_year"
decommission_col = "d_year"
capacity_col = "t_cap_filled"
generator_col = "gen_type_inferred"
site_col = "site_type"

df[commission_col] = pd.to_numeric(df[commission_col], errors="coerce")
df[decommission_col] = pd.to_numeric(df[decommission_col], errors="coerce")
df[capacity_col] = pd.to_numeric(df[capacity_col], errors="coerce")

df = df.dropna(subset=[commission_col])
df = df[df[commission_col] >= 1980]

df[decommission_col] = df[decommission_col].fillna(9999)

df[generator_col] = df[generator_col].fillna("Unknown")
df[site_col] = df[site_col].fillna("Unknown")

df["capacity_group"] = pd.cut(
    df[capacity_col],
    bins=[0, 1500, 3000, 6000, 10000],
    labels=["Small (0–1.5 MW)", "Medium (1.5–3 MW)", "Large (3–6 MW)", "Utility (6+ MW)"],
    include_lowest=True
).astype("object").fillna("Unknown")

# ==============================
# Build Yearly Active Turbine Table
# ==============================

years = range(int(df[commission_col].min()), 2026)

rows = []
for _, row in df.iterrows():
    for year in years:
        if row[commission_col] <= year <= row[decommission_col]:
            rows.append([
                year,
                row["capacity_group"],
                row[generator_col],
                row[site_col]
            ])

active_df = pd.DataFrame(rows, columns=["year", "capacity_group", "generator_type", "site_type"])

# ==============================
# Absolute Counts
# ==============================

absolute_capacity = active_df.groupby(["year", "capacity_group"]).size().reset_index(name="count")
absolute_gen = active_df.groupby(["year", "generator_type"]).size().reset_index(name="count")
absolute_site = active_df.groupby(["year", "site_type"]).size().reset_index(name="count")

# ==============================
# Proportional Shares
# ==============================

prop_capacity = absolute_capacity.copy()
prop_capacity["percentage"] = prop_capacity.groupby("year")["count"].transform(lambda x: x / x.sum())

prop_gen = absolute_gen.copy()
prop_gen["percentage"] = prop_gen.groupby("year")["count"].transform(lambda x: x / x.sum())

prop_site = absolute_site.copy()
prop_site["percentage"] = prop_site.groupby("year")["count"].transform(lambda x: x / x.sum())

# ==============================
# Output Directory
# ==============================

output_dir = "Wind_Turbine_Graphs"
os.makedirs(output_dir, exist_ok=True)

# ==============================
# Generate Graphs
# ==============================

# Absolute graphs
px.area(absolute_capacity, x="year", y="count", color="capacity_group",
        title="Active Turbines by Capacity Group").write_html(os.path.join(output_dir, "absolute_capacity.html"))

px.area(absolute_gen, x="year", y="count", color="generator_type",
        title="Active Turbines by Generator Type").write_html(os.path.join(output_dir, "absolute_generator.html"))

px.area(absolute_site, x="year", y="count", color="site_type",
        title="Active Turbines: Onshore vs Offshore").write_html(os.path.join(output_dir, "absolute_site_type.html"))

# Proportion graphs
px.area(prop_capacity, x="year", y="percentage", color="capacity_group",
        title="Proportion of Capacity Groups Over Time").write_html(os.path.join(output_dir, "proportion_capacity.html"))

px.area(prop_gen, x="year", y="percentage", color="generator_type",
        title="Proportion of Generator Types Over Time").write_html(os.path.join(output_dir, "proportion_generator.html"))

px.area(prop_site, x="year", y="percentage", color="site_type",
        title="Proportion: Onshore vs Offshore Over Time").write_html(os.path.join(output_dir, "proportion_site_type.html"))

print("Graphs generated in:", output_dir)
