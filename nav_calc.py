import pandas as pd

# Mock daily NAV data
data = {"Date": ["2025-01-01", "2025-01-02", "2025-01-03"],
        "Assets": [1000000, 1015000, 1023000],
        "Liabilities": [200000, 205000, 210000]}

df = pd.DataFrame(data)
df["NAV"] = df["Assets"] - df["Liabilities"]

print(df)
