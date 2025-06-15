import pandas as pd
import zipfile
from io import BytesIO
from pathlib import Path
import re
from mlxtend.frequent_patterns import apriori, association_rules
from mlxtend.preprocessing import TransactionEncoder

# === File paths (adjust to your system if needed)
historical_zip_path = Path("/Users/mahtab/Desktop/AIRE/Input/Historical_BOM.zip")
output_path = Path("/Users/mahtab/Desktop/AIRE/Output/Apriori_Only_Historical.xlsx")
max_files = 5  # Limit to first 5 files

# === Format component code
def format_component(code):
    code = str(code)
    match = re.search(r"(KM\d+)", code)
    return match.group(1) if match else code.split("/")[0]

# === Load first N Excel files from historical zip
def load_boms(zip_path, max_files=5):
    data = []
    with zipfile.ZipFile(zip_path, 'r') as archive:
        files = [f for f in archive.infolist() if f.filename.endswith(".xlsx") and "__MACOSX" not in f.filename][:max_files]
        for file_info in files:
            with archive.open(file_info) as file:
                try:
                    df = pd.read_excel(BytesIO(file.read()), engine="openpyxl", dtype=str)
                    if "Component" in df.columns and "kmfg material" in df.columns:
                        df = df[["Component", "kmfg material"]].dropna()
                        df.columns = ["Component", "Material"]
                        df["Component"] = df["Component"].apply(format_component)
                        df["Source_File"] = Path(file_info.filename).name
                        data.append(df)
                except Exception as e:
                    print(f"⚠️ Skipping file {file_info.filename}: {e}")
    return pd.concat(data, ignore_index=True) if data else pd.DataFrame()

# === Load data
hist_df = load_boms(historical_zip_path, max_files=max_files)

# === Clean and prepare items
hist_df["Component"] = hist_df["Component"].str.strip().str.upper()
hist_df["Material"] = hist_df["Material"].str.strip().str.upper()
hist_df["Item"] = hist_df["Component"] + "__" + hist_df["Material"]

# === Group into transactions
transaction_df = pd.DataFrame({"File": hist_df["Source_File"], "Item": hist_df["Item"]})
transaction_basket = transaction_df.groupby("File")["Item"].apply(list).tolist()
transaction_basket = [t for t in transaction_basket if len(t) > 0]

# === Encode transactions (dense mode)
te = TransactionEncoder()
te_array = te.fit(transaction_basket).transform(transaction_basket)
basket = pd.DataFrame(te_array, columns=te.columns_)

# === Run Apriori
frequent_itemsets = apriori(basket, min_support=0.07, use_colnames=True)
rules = association_rules(frequent_itemsets, metric="confidence", min_threshold=0.5)

# === Extract simple 1-to-1 rules
rules = rules[
    (rules['antecedents'].apply(lambda x: len(x) == 1)) &
    (rules['consequents'].apply(lambda x: len(x) == 1))
].copy()

rules["Antecedent"] = rules["antecedents"].apply(lambda x: list(x)[0])
rules["Consequent"] = rules["consequents"].apply(lambda x: list(x)[0])
rules["Component"] = rules["Antecedent"].str.extract(r"^([^_]+)__")
rules["Material"] = rules["Consequent"].str.extract(r"^.*?__(.*?)$")

# === Final rule set
final_rules = rules[["Component", "Material", "support", "confidence", "lift"]].rename(columns={
    "support": "Support",
    "confidence": "Confidence",
    "lift": "Lift"
})

# === Save to Excel
final_rules.to_excel(output_path, sheet_name="Apriori_IfThen", index=False)
print(f"✅ Apriori rules saved to: {output_path}")
