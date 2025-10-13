import pandas as pd
import requests
import time

# Read Excel
df = pd.read_excel("top_200_universities.xlsx")

# Prepare output columns
df["Total_Citations"] = 0
df["Yearly_Citations"] = None

# Loop through each university
for i, row in df.iterrows():
    university = row["University"]
    print(f"Fetching citations for: {university}")

    # Use OpenAlex API
    url = f"https://api.openalex.org/insts?filter=display_name.search:{university}"
    r = requests.get(url)
    data = r.json()

    if "results" in data and data["results"]:
        openalex_id = data["results"][0]["id"]
        citation_url = f"https://api.openalex.org/works?filter=institutions.id:{openalex_id}"
        c = requests.get(citation_url).json()
        df.loc[i, "Total_Citations"] = c.get("meta", {}).get("count", 0)
    else:
        df.loc[i, "Total_Citations"] = 0

    time.sleep(1)  # to avoid rate-limit

# Save to new Excel
df.to_excel("citation_output.xlsx", index=False)
print("✅ Done! Results saved to citation_output.xlsx")
