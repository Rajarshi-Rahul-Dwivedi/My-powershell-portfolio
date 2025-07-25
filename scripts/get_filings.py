#Hello, This Script will not function in it's current form please read the discription to update the script acordingly or contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import pandas as pd
import requests
from io import StringIO

app = FastAPI()

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Or your specific frontend domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Pydantic schema for request body
class FilingRequest(BaseModel):
    name: str
    email: str
    tickers: list
    start_year: int
    start_qtr: int
    end_year: int
    end_qtr: int
    form_type: str  # e.g. "8-K"

@app.post("/get-filings")
def get_filings(req: FilingRequest):
    user_credentials = f"{req.name} {req.email}"

    def download_master_idx(year, quarter):
        url = f"https://www.sec.gov/Archives/edgar/full-index/{year}/QTR{quarter}/master.idx"
        headers = {'User-Agent': user_credentials}
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            return pd.DataFrame()
        lines = response.text.splitlines()
        for i, line in enumerate(lines):
            if line.strip().startswith("CIK|Company Name|Form Type|Date Filed|Filename"):
                lines = lines[i + 2:]
                break
        data = "\n".join(lines)
        df = pd.read_csv(StringIO(data), sep="|", header=None,
                         names=["CIK", "Company Name", "Form Type", "Date Filed", "Filename"])
        df["Year"] = year
        df["Quarter"] = quarter
        return df

    def get_cik_to_ticker_map():
        url = "https://www.sec.gov/files/company_tickers.json"
        headers = {'User-Agent': user_credentials}
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            return {}
        data = response.json()
        return {str(v["cik_str"]).zfill(10): v["ticker"] for v in data.values()}

    all_dfs = []
    for year in range(req.start_year, req.end_year + 1):
        for quarter in range(1, 5):
            if (year == req.start_year and quarter < req.start_qtr) or \
               (year == req.end_year and quarter > req.end_qtr):
                continue
            df = download_master_idx(year, quarter)
            if not df.empty:
                all_dfs.append(df)

    if not all_dfs:
        return {"data": []}

    combined = pd.concat(all_dfs, ignore_index=True)
    if req.form_type:
        combined = combined[combined["Form Type"] == req.form_type]

    cik_map = get_cik_to_ticker_map()
    combined["CIK"] = combined["CIK"].astype(str).str.zfill(10)
    combined["Ticker"] = combined["CIK"].map(cik_map)
    combined = combined[combined["Ticker"].isin(req.tickers)]

    combined["Filing URL"] = "https://www.sec.gov/Archives/" + combined["Filename"].str.replace(".txt", "-index.html")
    return {"data": combined[["Ticker", "Form Type", "Date Filed", "Filing URL"]].to_dict(orient="records")}
