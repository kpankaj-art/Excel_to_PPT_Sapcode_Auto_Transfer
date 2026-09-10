# Excel → PPT Safe SAP Transfer

A Streamlit app that safely transfers SAP/Dealer/Customer codes from Excel into an existing PowerPoint.

## Matching rule

The app is intentionally conservative so one dealer's SAP code is not accidentally added to another dealer.

1. Match **Outlet/Dealer Name + Contact Number**.
2. If that combination occurs more than once in the PPT, require **Size** as the third match.
3. If there is still no unique match, the SAP code is **NOT transferred**.
4. A PPT slide can only be assigned to one Excel row; duplicate assignments are blocked.
5. The app produces a matching report showing MATCHED / AMBIGUOUS / NO_MATCH and the reason.

## Flexible Excel column names

Examples supported:

- SAP: `Customer Code`, `Dealer Code`, `SAP Code`, `Sapcode`, `CustomerCode`
- Name: `DEALER / NAME`, `Dealer Name`, `Outlet Name`, `Customer Name`, `Retailer Name`
- Address: `DEALER / ADDERSS`, `Dealer Address`, `Address`, `Outlet Address`
- Contact: `DEALER / CONTACT`, `Contact No`, `Mobile`, `Phone`
- District: `District name`, `District`, `Dist`
- Size: `Size` or separate `W` + `H`
- Media: `TYPE`, `Media Type`

## PPT handling

The app reads the dealer information already present on each PPT slide. It updates an existing `Sapcode` field when available. If the PPT does not contain a SAP code field, it inserts:

```text
Sapcode : 11085686
```

after the District line in the main dealer information block.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## GitHub + Streamlit Cloud

1. Create a GitHub repository.
2. Upload `app.py`, `requirements.txt`, and `README.md`.
3. Deploy the repository on Streamlit Community Cloud.
4. Select `app.py` as the application entry point.

No API key is required.
