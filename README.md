# SAP BOM Full Mapping Tool

This Streamlit app combines three existing tools into one workflow:

1. Upload system CPN/SPN data and generate `organized_cpn_spn.xlsx`
2. Upload system MPN/SPN data and generate `organized_mpn_spn.xlsx`
3. Upload customer Original BOM and run the final CPN/SPN/MPN mapping comparison

## Input files

### 1. System CPN/SPN data
Required source columns:
- `Material`
- `Customer material no.`

The tool automatically expands slash-separated CPN values.

### 2. System MPN/SPN data
Required source columns:
- `Material NO`
- `MFR. Name`
- `MFR. P/N`

### 3. Customer Original BOM
Required sheet name:
- `BOM`

Default data start row:
- `2`

Expected columns:
- A = Customer CPN
- B = Description
- C = Qty
- D = Location
- E/F = Primary MFG / MPN
- G/H, I/J... = Alternate MFG / MPN

## Output downloads

The app provides three downloadable Excel files:

1. `organized_cpn_spn.xlsx`
2. `organized_mpn_spn.xlsx`
3. `original_bom_mapping_result.xlsx`

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud deployment

1. Push these files to GitHub.
2. Go to Streamlit Community Cloud.
3. Select the repository.
4. Set main file path to `app.py`.
5. Deploy.
