
import streamlit as st
import pandas as pd
import re
from io import BytesIO

def is_valid_date(value):
    try:
        pd.to_datetime(value, errors='raise')
        return True
    except Exception:
        return False

def is_valid_time(value):
    try:
        parsed = pd.to_datetime(value, errors='raise').time()
        return True
    except Exception:
        return False

def is_valid_gps(value):
    if isinstance(value, str):
        value = value.replace(',', '.').strip()
    try:
        val = float(value)
        return -180.0 <= val <= 180.0
    except:
        return False

def valideer_excel(uploaded_file, gps_kolomnamen):
    fouten = []
    gps_overzicht = []

    xls = pd.ExcelFile(uploaded_file)

    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        gps_kolommen_herkend = [col for col in df.columns if str(col).strip() in gps_kolomnamen]

        total_rows = len(df)
        missing_gps_rows = 0

        for row_idx, row in df.iterrows():
            row_has_gps_missing = False
            for col_idx, value in enumerate(row):
                fouttypes = []

                kolomnaam = df.columns[col_idx]
                waarde = value

                if pd.isna(value):
                    continue

                kolomnaam_lc = str(kolomnaam).lower()

                if 'datum' in kolomnaam_lc:
                    if not is_valid_date(value):
                        fouttypes.append('Ongeldige datum')

                if 'tijd' in kolomnaam_lc:
                    if not is_valid_time(value):
                        fouttypes.append('Ongeldige tijd')

                if kolomnaam in gps_kolommen_herkend:
                    if pd.isna(value) or str(value).strip() == "":
                        row_has_gps_missing = True
                    elif not is_valid_gps(value):
                        fouttypes.append('Ongeldige GPS')

                if isinstance(value, str):
                    if re.search(r'[^ -~]', value):
                        fouttypes.append('Vreemde tekens')
                    if re.search(r'\s{2,}', value):
                        fouttypes.append('Dubbele spatie')

                if fouttypes:
                    fouten.append({
                        'sheet': sheet_name,
                        'rij': row_idx + 2,
                        'kolom': kolomnaam,
                        'waarde': waarde,
                        'fouten': ', '.join(fouttypes)
                    })

            if row_has_gps_missing:
                missing_gps_rows += 1

        if total_rows > 0 and gps_kolommen_herkend:
            gps_overzicht.append({
                'sheet': sheet_name,
                'gps_kolommen': gps_kolommen_herkend,
                'totaal_rijen': total_rows,
                'rijen_met_ontbrekende_gps': missing_gps_rows,
                'percentage_ontbrekend': round((missing_gps_rows / total_rows) * 100, 2)
            })

    return pd.DataFrame(fouten), pd.DataFrame(gps_overzicht)

# Streamlit UI
st.title("Excel Validatie Tool")

uploaded_file = st.file_uploader("Upload een Excel bestand", type=["xls", "xlsx"])

if uploaded_file is not None:
    gps_kolommen = [
        "Instap Long", "Instap Lat",
        "Uitstap Long", "Uitstap Lat",
        "Loos Long", "Loos Lat"
    ]

    fouten_df, gps_info_df = valideer_excel(uploaded_file, gps_kolommen)

    st.subheader("Foutenrapport")
    if fouten_df.empty:
        st.success("Geen fouten gevonden in het bestand.")
    else:
        st.dataframe(fouten_df)
        csv = fouten_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download foutenrapport als CSV", data=csv, file_name="foutenrapport.csv")

    st.subheader("Overzicht GPS-kolommen")
    if not gps_info_df.empty:
        st.dataframe(gps_info_df)
    else:
        st.info("Geen GPS-kolommen gevonden of herkend.")
