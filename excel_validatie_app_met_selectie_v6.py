
import streamlit as st
import pandas as pd
import re
import time
from dateutil.parser import parse

st.title("Excel Validatie App met Selecteerbare Kolommen")

uploaded_file = st.file_uploader("Upload een Excel bestand", type=["xls", "xlsx"])
if uploaded_file:
    start_time = time.time()
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name)

    st.subheader("Stap 1: Selecteer kolommen voor controles")

    kolomnamen = df.columns.tolist()
    gps_kolommen = st.multiselect("GPS-kolommen", kolomnamen, default=[k for k in kolomnamen if any(x in k.lower() for x in ["long", "lat"])])
    tijd_kolommen = st.multiselect("Tijd-kolommen", kolomnamen, default=[k for k in kolomnamen if "tijd" in k.lower()])
    postcode_kolommen = st.multiselect("Postcode-kolommen", kolomnamen, default=[k for k in kolomnamen if "postcode" in k.lower()])
    datum_kolommen = st.multiselect("Datum-kolommen", kolomnamen, default=[k for k in kolomnamen if "datum" in k.lower()])
    whitespace_kolommen = st.multiselect("Kolommen voor whitespace-controle", kolomnamen, default=kolomnamen)

    if st.button("Start validatie"):
        fouten = []
        gps_info = {"ontbrekend": 0, "ongeldig": 0, "correct": 0}
        totaal_gps_rijen = 0

        for row_idx, row in df.iterrows():
            for kolomnaam, value in row.items():
                if pd.isna(value):
                    value_str = ""
                else:
                    value_str = str(value)
                fouttypes = []

                # GPS-controle
                if kolomnaam in gps_kolommen:
                    if value_str.strip() == "":
                        gps_info["ontbrekend"] += 1
                    else:
                        try:
                            float_val = float(value_str.replace(",", "."))
                            if -180 <= float_val <= 180:
                                gps_info["correct"] += 1
                            else:
                                gps_info["ongeldig"] += 1
                                fouttypes.append("GPS buiten bereik")
                        except ValueError:
                            gps_info["ongeldig"] += 1
                            fouttypes.append("Ongeldig GPS-formaat")
                    totaal_gps_rijen += 1

                # Tijd-controle met dateutil
                if kolomnaam in tijd_kolommen:
                    if value_str.strip():
                        try:
                            parse(value_str.strip(), fuzzy=False)
                        except Exception:
                            fouttypes.append("Ongeldige tijd")

                
                # Datum-controle met dateutil
                if kolomnaam in datum_kolommen:
                    if value_str.strip():
                        try:
                            parse(value_str.strip(), fuzzy=False)
                        except Exception:
                            fouttypes.append("Ongeldige datum")

                
                
                # Postcode-controle met slimme splitsing
                if kolomnaam in postcode_kolommen:
                    if value_str.strip():
                        if any(sep in value_str for sep in [",", ";", "|"]):
                            delen = re.split(r"[ ,;|]+", value_str.strip())
                        else:
                            delen = [value_str.strip()]
                        for deel in delen:
                            if deel and not re.match(r"^[1-9][0-9]{3}\s?[A-Z]{2}$", deel):
                                fouttypes.append("Ongeldige postcode")
                                break



                # Whitespace-controle
                if kolomnaam in whitespace_kolommen:
                    if re.match(r'^\s+', value_str):
                        fouttypes.append("Leading whitespace")
                    if re.search(r'\s+$', value_str):
                        fouttypes.append("Trailing whitespace")

                if fouttypes:
                    fouten.append({
                        "sheet": sheet_name,
                        "rij": row_idx + 2,
                        "kolom": kolomnaam,
                        "waarde": value_str,
                        "fouttype": ", ".join(fouttypes),
                    })

        if fouten:
            fouten_df = pd.DataFrame(fouten)
            fouten_df["waarde"] = fouten_df["waarde"].fillna("ONTBREKEND").astype(str)
            st.subheader("Gevonden fouten")
            st.dataframe(fouten_df)

            csv = fouten_df.to_csv(index=False).encode("utf-8")
            st.download_button("Download foutenrapport als CSV", data=csv, file_name="foutenrapport.csv", mime="text/csv")
        else:
            st.success("Geen fouten gevonden.")

        eindtijd = time.time()
        st.caption(f"Validatie uitgevoerd in {round(eindtijd - start_time, 2)} seconden.")

        if totaal_gps_rijen > 0:
            st.subheader("Overzicht GPS-kolommen")
            gps_info_df = pd.DataFrame([gps_info])
            gps_info_df["% ontbrekend"] = round(100 * gps_info["ontbrekend"] / totaal_gps_rijen, 2)
            gps_info_df["% correct"] = round(100 * gps_info["correct"] / totaal_gps_rijen, 2)
            gps_info_df["% ongeldig"] = round(100 * gps_info["ongeldig"] / totaal_gps_rijen, 2)
            st.dataframe(gps_info_df)
