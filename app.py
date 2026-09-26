import io
import pandas as pd
import streamlit as st
from geopy.extra.rate_limiter import RateLimiter
from geopy.geocoders import Nominatim

st.title("📍 Excel Reverse Geocoder")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.dataframe(df.head())

    if "Latitude" in df.columns and "Longitude" in df.columns:
        if st.button("Start Geocoding", type="primary"):
            geolocator = Nominatim(user_agent="excel_address_lookup_jupyter")
            reverse_geocode = RateLimiter(geolocator.reverse, min_delay_seconds=1)

            progress_bar = st.progress(0)
            status_text = st.empty()
            addresses = []

            for index, row in df.iterrows():
                status_text.text(f"Processing row {index + 1} of {len(df)}...")
                try:
                    location = reverse_geocode(f"{row['Latitude']}, {row['Longitude']}")
                    addresses.append(location.address if location else "Not Found")
                except Exception:
                    addresses.append("Error")
                progress_bar.progress((index + 1) / len(df))

            df["Address"] = addresses
            status_text.success("Geocoding Complete!")
            st.dataframe(df)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)

            st.download_button(
                label="📥 Download Result Excel File",
                data=buffer.getvalue(),
                file_name="geocoded_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
