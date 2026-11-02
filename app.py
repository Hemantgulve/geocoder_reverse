import io
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Excel Reverse Geocoder", page_icon="📍")
st.title("📍 Excel Reverse Geocoder")

# ----------------------------------------------------------------------------
# Providers. Nominatim is free but capped at 1 request/sec by its usage policy,
# so for big files a keyed provider is MUCH faster (parallel requests).
# ----------------------------------------------------------------------------
PROVIDERS = {
    "Nominatim (free, no key, 1 req/sec)": {"rate": 1.0, "workers": 2, "key": False},
    "Geoapify (free API key, ~5 req/sec)": {"rate": 5.0, "workers": 8, "key": True},
    "LocationIQ (free API key, ~2 req/sec)": {"rate": 2.0, "workers": 4, "key": True},
}
USER_AGENT = "excel_address_lookup_streamlit/2.0"
PRECISION = 5  # ~1 m. Rows with the same rounded coordinates share one lookup.


class RateLimiter:
    """Thread-safe limiter: spaces request *starts*, lets responses overlap."""

    def __init__(self, per_sec):
        self.interval = 1.0 / per_sec
        self.lock = threading.Lock()
        self.next_slot = 0.0

    def wait(self):
        with self.lock:
            now = time.monotonic()
            start = max(now, self.next_slot)
            self.next_slot = start + self.interval
        if start > now:
            time.sleep(start - now)


def build_request(provider, api_key, lat, lon):
    if provider.startswith("Nominatim"):
        return (
            "https://nominatim.openstreetmap.org/reverse",
            {"lat": lat, "lon": lon, "format": "jsonv2"},
        )
    if provider.startswith("Geoapify"):
        return (
            "https://api.geoapify.com/v1/geocode/reverse",
            {"lat": lat, "lon": lon, "format": "json", "apiKey": api_key},
        )
    return (
        "https://us1.locationiq.com/v1/reverse",
        {"key": api_key, "lat": lat, "lon": lon, "format": "json"},
    )


def parse_response(provider, data):
    if provider.startswith("Geoapify"):
        results = data.get("results") or []
        return results[0].get("formatted", "Not Found") if results else "Not Found"
    return data.get("display_name") or "Not Found"


def reverse_one(session, limiter, provider, api_key, lat, lon, retries=3):
    url, params = build_request(provider, api_key, lat, lon)
    for attempt in range(retries):
        limiter.wait()
        try:
            r = session.get(url, params=params, timeout=15)
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(2 ** attempt)
                continue
            if r.status_code == 404:
                return "Not Found"
            r.raise_for_status()
            return parse_response(provider, r.json())
        except (requests.RequestException, ValueError):
            time.sleep(2 ** attempt)
    return "Error"


# Cache survives Streamlit reruns, so re-running never repeats finished lookups.
cache = st.session_state.setdefault("geo_cache", {})

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.dataframe(df.head())

    if "Latitude" not in df.columns or "Longitude" not in df.columns:
        st.error("The file must contain 'Latitude' and 'Longitude' columns.")
        st.stop()

    provider = st.selectbox("Geocoding provider", list(PROVIDERS))
    cfg = PROVIDERS[provider]
    api_key = ""
    if cfg["key"]:
        api_key = st.text_input("API key", type="password")

    # Clean + de-duplicate coordinates
    df["_lat"] = pd.to_numeric(df["Latitude"], errors="coerce").round(PRECISION)
    df["_lon"] = pd.to_numeric(df["Longitude"], errors="coerce").round(PRECISION)
    valid = df["_lat"].between(-90, 90) & df["_lon"].between(-180, 180)

    unique_coords = (
        df.loc[valid, ["_lat", "_lon"]].drop_duplicates().itertuples(index=False, name=None)
    )
    todo = [c for c in unique_coords if (provider, c) not in cache]
    n_unique = len(todo) + sum(
        1 for k in cache if k[0] == provider
    )  # rough info only

    est = len(todo) / cfg["rate"]
    st.info(
        f"{len(df)} rows → {int(valid.sum())} valid, "
        f"**{len(todo)} lookups needed** (duplicates & cached skipped). "
        f"Estimated time: ~{est / 60:.1f} min."
    )

    if st.button("Start Geocoding", type="primary"):
        if cfg["key"] and not api_key:
            st.error("Please enter an API key for this provider.")
            st.stop()

        progress_bar = st.progress(0.0)
        status_text = st.empty()

        if todo:
            limiter = RateLimiter(cfg["rate"])
            session = requests.Session()
            session.headers.update({"User-Agent": USER_AGENT})
            done = 0
            with ThreadPoolExecutor(max_workers=cfg["workers"]) as pool:
                futures = {
                    pool.submit(
                        reverse_one, session, limiter, provider, api_key, lat, lon
                    ): (lat, lon)
                    for lat, lon in todo
                }
                for fut in as_completed(futures):
                    result = fut.result()
                    if result != "Error":  # don't cache failures; retry next time
                        cache[(provider, futures[fut])] = result
                    else:
                        cache[("_err", futures[fut])] = result
                    done += 1
                    if done % 5 == 0 or done == len(todo):
                        status_text.text(f"Looked up {done} of {len(todo)} locations...")
                        progress_bar.progress(done / len(todo))

        def lookup(row_lat, row_lon, ok):
            if not ok:
                return "Invalid coordinates"
            return cache.get((provider, (row_lat, row_lon)), "Error")

        df["Address"] = [
            lookup(a, b, ok) for a, b, ok in zip(df["_lat"], df["_lon"], valid)
        ]
        out = df.drop(columns=["_lat", "_lon"])

        progress_bar.progress(1.0)
        n_err = int((out["Address"] == "Error").sum())
        status_text.success(
            "Geocoding Complete!" + (f" ({n_err} rows failed — click Start again to retry them.)" if n_err else "")
        )
        st.dataframe(out)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            out.to_excel(writer, index=False)

        st.download_button(
            label="📥 Download Result Excel File",
            data=buffer.getvalue(),
            file_name="geocoded_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
