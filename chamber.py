from __future__ import annotations

import os
from datetime import date, datetime, timedelta
from pathlib import Path

import certifi
import pandas as pd
import plotly.express as px
import streamlit as st
from pymongo import MongoClient
from pymongo.errors import PyMongoError, ServerSelectionTimeoutError
from streamlit_calendar import calendar

CHAMBERS = ["Chamber A", "Chamber B", "Chamber C", "Chamber D", "Chamber E"]
REQUIRED_COLUMNS = ["User", "Project", "Chamber", "Date", "Start", "End"]
DEFAULT_DB_NAME = "chamber_booking"
DEFAULT_COLLECTION_NAME = "bookings"
LEGACY_EXCEL_FILE = Path(__file__).with_name("booking.xlsx")
ROOT_SECRETS_PATH = Path(__file__).resolve().parent / ".streamlit" / "secrets.toml"


st.set_page_config(page_title="Chamber Booking", layout="wide")
st.markdown(
    """
    <style>
        .stApp {background-color: #F4F8FF;}
        .stApp, .stApp p, .stApp div, .stApp span, .stApp label, .stApp li {
            color: #000000;
        }
        h1, h2, h3 {color: #000000;}
        [data-testid="stSidebar"] {
            background-color: #0F172A;
        }
        [data-testid="stSidebar"],
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] div,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] * {
            color: #FFFFFF !important;
        }
        [data-testid="stSidebar"] input,
        [data-testid="stSidebar"] label p,
        [data-testid="stSidebar"] label span {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }
        .stButton button,
        .stForm button,
        div[data-testid="stFormSubmitButton"] button {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            background-color: #1D4ED8 !important;
            border: 1px solid #1D4ED8 !important;
        }
        .stButton button:hover,
        .stForm button:hover,
        div[data-testid="stFormSubmitButton"] button:hover {
            color: #FFFFFF !important;
            background-color: #1E40AF !important;
            border-color: #1E40AF !important;
        }
        .stButton button *,
        .stForm button *,
        div[data-testid="stFormSubmitButton"] button *,
        .stButton button p,
        .stButton button span,
        .stButton button div,
        .stForm button p,
        .stForm button span,
        .stForm button div,
        div[data-testid="stFormSubmitButton"] button p,
        div[data-testid="stFormSubmitButton"] button span,
        div[data-testid="stFormSubmitButton"] button div {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }
        .stTextInput input,
        .stDateInput input,
        .stSelectbox [data-baseweb="select"] > div,
        .stForm input,
        .stForm textarea {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            background-color: #1F2937 !important;
        }
        .stSelectbox svg,
        .stDateInput svg {
            fill: #FFFFFF !important;
        }
        div[role="listbox"] ul,
        div[role="listbox"] li,
        div[role="option"] {
            color: #FFFFFF !important;
            background-color: #1F2937 !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


def load_secret_value(key: str) -> str | None:
    try:
        secret_value = st.secrets.get(key)
        if secret_value:
            return str(secret_value)
    except Exception:
        pass

    env_value = os.getenv(key)
    if env_value:
        return env_value

    if ROOT_SECRETS_PATH.exists():
        try:
            for raw_line in ROOT_SECRETS_PATH.read_text(encoding="utf-8").splitlines():
                line = raw_line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                parsed_key, value = line.split("=", 1)
                if parsed_key.strip() == key:
                    return value.strip().strip('"').strip("'")
        except Exception:
            return None

    return None


@st.cache_resource
def init_collection():
    uri = load_secret_value("MONGODB_URI")
    if not uri:
        raise RuntimeError("Missing MONGODB_URI")

    client = MongoClient(
        uri,
        tls=True,
        tlsCAFile=certifi.where(),
        serverSelectionTimeoutMS=10000,
    )
    client.admin.command("ping")

    db_name = load_secret_value("MONGODB_DB") or DEFAULT_DB_NAME
    collection_name = load_secret_value("MONGODB_COLLECTION") or DEFAULT_COLLECTION_NAME
    collection = client[db_name][collection_name]
    collection.create_index([("Date", 1), ("Chamber", 1), ("Start", 1)])
    return collection, db_name, collection_name


def normalize_date_value(value) -> str | None:
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()

    text = str(value).strip()
    if not text:
        return None

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return text
    return parsed.date().isoformat()


def normalize_time_value(value) -> str | None:
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.strftime("%H:%M:%S")
    if hasattr(value, "strftime") and not isinstance(value, str):
        try:
            return value.strftime("%H:%M:%S")
        except Exception:
            pass

    text = str(value).strip()
    if not text:
        return None

    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(text, fmt).strftime("%H:%M:%S")
        except ValueError:
            continue

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%H:%M:%S")


def parse_booking_time(value: str) -> datetime.time:
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(str(value), fmt).time()
        except ValueError:
            continue
    raise ValueError(f"Unsupported time format: {value}")


def bootstrap_legacy_excel(collection) -> int:
    if collection.count_documents({}, limit=1) > 0 or not LEGACY_EXCEL_FILE.exists():
        return 0

    try:
        legacy_df = pd.read_excel(LEGACY_EXCEL_FILE)
    except Exception:
        return 0

    documents = []
    for _, row in legacy_df.iterrows():
        normalized = {
            "User": str(row.get("User", "")).strip(),
            "Project": str(row.get("Project", "")).strip(),
            "Chamber": str(row.get("Chamber", "")).strip(),
            "Date": normalize_date_value(row.get("Date")),
            "Start": normalize_time_value(row.get("Start")),
            "End": normalize_time_value(row.get("End")),
            "created_at_utc": datetime.utcnow(),
        }
        if all(normalized.get(column) for column in REQUIRED_COLUMNS):
            documents.append(normalized)

    if documents:
        collection.insert_many(documents)
    return len(documents)


def load_bookings(collection) -> pd.DataFrame:
    documents = list(collection.find().sort([("Date", 1), ("Start", 1), ("Chamber", 1)]))
    if not documents:
        return pd.DataFrame(columns=["_id", *REQUIRED_COLUMNS])

    df = pd.DataFrame(documents)
    for column in REQUIRED_COLUMNS:
        if column not in df.columns:
            df[column] = None

    df["Date"] = df["Date"].apply(normalize_date_value)
    df["Start"] = df["Start"].apply(normalize_time_value)
    df["End"] = df["End"].apply(normalize_time_value)
    return df[["_id", *REQUIRED_COLUMNS]]


def calculate_hours(row: pd.Series) -> float:
    start = datetime.combine(date.today(), parse_booking_time(row["Start"]))
    end = datetime.combine(date.today(), parse_booking_time(row["End"]))
    return (end - start).total_seconds() / 3600


TIME_SLOTS = []
current = datetime.strptime("00:00", "%H:%M")
while current.strftime("%H:%M") != "23:30":
    TIME_SLOTS.append(current.strftime("%H:%M"))
    current += timedelta(minutes=30)
TIME_SLOTS.append("23:30")

try:
    bookings_collection, db_name, collection_name = init_collection()
except RuntimeError:
    st.error("Missing `MONGODB_URI` in Streamlit secrets or environment variables.")
    st.info(
        "For Streamlit Community Cloud, set `MONGODB_URI` in app secrets before deploying this page."
    )
    st.stop()
except ServerSelectionTimeoutError as exc:
    st.error("Could not connect to MongoDB.")
    st.code(str(exc))
    st.stop()

try:
    migrated_count = bootstrap_legacy_excel(bookings_collection)
    df = load_bookings(bookings_collection)
except PyMongoError as exc:
    st.error("MongoDB query failed.")
    st.code(str(exc))
    st.stop()

st.sidebar.title("Navigation")
st.sidebar.caption(f"MongoDB: `{db_name}.{collection_name}`")
if migrated_count > 0 and not st.session_state.get("legacy_excel_migrated_notice"):
    st.sidebar.success(f"Migrated {migrated_count} bookings from booking.xlsx")
    st.session_state["legacy_excel_migrated_notice"] = True

page = st.sidebar.radio("Select Page", ["Dashboard", "Booking", "Remove Booking", "Calendar"])

st.title("🏭 Chamber Booking System")

today = date.today().isoformat()
today_df = df[df["Date"].astype(str) == today]

if page == "Dashboard":
    st.subheader("Today's Chamber Status")
    cols = st.columns(len(CHAMBERS))
    for index, chamber in enumerate(CHAMBERS):
        chamber_bookings = today_df[today_df["Chamber"] == chamber]
        with cols[index]:
            st.markdown(
                f"""
                <div style="
                    border-radius:20px;
                    padding:20px;
                    background-color:white;
                    text-align:center;
                    height:320px;
                    border:3px solid #1976D2;
                    box-shadow:0px 3px 8px rgba(0,0,0,0.1);
                ">
                <h1 style="
                    color:#000000;
                    font-size:28px;
                ">
                    {chamber}
                </h1>
                """,
                unsafe_allow_html=True,
            )
            if chamber_bookings.empty:
                st.success("🟢 Available")
            else:
                st.error("🔴 Booked")
                for _, row in chamber_bookings.iterrows():
                    st.markdown(
                        f"""
                        <hr>
                        <h4 style="color:#000000;">
                            {row["Project"]}
                        </h4>
                        {row["User"]}
                        <br>
                        {row["Start"]} - {row["End"]}
                        """,
                        unsafe_allow_html=True,
                    )
            st.markdown("</div>", unsafe_allow_html=True)

    st.divider()
    st.subheader("Project Utilization Today")
    if not today_df.empty:
        dashboard_df = today_df.copy()
        dashboard_df["Hours"] = dashboard_df.apply(calculate_hours, axis=1)
        summary = dashboard_df.groupby("Project")["Hours"].sum().reset_index()
        total_hours = summary["Hours"].sum()
        summary["Usage %"] = (summary["Hours"] / total_hours) * 100
        fig = px.pie(summary, names="Project", values="Usage %", title="Today's Chamber Usage")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No booking today")

elif page == "Booking":
    st.subheader("New Booking")
    with st.form("booking_form"):
        user = st.text_input("User Name *")
        project = st.text_input("Project Name *")
        chamber = st.selectbox("Select Chamber", CHAMBERS)
        booking_date = st.date_input("Booking Date", min_value=date.today())
        start_slot = st.selectbox("Start Time", TIME_SLOTS, index=0)
        end_slot = st.selectbox("End Time", TIME_SLOTS, index=2)
        submit = st.form_submit_button("Book Chamber")

    if submit:
        if user.strip() == "" or project.strip() == "":
            st.error("❌ Please fill all required fields")
        else:
            start_time = datetime.strptime(start_slot, "%H:%M").time()
            end_time = datetime.strptime(end_slot, "%H:%M").time()
            new_start = datetime.combine(booking_date, start_time)
            new_end = datetime.combine(booking_date, end_time)

            if new_start < datetime.now():
                st.error("❌ Cannot book past time")
            elif new_end <= new_start:
                st.error("❌ End time must be later than start time")
            else:
                conflict = False
                same_day = df[
                    (df["Date"].astype(str) == booking_date.isoformat()) & (df["Chamber"] == chamber)
                ]
                for _, row in same_day.iterrows():
                    existing_start = datetime.combine(booking_date, parse_booking_time(row["Start"]))
                    existing_end = datetime.combine(booking_date, parse_booking_time(row["End"]))
                    if new_start < existing_end and new_end > existing_start:
                        conflict = True
                        break

                if conflict:
                    st.error("❌ Time conflict detected!")
                else:
                    document = {
                        "User": user.strip(),
                        "Project": project.strip(),
                        "Chamber": chamber,
                        "Date": booking_date.isoformat(),
                        "Start": start_time.strftime("%H:%M:%S"),
                        "End": end_time.strftime("%H:%M:%S"),
                        "created_at_utc": datetime.utcnow(),
                    }
                    try:
                        bookings_collection.insert_one(document)
                    except PyMongoError as exc:
                        st.error("❌ Booking could not be saved to MongoDB")
                        st.code(str(exc))
                    else:
                        st.success("✅ Booking Successful!")
                        st.rerun()

elif page == "Remove Booking":
    st.subheader("Remove Booking")
    filter_date = st.date_input("Select Date", value=date.today())
    filtered_df = df[df["Date"].astype(str) == filter_date.isoformat()]

    if not filtered_df.empty:
        remove_options = []
        for _, row in filtered_df.iterrows():
            label = f"{row['Chamber']} | {row['Project']} | {row['Start']} - {row['End']}"
            remove_options.append((row["_id"], label))

        selected_remove = st.selectbox(
            "Select Booking",
            remove_options,
            format_func=lambda option: option[1],
        )
        if st.button("Remove Booking"):
            try:
                result = bookings_collection.delete_one({"_id": selected_remove[0]})
            except PyMongoError as exc:
                st.error("❌ Booking could not be removed")
                st.code(str(exc))
            else:
                if result.deleted_count == 1:
                    st.success("✅ Booking Removed")
                    st.rerun()
                else:
                    st.warning("Booking was already removed")
    else:
        st.info("No booking found")

elif page == "Calendar":
    st.subheader("Chamber Calendar")
    selected_chamber = st.selectbox("Select Chamber", CHAMBERS)
    chamber_df = df[df["Chamber"] == selected_chamber].copy()
    chamber_df = chamber_df[pd.to_datetime(chamber_df["Date"], errors="coerce").dt.date >= date.today()]

    events = []
    for _, row in chamber_df.iterrows():
        events.append(
            {
                "title": f"{row['Project']} | {row['User']}",
                "start": f"{row['Date']}T{row['Start']}",
                "end": f"{row['Date']}T{row['End']}",
            }
        )

    calendar_options = {
        "initialView": "timeGridWeek",
        "slotMinTime": "00:00:00",
        "slotMaxTime": "24:00:00",
        "scrollTime": "09:00:00",
        "weekends": False,
        "height": 750,
    }
    calendar(events=events, options=calendar_options)
