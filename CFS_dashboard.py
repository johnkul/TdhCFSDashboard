from pathlib import Path
from difflib import SequenceMatcher
from io import BytesIO
import hashlib
import re
import time

import pandas as pd
import plotly.express as px
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
DATA_FILE_NAME = "CFS_QUESTIONNAIRE_Tdh_Kenya_T1.xlsx"
DATA_DIR = BASE_DIR / "data"
LOCAL_DATA_FALLBACK = Path(r"C:\Users\jekai\Desktop\DSC-Johntrial\CFS_Data\data") / DATA_FILE_NAME
DATA_PATH = DATA_DIR / DATA_FILE_NAME
if not DATA_PATH.exists() and LOCAL_DATA_FALLBACK.exists():
    DATA_PATH = LOCAL_DATA_FALLBACK
LOGO_PATH = BASE_DIR / "assets" / "tdh-logo.png"
CSS_PATH = BASE_DIR / "assets" / "styles.css"
PREPARED_CACHE_PATH = Path(__file__).with_name(".cfs_dashboard_prepared_cache.pkl")
CACHE_VERSION = "cfs-dashboard-prepared-v15"

MISSING = "Missing / unspecified"
REVIEW = "Needs review"
AGE_GROUP_ORDER = ["0-5 years", "6-12 years", "13-17 years", MISSING]
YES_NO_ORDER = ["Yes", "No", MISSING]
GENDER_ORDER = ["Girls", "Boys", "Transgender", MISSING]
TOP_N_OPTIONS = [5, 10, 15, 25]

ISSUE_COLUMNS = {
    "issue_new_arrival_lack_of_card": "New arrival / Lack of card",
    "issue_disability": "Disability",
    "issue_basic_needs": "Basic needs",
    "issue_deceased_parent": "Deceased parent",
    "issue_education": "Education",
    "issue_psychosocial_support": "Psychosocial support",
    "issue_neglected": "Neglected",
    "issue_parents_separated": "Parents separated",
    "issue_child_out_of_wedlock": "Child out of wedlock",
    "issue_food": "Food",
    "issue_clothing": "Clothing",
    "issue_shelter": "Shelter",
    "issue_reporting_protection_concern": "Reporting a protection concern",
    "issue_need_profiling_registration_unhcr": "Needs profiling / registration by UNHCR",
    "issue_none": "None",
}

SUPPORT_COLUMNS = {
    "support_psychological_first_aid": "Psychological First Aid",
    "support_play_art_therapy": "Play and art therapy",
    "support_psychoeducation": "Psychoeducation",
    "support_none": "None",
}

STAFF_MAP = {
    "mohamed sidi": "Mohamed Sidi",
    "mohamed": "Mohamed Sidi",
    "halima amin": "Halima Amin",
    "moge garad": "Moge Garad",
    "moge": "Moge Garad",
    "more garad": "Moge Garad",
    "daud hussein": "Daud Hussein",
    "maryan": "Maryan",
    "maslah mohamed hassan": "Maslah Mohamed Hassan",
    "maslah mohamed hasssan": "Maslah Mohamed Hassan",
    "maslah mohamed": "Maslah Mohamed Hassan",
    "maslsh mohamed hassan": "Maslah Mohamed Hassan",
    "maslish mohamed": "Maslah Mohamed Hassan",
    "maslah kohamed hassan": "Maslah Mohamed Hassan",
    "maslah kohamed": "Maslah Mohamed Hassan",
    "maslsh mohamed hasssan": "Maslah Mohamed Hassan",
    "ndayikeje ferdinand": "Ndayikeje Ferdinand",
    "ferdinand ndayikeje": "Ndayikeje Ferdinand",
    "ndayikeje": "Ndayikeje Ferdinand",
    "teresia natire thomas": "Teresia Natire Thomas",
    "teresia natire": "Teresia Natire Thomas",
    "haret derow ibrahim": "Haret Derow Ibrahim",
    "haret derow": "Haret Derow Ibrahim",
    "hared derow": "Haret Derow Ibrahim",
    "beatrice akwero": "Beatrice Akwero",
    "david otifo": "David Otifo",
    "farah mohamed hussein": "Farah Mohamed Hussein",
    "musdaf mohamed hassan": "Musdaf Mohamed Hassan",
    "musdaf mohamed": "Musdaf Mohamed Hassan",
    "hirwa gentille": "Hirwa Gentille",
    "oliek omot": "Oliek Omot",
    "nelson amanya": "Nelson Amanya",
    "dimo justin": "Dimo Justin",
    "dimo": "Dimo Justin",
    "louis kyanza": "Louis Kyanza",
    "dominic nangiro lomil": "Dominic Nangiro Lomil",
    "leer biel leer": "Leer Biel Leer",
    "spora niyikiza": "Spora Niyikiza",
    "niyikiza spora": "Spora Niyikiza",
    "jean claude": "Jean Claude",
    "fowzia omar": "Fowzia Omar",
    "yaak akech": "Yaak Akech",
    "peter kingombe": "Peter Kingombe",
    "dual ador arok": "Dual Ador Arok",
    "dual ador": "Dual Ador Arok",
    "gatwech bayak": "Gatwech Bayak",
    "armele ngakani": "Armele Ngakani",
    "john wani": "John Wani",
    "safari david": "Safari David",
    "epusie brenda": "Epusie Brenda",
    "zahara issack": "Zahara Issack",
    "zahara": "Zahara Issack",
    "nyakhor buob": "Nyakhor Buob Tang",
    "nyqkhor buob tang": "Nyakhor Buob Tang",
    "nyakhor": "Nyakhor Buob Tang",
    "halimo ahmed": "Halimo Ahmed",
    "halimo": "Halimo Ahmed",
    "agnes ingiara": "Agnes Ingiara",
    "oweteshe mirindi": "Oweteshe Mirindi",
    "rahmo abdi": "Rahmo Abdi",
    "ongoro john": "Ongoro John",
    "abdikadir osman": "Abdikadir Osman",
    "salma said": "Salma Said",
    "abdiwakil ali": "Abdiwakil Ali",
    "manow muse": "Manow Muse",
    "jama mohamed": "Jama Mohamed",
}

LOCATION_MAP = {
    "hagadera camp": "Hagadera",
    "hagadera": "Hagadera",
    "ifo 1": "Ifo 1",
    "ifo 2": "Ifo 2",
    "dagahaley": "Dagahaley",
    "kalobeyei reception center": "Kalobeyei Reception Centre",
    "kalobeyei reception centre": "Kalobeyei Reception Centre",
    "kalobeyei village 1": "Kalobeyei Village 1",
    "kalobeyei village 2": "Kalobeyei Village 2",
    "kalobeyei village 3": "Kalobeyei Village 3",
    "kalobeyei host": "Kalobeyei Host Community",
}

CFS_MAP = {
    "kalobeyei reception center cfs": "Kalobeyei Reception Centre CFS",
    "kalobeyei reception centre cfs": "Kalobeyei Reception Centre CFS",
    "host community": "Host Community CFS",
    "ifo 2 mobile cfs": "Ifo 2 Mobile CFS",
    "ifo mobile cfs 1": "Ifo Mobile CFS 1",
}

GAME_MAP = {
    "take5": "Take 5",
    "take 5": "Take 5",
    "take5 training": "Take 5",
    "take 5 training": "Take 5",
    "take5 routine": "Take 5",
    "take 5 routine": "Take 5",
    "ismf": "I Support My Friend",
    "imsf": "I Support My Friend",
    "i support my friend": "I Support My Friend",
    "psycho education": "Psychoeducation",
    "psychoeducation": "Psychoeducation",
    "merry go round": "Merry-go-round",
    "see saw": "See-saw",
    "slider": "Slider",
    "see saw and slider": "See-saw / Slider",
    "see saw and sliders": "See-saw / Slider",
    "modelling": "Modelling clay",
    "modelling clay": "Modelling clay",
    "meddling clay": "Modelling clay",
    "book reading": "Story / book reading",
    "reading story books": "Story / book reading",
    "story telling": "Story / book reading",
    "playing": "Free play",
}

ISSUE_OTHER_MAP = {
    "playing": "Play / child engagement",
    "playing with the kids": "Play / child engagement",
    "playing with other children": "Play / child engagement",
    "playing colouring book": "Play / child engagement",
    "playing and reading story book": "Play / child engagement",
    "take5": "Take 5",
    "take 5": "Take 5",
    "take5 training": "Take 5",
    "take 5 training": "Take 5",
    "take5 routine": "Take 5",
    "take 5 routine": "Take 5",
    "scholastic materials": "Scholastic materials",
    "materials support": "Scholastic materials",
    "card separation": "Card separation",
    "separation card": "Card separation",
    "abduction": "Threat / abduction",
    "threat of abduction": "Threat / abduction",
    "gbv": "GBV",
    "gbv survivor": "GBV",
    "defilement": "GBV",
    "physical assault": "Protection concern",
    "medical": "Medical concern",
    "medical concern": "Medical concern",
    "basic needs": "Basic needs",
    "bamba chakula issues": "Food assistance issue",
    "alternative food collector": "Food assistance issue",
    "pl": "Unclear / needs review",
    "o00": "Unclear / needs review",
}

AGENCY_MAP = {
    "dras": "DRS",
    "drs": "DRS",
    "dr s": "DRS",
    "dra": "DRS",
    "unhcr": "UNHCR",
    "uhcr": "UNHCR",
    "unchr": "UNHCR",
    "un": "UNHCR",
    "unhcr and dras": "UNHCR / DRS",
    "dras and unhcr": "UNHCR / DRS",
    "dras and unchr": "UNHCR / DRS",
    "unhcr dras": "UNHCR / DRS",
    "dras unhcr": "UNHCR / DRS",
    "lwf and dras": "LWF and DRS",
    "lwf and drs": "LWF and DRS",
    "dras and lwf": "DRS and LWF",
    "drs and lwf": "DRS and LWF",
    "lwf": "LWF",
    "lwfl": "LWF",
    "lwf education": "LWF Education",
    "lwf education sector": "LWF Education",
    "hi": "HI",
    "wfp": "WFP",
    "pwj": "PWJ",
    "peace wind japan": "PWJ",
    "nrc": "NRC",
    "drc": "DRC",
    "rck": "RCK",
    "krcs": "KRCS",
    "police and krcs": "Police / KRCS",
    "save the children": "Save the Children",
    "save children": "Save the Children",
    "iftin primary": "Iftin Primary",
    "000": "Unknown",
    "n": "Unknown",
}


def norm_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def smart_title(value):
    if pd.isna(value) or str(value).strip() == "":
        return MISSING
    text = re.sub(r"\s+", " ", str(value).strip())
    upper_tokens = {"cfs", "tdh", "unhcr", "drs", "dras", "lwf", "hi", "wfp", "nrc", "drc", "rck", "krcs", "pss", "gbv"}
    lower_tokens = {"and", "or", "of", "the", "for", "to", "in", "by", "with", "at"}
    titled = []
    for token in text.split():
        key = token.lower()
        if key in upper_tokens:
            titled.append("DRS" if key == "dras" else token.upper())
        elif key in lower_tokens:
            titled.append(key)
        else:
            titled.append(token.capitalize())
    return " ".join(titled)


def yes_no(value):
    if pd.isna(value):
        return MISSING
    numeric_value = pd.to_numeric(value, errors="coerce")
    if pd.notna(numeric_value):
        if numeric_value == 1:
            return "Yes"
        if numeric_value == 0:
            return "No"
    key = norm_text(value)
    if key in {"yes", "y", "1", "true"}:
        return "Yes"
    if key in {"no", "n", "0", "false"}:
        return "No"
    return MISSING


def clean_gender(value):
    key = norm_text(value)
    if not key:
        return MISSING
    if key in {"girl", "girls", "female", "f"}:
        return "Girls"
    if key in {"boy", "boys", "male", "m"}:
        return "Boys"
    if key in {
        "transgender",
        "trans gender",
        "trans",
        "tg",
        "other",
        "others",
        "other gender",
        "other specify",
        "others specify",
        "specify other",
        "trans boy",
        "trans girl",
        "trans male",
        "trans female",
    }:
        return "Transgender"
    return MISSING


def is_yes(value):
    return yes_no(value) == "Yes"


def harmonize_from_map(value, mapping, fallback_title=True):
    key = norm_text(value)
    if not key:
        return MISSING
    if key in mapping:
        return mapping[key]
    return smart_title(value) if fallback_title else value


def fuzzy_harmonize(value, mapping, canonical_values=None, cutoff=0.86):
    key = norm_text(value)
    if not key:
        return MISSING
    if key in mapping:
        return mapping[key]
    candidates = canonical_values or sorted(set(mapping.values()))
    best_score = 0
    best_value = None
    for candidate in candidates:
        score = SequenceMatcher(None, key, norm_text(candidate)).ratio()
        if score > best_score:
            best_score = score
            best_value = candidate
    if best_value and best_score >= cutoff:
        return best_value
    return smart_title(value)


def shorten_disability_type(value):
    key = norm_text(value)
    if not key:
        return MISSING
    if "downsyndrome" in key or "down syndrome" in key:
        return "Down syndrome"
    if "neurological" in key or "adhd" in key or "autism" in key or "dyslexia" in key:
        return "Neurological impairments"
    if "chronic" in key or "diabetes" in key or "blood pressure" in key:
        return "Chronic illnesses"
    return smart_title(value)


def clean_referral_destination(value):
    key = norm_text(value)
    if not key:
        return MISSING
    if key == "pss":
        return "PSS"
    if "case" in key:
        return "Case Management"
    if "empower" in key:
        return "Empowerment"
    if "external" in key:
        return "External referrals"
    return smart_title(value)


def age_group(age):
    if pd.isna(age):
        return MISSING
    try:
        age = float(age)
    except Exception:
        return MISSING
    if 0 <= age <= 5:
        return "0-5 years"
    if 6 <= age <= 12:
        return "6-12 years"
    if 13 <= age <= 17:
        return "13-17 years"
    return MISSING


def choose_first_nonmissing(row, columns):
    for col in columns:
        value = row.get(col, None)
        if pd.notna(value) and str(value).strip() != "":
            return value
    return None


def combine_location(row):
    if str(row.get("settlement_clean", "")).lower().startswith("kalobeyei"):
        return choose_first_nonmissing(row, ["camp_location_alt", "specific_camp_location", "exact_registered_location"])
    return choose_first_nonmissing(row, ["specific_camp_location", "camp_location_alt", "exact_registered_location"])


def combine_cfs(row):
    return choose_first_nonmissing(row, ["child_friendly_space_visited", "cfs_visited"])


def is_take5_text(*values):
    for value in values:
        key = norm_text(value)
        if not key:
            continue
        if re.search(r"(^|\s)take\s+5($|\s|[a-z])", key):
            return True
        if re.search(r"(^|\s)take5($|\s|[a-z])", key):
            return True
    return False


def clean_game(row):
    raw = row.get("games_played", None)
    other = row.get("game_other_specify", None)
    raw_key = norm_text(raw)
    other_key = norm_text(other)
    if is_take5_text(raw, other):
        return "Take 5"
    if raw_key in {"", "other", "others", "specify other"} and norm_text(other):
        return fuzzy_harmonize(other, GAME_MAP, cutoff=0.84)
    if norm_text(raw):
        return fuzzy_harmonize(raw, GAME_MAP, cutoff=0.84)
    if norm_text(other):
        return fuzzy_harmonize(other, GAME_MAP, cutoff=0.84)
    return MISSING


def clean_staff(value):
    return fuzzy_harmonize(value, STAFF_MAP, cutoff=0.88)


def clean_issue_other(value):
    if is_take5_text(value):
        return "Take 5"
    return fuzzy_harmonize(value, ISSUE_OTHER_MAP, cutoff=0.84)


def ordered_unique(series, order=None):
    values = [v for v in series.dropna().unique().tolist() if str(v).strip() != ""]
    values = [v for v in values if v != MISSING] + ([MISSING] if MISSING in values else [])
    if order:
        order_map = {v: i for i, v in enumerate(order)}
        return sorted(values, key=lambda v: (order_map.get(v, 999), str(v)))
    return sorted(values, key=lambda v: str(v))


@st.cache_data(show_spinner=False, ttl=300)
def load_and_prepare(modified_time):
    try:
        if PREPARED_CACHE_PATH.exists():
            payload = pd.read_pickle(PREPARED_CACHE_PATH)
            if (
                payload.get("cache_version") == CACHE_VERSION
                and payload.get("source_path") == str(DATA_PATH)
                and payload.get("modified_time") == modified_time
            ):
                return payload["df"], payload["issue_long"], payload["support_long"], payload["game_long"]
    except Exception:
        pass

    df = pd.read_excel(DATA_PATH)
    df = df.copy()
    df.insert(0, "record_id", range(1, len(df) + 1))

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["month"] = df["date"].dt.to_period("M").astype("string")
    df["consent_clean"] = df["consent"].map(yes_no)
    df["staff_clean"] = df["staff_filling_form"].map(clean_staff)
    df["gender_clean"] = df["child_gender"].map(clean_gender)
    df["disability_status_clean"] = df["child_living_with_disability"].map(yes_no)
    df["first_visit_clean"] = df["first_visit_tdh_cfs"].map(yes_no)
    df["referral_made_clean"] = df["referral_made"].map(yes_no)
    df["referral_destination_grouped"] = df["referral_destination"].map(clean_referral_destination)
    df["external_referral_agency_clean"] = df["external_referral_agency"].map(lambda v: fuzzy_harmonize(v, AGENCY_MAP, cutoff=0.86))

    df["age_clean"] = pd.to_numeric(df["child_age"], errors="coerce")
    df["age_group"] = df["age_clean"].map(age_group)
    df["age_group"] = pd.Categorical(df["age_group"], categories=AGE_GROUP_ORDER, ordered=True)
    df["gender_clean"] = pd.Categorical(df["gender_clean"], categories=GENDER_ORDER, ordered=True)
    df["first_visit_clean"] = pd.Categorical(df["first_visit_clean"], categories=YES_NO_ORDER, ordered=True)
    df["disability_status_clean"] = pd.Categorical(df["disability_status_clean"], categories=YES_NO_ORDER, ordered=True)
    df["referral_made_clean"] = pd.Categorical(df["referral_made_clean"], categories=YES_NO_ORDER, ordered=True)

    df["disability_type_clean"] = df["disability_type"].map(lambda v: harmonize_from_map(v, {}, fallback_title=True))
    df["disability_type_display"] = df["disability_type"].map(shorten_disability_type)
    df.loc[df["disability_status_clean"].astype("string") != "Yes", "disability_type_display"] = MISSING

    df["settlement_clean"] = df["camp_of_information_seeking"].map(lambda v: harmonize_from_map(v, {}, fallback_title=True))
    df["location_raw"] = df.apply(combine_location, axis=1)
    df["location_clean"] = df["location_raw"].map(lambda v: fuzzy_harmonize(v, LOCATION_MAP, cutoff=0.86))
    df["cfs_raw"] = df.apply(combine_cfs, axis=1)
    df["cfs_clean"] = df["cfs_raw"].map(lambda v: fuzzy_harmonize(v, CFS_MAP, cutoff=0.86))
    df["games_played_clean"] = df.apply(clean_game, axis=1)
    df["issue_other_specify_clean"] = df["issue_other_specify"].map(clean_issue_other)

    issue_frames = []
    for col, label in ISSUE_COLUMNS.items():
        if col in df.columns:
            mask = df[col].map(is_yes)
            tmp = df.loc[mask, ["record_id"]].copy()
            tmp["issue_clean"] = label
            issue_frames.append(tmp)
    if "issue_other" in df.columns and "issue_other_specify_clean" in df.columns:
        mask = df["issue_other"].map(is_yes) & ~df["issue_other_specify_clean"].isin([MISSING, REVIEW])
        tmp = df.loc[mask, ["record_id", "issue_other_specify_clean"]].copy()
        tmp = tmp.rename(columns={"issue_other_specify_clean": "issue_clean"})
        issue_frames.append(tmp)
    issue_long = pd.concat(issue_frames, ignore_index=True) if issue_frames else pd.DataFrame(columns=["record_id", "issue_clean"])
    issue_long = issue_long.drop_duplicates(["record_id", "issue_clean"]).reset_index(drop=True)

    support_frames = []
    for col, label in SUPPORT_COLUMNS.items():
        if col in df.columns:
            mask = df[col].map(is_yes)
            tmp = df.loc[mask, ["record_id"]].copy()
            tmp["support_clean"] = label
            support_frames.append(tmp)
    support_long = pd.concat(support_frames, ignore_index=True) if support_frames else pd.DataFrame(columns=["record_id", "support_clean"])
    support_long = support_long.drop_duplicates(["record_id", "support_clean"]).reset_index(drop=True)

    game_long = (
        df.loc[~df["games_played_clean"].isin([MISSING, REVIEW]), ["record_id", "games_played_clean"]]
        .rename(columns={"games_played_clean": "game_clean"})
        .drop_duplicates()
        .reset_index(drop=True)
    )

    issues_combined = issue_long.groupby("record_id")["issue_clean"].apply(lambda x: "; ".join(sorted(set(x)))).rename("issues_combined")
    support_combined = support_long.groupby("record_id")["support_clean"].apply(lambda x: "; ".join(sorted(set(x)))).rename("support_combined")
    df = df.merge(issues_combined, on="record_id", how="left").merge(support_combined, on="record_id", how="left")
    df["issues_combined"] = df["issues_combined"].fillna(MISSING)
    df["support_combined"] = df["support_combined"].fillna(MISSING)

    try:
        pd.to_pickle(
            {
                "cache_version": CACHE_VERSION,
                "source_path": str(DATA_PATH),
                "modified_time": modified_time,
                "df": df,
                "issue_long": issue_long,
                "support_long": support_long,
                "game_long": game_long,
            },
            PREPARED_CACHE_PATH,
        )
    except Exception:
        pass

    return df, issue_long, support_long, game_long


def table_with_total(data, index, columns=None, value_col="record_id"):
    if data.empty:
        return pd.DataFrame()
    columns = columns or []
    grouped = data.groupby(index + columns, observed=False)[value_col].count()
    if columns:
        table = grouped.unstack(columns, fill_value=0)
    else:
        table = grouped.to_frame("Count")
    if isinstance(table.columns, pd.MultiIndex):
        table[("Total",) + ("",) * (table.columns.nlevels - 1)] = table.sum(axis=1)
    elif "Count" in table.columns and len(table.columns) == 1:
        table = drop_zero_missing_table_labels(table)
        if index == ["age_group"]:
            table = table.sort_index()
        else:
            table = table.sort_values("Count", ascending=False)
        table.loc["Grand Total", "Count"] = table["Count"].sum()
        return table
    else:
        table["Total"] = table.sum(axis=1)
    total_col = ("Total",) + ("",) * (table.columns.nlevels - 1) if isinstance(table.columns, pd.MultiIndex) else "Total"
    table = drop_zero_missing_table_labels(table)
    if "age_group" in index:
        table = table.sort_index()
    else:
        table = table.sort_values(total_col, ascending=False)
    table.loc["Grand Total"] = table.sum(numeric_only=True)
    return table


def add_interview_date_columns(table, data, group_col, date_col):
    if table.empty or data.empty or group_col not in data.columns or date_col not in data.columns:
        return table

    out = table.copy()
    date_summary = data.dropna(subset=[date_col]).groupby(group_col, observed=False)[date_col].agg(["min", "max"])

    def fmt_date(value):
        if pd.isna(value):
            return ""
        return pd.to_datetime(value).strftime("%d %b %Y")

    first_dates = {}
    latest_dates = {}
    for row_label in out.index:
        if str(row_label) == "Grand Total":
            first_dates[row_label] = fmt_date(data[date_col].min())
            latest_dates[row_label] = fmt_date(data[date_col].max())
        elif row_label in date_summary.index:
            first_dates[row_label] = fmt_date(date_summary.loc[row_label, "min"])
            latest_dates[row_label] = fmt_date(date_summary.loc[row_label, "max"])
        else:
            first_dates[row_label] = ""
            latest_dates[row_label] = ""

    if isinstance(out.columns, pd.MultiIndex):
        suffix = ("",) * (out.columns.nlevels - 1)
        out[("First Date of Interview",) + suffix] = pd.Series(first_dates)
        out[("Latest Interview Date",) + suffix] = pd.Series(latest_dates)
    else:
        out["First Date of Interview"] = pd.Series(first_dates)
        out["Latest Interview Date"] = pd.Series(latest_dates)
    return out


def label_has_missing(label):
    if isinstance(label, tuple):
        return any(str(part) == MISSING for part in label)
    return str(label) == MISSING


def numeric_sum(values):
    return pd.to_numeric(values, errors="coerce").fillna(0).sum()


def drop_zero_missing_table_labels(table):
    if table.empty:
        return table

    keep_rows = []
    for row_label, row in table.iterrows():
        drop_row = label_has_missing(row_label) and numeric_sum(row) == 0
        keep_rows.append(not drop_row)
    table = table.loc[keep_rows]

    keep_cols = []
    for col_label in table.columns:
        drop_col = label_has_missing(col_label) and numeric_sum(table[col_label]) == 0
        keep_cols.append(not drop_col)
    return table.loc[:, keep_cols]


def style_total(table):
    if table.empty:
        return table

    def row_style(row):
        if str(row.name) == "Grand Total":
            return ["font-weight: 800; background-color: #dbeafe; color: #0f172a;" for _ in row]
        return ["" for _ in row]

    styler = (
        table.style
        .apply(row_style, axis=1)
        .format(precision=0, thousands=",")
        .set_table_styles(
            [
                {"selector": "th", "props": [("background-color", "#eaf2fb"), ("color", "#111827"), ("font-weight", "800"), ("padding", "9px 10px"), ("border", "1px solid #d5dee9")]},
                {"selector": "td", "props": [("color", "#1f2937"), ("padding", "8px 10px"), ("border", "1px solid #e5edf5")]},
                {"selector": "tbody tr:nth-child(even)", "props": [("background-color", "#f8fafc")]},
            ]
        )
    )
    return styler


def count_table(data, columns, order_col=None):
    if data.empty:
        return pd.DataFrame(columns=columns + ["Count"])
    out = data.groupby(columns, observed=False).size().reset_index(name="Count")
    out = out[out["Count"] > 0]
    if order_col == "age_group":
        out[order_col] = pd.Categorical(out[order_col], categories=AGE_GROUP_ORDER, ordered=True)
        out = out.sort_values([order_col, "Count"], ascending=[True, False])
    else:
        out = out.sort_values("Count", ascending=False)
    return out


def sort_chart_by_total(data, category_col, value_col="Count", keep_age_order=False, axis="x"):
    if data.empty:
        return data, []
    if keep_age_order or category_col == "age_group":
        order = [v for v in AGE_GROUP_ORDER if v in set(data[category_col].astype("string"))]
        return data, order
    ascending = axis == "y"
    totals = (
        data.groupby(category_col, observed=False)[value_col]
        .sum()
        .sort_values(ascending=ascending)
    )
    order = totals.index.astype("string").tolist()
    return data, order


def top_n_chart_control(data, category_col, key, label="Chart category slicer", default=10):
    if data.empty or category_col not in data.columns:
        return data
    category_count = data[category_col].nunique(dropna=True)
    if category_count <= min(TOP_N_OPTIONS):
        return data
    control_signature = dataframe_signature(data)
    control_key = f"{key}_{control_signature}"
    default_index = 0
    mode = st.radio(
        label,
        ["Highest", "Lowest"],
        index=0,
        horizontal=False,
        key=f"chart_rank_mode_{control_key}",
        help="This limits the chart only. The table above remains complete.",
    )
    selected_n = st.radio(
        "Number of categories",
        TOP_N_OPTIONS,
        index=default_index,
        horizontal=False,
        key=f"chart_top_n_{control_key}",
    )
    totals = data.groupby(category_col, observed=False)["Count"].sum().sort_values(ascending=(mode == "Lowest"))
    keep = totals.head(selected_n).index
    return data[data[category_col].isin(keep)].copy()


def add_grand_total_row(df, label_col):
    if df.empty:
        return df
    total = {c: "" for c in df.columns}
    total[label_col] = "Grand Total"
    total["Count"] = df["Count"].sum()
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


def top_n_with_total(df, n, label_col):
    if df.empty:
        return df
    total = {c: "" for c in df.columns}
    total[label_col] = "Grand Total"
    total["Count"] = df["Count"].sum()
    return pd.concat([df.head(n), pd.DataFrame([total])], ignore_index=True)


def style_simple_total(df):
    if df.empty:
        return df

    def row_style(row):
        if "Grand Total" in [str(x) for x in row.values]:
            return ["font-weight: 800; background-color: #dbeafe; color: #0f172a;" for _ in row]
        return ["" for _ in row]

    return (
        df.style
        .apply(row_style, axis=1)
        .format({"Count": "{:,.0f}"})
        .set_table_styles(
            [
                {"selector": "th", "props": [("background-color", "#eaf2fb"), ("color", "#111827"), ("font-weight", "800"), ("padding", "9px 10px"), ("border", "1px solid #d5dee9")]},
                {"selector": "td", "props": [("color", "#1f2937"), ("padding", "8px 10px"), ("border", "1px solid #e5edf5")]},
                {"selector": "tbody tr:nth-child(even)", "props": [("background-color", "#f8fafc")]},
            ]
        )
    )


def file_slug(value):
    value = re.sub(r"[^a-zA-Z0-9]+", "_", str(value)).strip("_").lower()
    return value[:90] or "dashboard_export"


def flatten_table_for_export(table):
    flat = table.copy()
    if isinstance(flat.index, pd.MultiIndex):
        flat.index = [" | ".join(str(part) for part in idx if str(part) != "") for idx in flat.index]
    if isinstance(flat.columns, pd.MultiIndex):
        flat.columns = [" | ".join(str(part) for part in col if str(part) != "") for col in flat.columns]
    flat = flat.reset_index()
    flat.columns = [str(col) if str(col) != "index" else "Category" for col in flat.columns]
    return flat


def dataframe_signature(table):
    if table.empty:
        return "empty"
    flat = flatten_table_for_export(table).astype("string").fillna("")
    metadata = f"{flat.shape}|{'|'.join(flat.columns)}".encode("utf-8", errors="ignore")
    hashed_values = pd.util.hash_pandas_object(flat, index=True).values.tobytes()
    return hashlib.md5(metadata + hashed_values).hexdigest()[:12]


def chart_signature(fig):
    return hashlib.md5(fig.to_json().encode("utf-8", errors="ignore")).hexdigest()[:12]


def dataframe_to_png_bytes(table, title):
    try:
        import matplotlib.pyplot as plt
    except Exception:
        return None, "Image export for tables requires matplotlib."
    if table.empty:
        return None, "No table data available to export."

    flat = flatten_table_for_export(table)
    row_count = len(flat)
    col_count = len(flat.columns)
    width = min(24, max(9, col_count * 1.45))
    height = min(40, max(3.2, row_count * 0.34 + 1.7))

    fig, ax = plt.subplots(figsize=(width, height))
    ax.axis("off")
    ax.set_title(title, fontsize=14, fontweight="bold", loc="left", pad=12)
    table_artist = ax.table(
        cellText=flat.astype(str).values,
        colLabels=flat.columns,
        loc="upper left",
        cellLoc="left",
        colLoc="left",
    )
    table_artist.auto_set_font_size(False)
    table_artist.set_fontsize(8)
    table_artist.scale(1, 1.25)
    for (row, col), cell in table_artist.get_celld().items():
        cell.set_edgecolor("#d5dee9")
        if row == 0:
            cell.set_facecolor("#eaf2fb")
            cell.set_text_props(weight="bold", color="#111827")
        elif any(str(value) == "Grand Total" for value in flat.iloc[row - 1].values):
            cell.set_facecolor("#dbeafe")
            cell.set_text_props(weight="bold", color="#0f172a")
        elif row % 2 == 0:
            cell.set_facecolor("#f8fafc")
        else:
            cell.set_facecolor("#ffffff")
    buffer = BytesIO()
    fig.savefig(buffer, format="png", dpi=180, bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer.getvalue(), None


def table_image_download(table, key, title):
    if table.empty:
        return
    current_signature = dataframe_signature(table)
    scoped_key = f"{key}_{current_signature}"
    image_key = f"table_png_{scoped_key}"
    message_key = f"table_png_message_{scoped_key}"
    st.markdown("<div class='export-row'>", unsafe_allow_html=True)
    prep_col, download_col, csv_col = st.columns([1, 1, 1])
    with prep_col:
        if st.button("Generate table PNG", key=f"prepare_table_{scoped_key}", use_container_width=True):
            image, message = dataframe_to_png_bytes(table, title)
            st.session_state[image_key] = image
            st.session_state[message_key] = message
    with download_col:
        if st.session_state.get(image_key):
            st.download_button(
                "Download table PNG",
                data=st.session_state[image_key],
                file_name=f"{file_slug(title)}.png",
                mime="image/png",
                key=f"download_table_{scoped_key}",
                use_container_width=True,
            )
        else:
            st.button("Download table PNG", key=f"download_table_placeholder_{scoped_key}", disabled=True, use_container_width=True)
    with csv_col:
        st.download_button(
            "Download table CSV",
            data=flatten_table_for_export(table).to_csv(index=False).encode("utf-8"),
            file_name=f"{file_slug(title)}.csv",
            mime="text/csv",
            key=f"download_table_csv_{scoped_key}",
            use_container_width=True,
        )
    if st.session_state.get(message_key):
        st.caption(st.session_state[message_key])
    st.markdown("</div>", unsafe_allow_html=True)


def render_table(table, key, title, simple=False):
    st.dataframe(style_simple_total(table) if simple else style_total(table), use_container_width=True)
    table_image_download(table, key, title)


def chart_image_download(fig, key, title):
    current_signature = chart_signature(fig)
    scoped_key = f"{key}_{current_signature}"
    image_key = f"chart_png_{scoped_key}"
    message_key = f"chart_png_message_{scoped_key}"
    st.markdown("<div class='export-row'>", unsafe_allow_html=True)
    prep_col, download_col = st.columns([1, 1])
    with prep_col:
        if st.button("Generate chart PNG", key=f"prepare_chart_{scoped_key}", use_container_width=True):
            try:
                st.session_state[image_key] = fig.to_image(format="png", scale=2)
                st.session_state[message_key] = None
            except Exception:
                st.session_state[image_key] = None
                st.session_state[message_key] = "Chart image export requires the kaleido package in the Streamlit environment."
    with download_col:
        if st.session_state.get(image_key):
            st.download_button(
                "Download chart PNG",
                data=st.session_state[image_key],
                file_name=f"{file_slug(title)}.png",
                mime="image/png",
                key=f"download_chart_{scoped_key}",
                use_container_width=True,
            )
        else:
            st.button("Download chart PNG", key=f"download_chart_placeholder_{scoped_key}", disabled=True, use_container_width=True)
    if st.session_state.get(message_key):
        st.caption(st.session_state[message_key])
    st.markdown("</div>", unsafe_allow_html=True)


def multi_response_tables(long_data, value_col, label):
    if long_data.empty or value_col not in long_data.columns:
        empty_distribution = pd.DataFrame(columns=["Selection count", "Count"])
        empty_combinations = pd.DataFrame(columns=[f"{label} combination", "Count"])
        return empty_distribution, empty_combinations

    cleaned = (
        long_data[["record_id", value_col]]
        .dropna()
        .drop_duplicates()
        .copy()
    )
    cleaned = cleaned[~cleaned[value_col].astype("string").isin([MISSING, REVIEW])]
    if cleaned.empty:
        empty_distribution = pd.DataFrame(columns=["Selection count", "Count"])
        empty_combinations = pd.DataFrame(columns=[f"{label} combination", "Count"])
        return empty_distribution, empty_combinations

    per_record = cleaned.groupby("record_id", observed=False)[value_col].agg(
        lambda values: sorted(set(str(value) for value in values if str(value).strip()))
    )
    selection_counts = per_record.map(len)
    distribution = (
        selection_counts.value_counts()
        .sort_index()
        .rename_axis("Selection count")
        .reset_index(name="Count")
    )
    distribution["Selection count"] = distribution["Selection count"].map(
        lambda count: "1 selection" if count == 1 else f"{count} selections"
    )
    distribution = add_grand_total_row(distribution, "Selection count")

    combinations = (
        per_record.map(lambda values: " + ".join(values))
        .value_counts()
        .rename_axis(f"{label} combination")
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )
    combinations = top_n_with_total(combinations, 30, f"{label} combination")
    return distribution, combinations


def bar_chart(data, x, y, color=None, title="", horizontal=False, height=420, category_orders=None):
    if data.empty:
        st.info("No records available for this chart.")
        return
    chart_data = data.copy()
    y_order = []
    auto_orders = {
        "age_group": AGE_GROUP_ORDER,
        "gender_clean": GENDER_ORDER,
        "first_visit_clean": YES_NO_ORDER,
    }
    if horizontal:
        chart_data, y_order = sort_chart_by_total(chart_data, y, keep_age_order=(y == "age_group"), axis="y")
        if y_order:
            auto_orders[y] = y_order
    elif x != "age_group" and "Count" in chart_data.columns:
        chart_data, x_order = sort_chart_by_total(chart_data, x)
        if x_order:
            auto_orders[x] = x_order
    if category_orders:
        auto_orders.update(category_orders)
    fig = px.bar(
        chart_data,
        x=x,
        y=y,
        color=color,
        text="Count" if "Count" in data.columns else None,
        orientation="h" if horizontal else "v",
        title=title,
        category_orders=auto_orders,
        color_discrete_sequence=px.colors.qualitative.Set2,
    )
    fig.update_traces(texttemplate="%{text:,}", textposition="outside", cliponaxis=False)
    fig.update_layout(
        height=height,
        margin=dict(l=12, r=24, t=58, b=18),
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend_title_text="",
        title_font=dict(size=18),
        font=dict(size=13, color="#111827"),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#e9eef5", zeroline=False)
    fig.update_yaxes(showgrid=False, zeroline=False)
    if horizontal and y_order:
        fig.update_yaxes(categoryorder="array", categoryarray=y_order)
    chart_key = file_slug(title)
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={
            "displayModeBar": True,
            "responsive": True,
            "toImageButtonOptions": {"format": "png", "filename": chart_key, "scale": 2},
        },
    )
    chart_image_download(fig, chart_key, title)


def pie_chart(data, names, values, title=""):
    if data.empty:
        st.info("No records available for this chart.")
        return
    fig = px.pie(
        data,
        names=names,
        values=values,
        hole=0.42,
        title=title,
        category_orders={names: YES_NO_ORDER},
        color_discrete_sequence=px.colors.qualitative.Set2,
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    fig.update_layout(
        height=420,
        margin=dict(l=12, r=24, t=58, b=18),
        paper_bgcolor="white",
        title_font=dict(size=18),
        font=dict(size=13, color="#111827"),
        legend_title_text="",
    )
    chart_key = file_slug(title)
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={
            "displayModeBar": True,
            "responsive": True,
            "toImageButtonOptions": {"format": "png", "filename": chart_key, "scale": 2},
        },
    )
    chart_image_download(fig, chart_key, title)


def sliced_bar_chart(data, category_col, slicer_key, x, y, color=None, title="", horizontal=False, height=420, category_orders=None, default=10):
    if data.empty or category_col not in data.columns or data[category_col].nunique(dropna=True) <= min(TOP_N_OPTIONS):
        bar_chart(data, x, y, color, title, horizontal, height, category_orders)
        return
    control_col, chart_col = st.columns([0.23, 0.77])
    with control_col:
        st.markdown("<div class='chart-control-title'>Chart controls</div>", unsafe_allow_html=True)
        chart_data = top_n_chart_control(data, category_col, slicer_key, default=default)
    with chart_col:
        bar_chart(chart_data, x, y, color, title, horizontal, height, category_orders)



def section_note(text):
    st.caption(text)


def slug(value):
    value = re.sub(r"[^a-zA-Z0-9]+", "_", str(value)).strip("_").lower()
    return value[:80] or "blank"


def scope_switch(key):
    scope_key = f"filter_{key}_scope"
    if scope_key not in st.session_state:
        st.session_state[scope_key] = "All"
    options = ["All", "Custom"]
    if hasattr(st, "pills"):
        try:
            selected_scope = st.pills(
                "Scope",
                options,
                selection_mode="single",
                key=scope_key,
                label_visibility="collapsed",
            )
            return selected_scope or "All"
        except Exception:
            pass
    if hasattr(st, "segmented_control"):
        try:
            selected_scope = st.segmented_control(
                "Scope",
                options,
                key=scope_key,
                label_visibility="collapsed",
            )
            return selected_scope or "All"
        except Exception:
            pass
    return st.radio(
        "Scope",
        options,
        horizontal=True,
        key=scope_key,
        label_visibility="collapsed",
    )


def value_selector(options, key, searchable_threshold=30):
    selected_key = f"filter_{key}_selected"
    if selected_key not in st.session_state:
        st.session_state[selected_key] = []
    st.session_state[selected_key] = [v for v in st.session_state[selected_key] if v in options]

    if hasattr(st, "pills") and len(options) <= searchable_threshold:
        try:
            selected = st.pills(
                "Choose values",
                options,
                selection_mode="multi",
                key=selected_key,
                label_visibility="collapsed",
            )
            return selected or []
        except Exception:
            pass

    st.caption("Search and choose one or more values.")
    return st.multiselect(
        "Choose values",
        options,
        key=selected_key,
        placeholder="Type to search",
        label_visibility="collapsed",
    )


def multi_choice_selector(label, options, key, disabled=False, disabled_message=None, order=None):
    options = ordered_unique(pd.Series(options), order=order)
    with st.sidebar.expander(label, expanded=not disabled):
        if disabled:
            st.caption(disabled_message or "Complete the previous filter first.")
            return [], False
        if not options:
            st.caption("No options available.")
            return [], False

        st.caption(f"{len(options):,} available")
        selected_scope = scope_switch(key)
        if selected_scope == "All":
            st.caption("All available values are included for this level.")
            return [], False

        selected = value_selector(options, key)
        if not selected:
            st.caption("Select at least one value to continue.")
        else:
            st.caption(f"{len(selected):,} selected")
        return selected, True


def apply_filter(data, column, selected, explicit):
    if not explicit:
        return data.copy()
    if not selected:
        return data.iloc[0:0].copy()
    return data[data[column].isin(selected)].copy()


def format_selection(selected, explicit, all_label):
    if not explicit:
        return all_label
    if not selected:
        return "None selected"
    return ", ".join(map(str, selected))


def metric_card(label, value, helper=None):
    helper_html = f"<span>{helper}</span>" if helper else ""
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-helper">{helper_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def insight_card(label, value, helper=None):
    helper_html = f"<div class='insight-helper'>{helper}</div>" if helper else ""
    st.markdown(
        f"""
        <div class="insight-card">
            <div class="insight-label">{label}</div>
            <div class="insight-value">{value}</div>
            {helper_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def top_category(data, column, exclude=None):
    if data.empty or column not in data.columns:
        return "N/A", 0
    exclude = set(exclude or [])
    series = data[column].astype("string")
    series = series[~series.isin(exclude)]
    counts = series.value_counts(dropna=True)
    if counts.empty:
        return "N/A", 0
    return str(counts.index[0]), int(counts.iloc[0])


def yes_count(data, column):
    if data.empty or column not in data.columns:
        return 0
    return int(data[column].astype("string").eq("Yes").sum())


def render_badges(items):
    if not items:
        return
    badge_html = "".join(f"<span class='status-badge status-{kind}'>{label}</span>" for label, kind in items)
    st.markdown(f"<div class='badge-row'>{badge_html}</div>", unsafe_allow_html=True)


def render_filter_chips(items):
    chips = []
    for label, value, active in items:
        css_class = "filter-chip active" if active else "filter-chip"
        chips.append(f"<span class='{css_class}'><strong>{label}</strong>: {value}</span>")
    st.markdown(f"<div class='filter-chip-row'>{''.join(chips)}</div>", unsafe_allow_html=True)


def category_share_insight(data, category_col, subject, value_col="Count"):
    if data.empty or category_col not in data.columns or value_col not in data.columns:
        return f"No {subject.lower()} are available for the current filters."
    totals = data.groupby(category_col, observed=False)[value_col].sum().sort_values(ascending=False)
    totals = totals[totals > 0]
    if totals.empty:
        return f"No {subject.lower()} are available for the current filters."
    top_label = str(totals.index[0])
    top_value = int(totals.iloc[0])
    total_value = int(totals.sum())
    share = top_value / total_value if total_value else 0
    return f"{top_label} is the leading {subject.lower()}, accounting for {share:.1%} of {total_value:,} records in the current filter."


def render_note(text):
    if text:
        st.markdown(f"<div class='insight-note'>{text}</div>", unsafe_allow_html=True)


def render_analysis_block(title, table, chart_renderer, note, key, simple=False, badges=None):
    st.markdown(f"<div class='analysis-title'>{title}</div>", unsafe_allow_html=True)
    render_badges(badges or [])
    chart_tab, table_tab, notes_tab = st.tabs(["Chart", "Table", "Notes"])
    with chart_tab:
        chart_renderer()
    with table_tab:
        render_table(table, f"{key}_table", title, simple=simple)
    with notes_tab:
        render_note(note)


def dashboard_nav(options, key="dashboard_section"):
    if st.session_state.get(key) not in options:
        st.session_state[key] = options[0]
    if hasattr(st, "pills"):
        try:
            selected = st.pills(
                "Dashboard section",
                options,
                selection_mode="single",
                key=key,
                label_visibility="collapsed",
            )
            return selected or options[0]
        except Exception:
            pass
    if hasattr(st, "segmented_control"):
        try:
            selected = st.segmented_control(
                "Dashboard section",
                options,
                key=key,
                label_visibility="collapsed",
            )
            return selected or options[0]
        except Exception:
            pass
    return st.radio(
        "Dashboard section",
        options,
        index=options.index(st.session_state.get(key, options[0])),
        key=key,
        horizontal=True,
        label_visibility="collapsed",
    )


def load_external_css(path):
    if not path.exists():
        return
    try:
        css = path.read_text(encoding="utf-8")
    except Exception as exc:
        st.caption(f"Custom CSS could not be loaded: {exc}")
        return
    if css.strip():
        st.markdown(f"<style>\n{css}\n</style>", unsafe_allow_html=True)


st.set_page_config(
    page_title="Tdh Kenya CFS Dashboard",
    page_icon="CFS",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root {
        --ink: #111827;
        --muted: #475569;
        --line: #cbd5e1;
        --panel: #ffffff;
        --soft: #f5f8fc;
        --brand: #0f6fbd;
        --brand-dark: #094f8d;
    }
    html, body, [class*="css"] {
        color: var(--ink);
    }
    p, span, label, div, .stMarkdown, .stCaption, [data-testid="stMarkdownContainer"] {
        color: var(--ink);
    }
    .main .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
        max-width: 1480px;
    }
    h1, h2, h3 {
        color: var(--ink);
        letter-spacing: 0;
    }
    section[data-testid="stSidebar"] {
        background: #f8fafc;
        border-right: 1px solid var(--line);
    }
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span {
        color: #1f2937 !important;
    }
    div[data-testid="stExpander"] {
        border: 1px solid var(--line);
        border-radius: 8px;
        background: #ffffff;
    }
    .dashboard-title {
        padding: 1rem 1.2rem;
        border: 1px solid var(--line);
        border-radius: 8px;
        background: linear-gradient(90deg, #ffffff 0%, #e8f2ff 100%);
        margin-bottom: 1rem;
    }
    .dashboard-title h1 {
        margin: 0;
        font-size: 2rem;
        line-height: 1.2;
    }
    .dashboard-title p {
        margin: .35rem 0 0 0;
        color: var(--muted);
        font-size: .98rem;
    }
    .record-status {
        display: inline-flex;
        align-items: center;
        gap: .45rem;
        padding: .65rem .9rem;
        margin: -.35rem 0 1rem 0;
        background: #eef7f2;
        border: 1px solid #cce6d8;
        border-radius: 8px;
        color: #1f5135;
        font-weight: 700;
        font-size: .95rem;
    }
    .filter-summary {
        padding: .85rem 1rem;
        background: #f1f7fd;
        border: 1px solid #cfe0f5;
        border-radius: 8px;
        color: #213547;
        margin-bottom: 1rem;
        font-size: .95rem;
    }
    .filter-group-title {
        margin: .9rem 0 .35rem 0;
        padding: .45rem .65rem;
        border-left: 4px solid var(--brand);
        background: #eef6ff;
        color: #0f172a;
        font-weight: 900;
        border-radius: 6px;
    }
    .filter-chip-row {
        display: flex;
        flex-wrap: wrap;
        gap: .45rem;
        margin: -.25rem 0 1rem 0;
    }
    .filter-chip {
        display: inline-flex;
        align-items: center;
        gap: .25rem;
        padding: .45rem .65rem;
        border-radius: 999px;
        border: 1px solid #d5dee9;
        background: #ffffff;
        color: #334155;
        font-size: .85rem;
        font-weight: 700;
    }
    .filter-chip.active {
        border-color: #8fc3f2;
        background: #eaf5ff;
        color: #0b4f84;
    }
    .badge-row {
        display: flex;
        flex-wrap: wrap;
        gap: .4rem;
        margin: .25rem 0 .65rem 0;
    }
    .status-badge {
        display: inline-flex;
        align-items: center;
        padding: .32rem .55rem;
        border-radius: 999px;
        font-size: .76rem;
        font-weight: 900;
        border: 1px solid #d5dee9;
    }
    .status-good {
        background: #ecfdf3;
        border-color: #b7e4c7;
        color: #166534;
    }
    .status-review {
        background: #fff7ed;
        border-color: #fed7aa;
        color: #9a3412;
    }
    .status-harmonized {
        background: #eef6ff;
        border-color: #bfdbfe;
        color: #1d4ed8;
    }
    .status-missing {
        background: #f8fafc;
        border-color: #cbd5e1;
        color: #475569;
    }
    .analysis-title {
        margin: 1.15rem 0 .2rem 0;
        padding-top: .35rem;
        color: #0f172a;
        font-size: 1.2rem;
        line-height: 1.3;
        font-weight: 900;
    }
    .chart-control-title {
        margin: .45rem 0 .65rem 0;
        padding: .55rem .65rem;
        border-radius: 8px;
        border: 1px solid #d5dee9;
        background: #f8fbff;
        color: #0f172a;
        font-size: .9rem;
        font-weight: 900;
        text-align: center;
    }
    .insight-note {
        padding: .9rem 1rem;
        border-radius: 8px;
        border: 1px solid #d5dee9;
        background: #fbfdff;
        color: #1f2937;
        font-weight: 700;
    }
    .export-row {
        margin: .35rem 0 1rem 0;
    }
    .metric-card {
        border: 1px solid var(--line);
        background: var(--panel);
        border-radius: 8px;
        padding: 1rem;
        min-height: 112px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, .04);
    }
    .metric-label {
        color: #334155;
        font-size: .84rem;
        font-weight: 700;
        text-transform: uppercase;
    }
    .metric-value {
        color: var(--brand-dark);
        font-size: 1.72rem;
        font-weight: 800;
        margin-top: .25rem;
    }
    .metric-helper {
        color: #475569;
        font-size: .78rem;
        margin-top: .15rem;
    }
    .quick-insights {
        margin: .6rem 0 1.1rem 0;
        padding: 1.05rem;
        border: 1px solid #d5dee9;
        border-radius: 8px;
        background: #f8fbff;
        box-shadow: 0 2px 8px rgba(15, 23, 42, .06);
    }
    .section-heading {
        margin: 1.25rem 0 .55rem 0;
        color: #0f172a;
        font-size: 1.55rem;
        line-height: 1.25;
        font-weight: 900;
    }
    .dashboard-section-heading {
        font-size: 1.75rem;
        letter-spacing: 0;
        margin-top: 1.35rem;
    }
    .section-subtitle {
        margin: -.25rem 0 .8rem 0;
        color: #334155;
        font-size: 1rem;
        font-weight: 700;
    }
    .insight-card {
        min-height: 104px;
        padding: .95rem 1rem;
        border: 1px solid #d5dee9;
        border-radius: 8px;
        background: #ffffff;
        box-shadow: 0 1px 3px rgba(15, 23, 42, .05);
    }
    .insight-label {
        color: #334155;
        font-size: .78rem;
        font-weight: 800;
        text-transform: uppercase;
    }
    .insight-value {
        color: #0f172a;
        font-size: 1.28rem;
        font-weight: 900;
        margin-top: .25rem;
        line-height: 1.25;
    }
    .insight-helper {
        color: #475569;
        font-size: .78rem;
        margin-top: .2rem;
    }
    div[data-testid="stTabs"] button p {
        color: #111827 !important;
        font-weight: 850 !important;
        font-size: .98rem !important;
    }
    .section-block {
        padding: .9rem 0 .25rem 0;
        border-top: 1px solid var(--line);
        margin-top: .8rem;
    }
    div[data-testid="stDataFrame"] {
        border: 1px solid #d5dee9;
        border-radius: 8px;
        overflow: hidden;
        background: #ffffff;
    }
    .dataframe th {
        font-weight: 800 !important;
        color: #111827 !important;
        background: #eaf2fb !important;
        padding: 9px 10px !important;
    }
    .dataframe td {
        color: #1f2937 !important;
        padding: 8px 10px !important;
        border-color: #e5edf5 !important;
    }
    div[data-testid="stDataFrame"] [role="gridcell"],
    div[data-testid="stDataFrame"] [role="columnheader"] {
        color: #1f2937 !important;
        font-size: 13px !important;
    }
    button[kind="primary"] {
        background: var(--brand) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

load_external_css(CSS_PATH)

title_col, logo_col = st.columns([5, 1])
with title_col:
    st.markdown(
        """
        <div class="dashboard-title">
            <h1>Tdh Kenya Child Friendly Spaces Dashboard</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )
with logo_col:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), use_container_width=True)
record_status_slot = st.empty()

if not DATA_PATH.exists():
    st.error(f"Data file not found: {DATA_PATH}")
    st.stop()

if st.sidebar.button("Refresh data and clear cache", use_container_width=True):
    st.cache_data.clear()
    try:
        if PREPARED_CACHE_PATH.exists():
            PREPARED_CACHE_PATH.unlink()
    except Exception:
        pass
    for state_key in list(st.session_state.keys()):
        if state_key.startswith("filter_") or state_key in {"date_filter", "date_from_filter", "date_to_filter"}:
            st.session_state.pop(state_key, None)
    st.rerun()

modified_time = DATA_PATH.stat().st_mtime
df, issue_long, support_long, game_long = load_and_prepare(modified_time)

st.sidebar.header("Filters")
reset_col, refresh_col = st.sidebar.columns(2)
with reset_col:
    if st.button("Reset filters", use_container_width=True):
        for state_key in list(st.session_state.keys()):
            if state_key.startswith("filter_") or state_key in {"date_filter", "date_from_filter", "date_to_filter"}:
                st.session_state.pop(state_key, None)
        st.rerun()
with refresh_col:
    st.caption("Reset keeps the prepared data cache.")
st.sidebar.caption("Use each level in order. Later filters unlock after the location path is defined.")
st.sidebar.markdown("<div class='filter-group-title'>Date & Location Filters</div>", unsafe_allow_html=True)
valid_dates = df["date"].dropna()
if valid_dates.empty:
    st.sidebar.warning("No valid dates found in the data.")
    date_scope = df.copy()
    start_date = end_date = None
else:
    min_date = valid_dates.min().date()
    max_date = valid_dates.max().date()
    st.sidebar.markdown("**1. Date range**")
    from_col, to_col = st.sidebar.columns(2)
    with from_col:
        from_date = st.date_input("From:", value=min_date, min_value=min_date, max_value=max_date, key="date_from_filter")
    with to_col:
        to_date = st.date_input("To:", value=max_date, min_value=min_date, max_value=max_date, key="date_to_filter")
    if from_date > to_date:
        st.sidebar.warning("From date is after To date; dates were swapped for this view.")
        from_date, to_date = to_date, from_date
    start_date, end_date = pd.to_datetime(from_date), pd.to_datetime(to_date)
    date_scope = df[df["date"].between(start_date, end_date, inclusive="both")].copy()

selected_settlements, settlement_explicit = multi_choice_selector(
    "2. Camp Name",
    date_scope["settlement_clean"],
    "settlement",
)
settlement_scope = apply_filter(date_scope, "settlement_clean", selected_settlements, settlement_explicit)

location_disabled = not settlement_explicit or not selected_settlements or settlement_scope.empty
selected_locations, location_explicit = multi_choice_selector(
    "3. Specific camp location",
    settlement_scope["location_clean"] if not settlement_scope.empty else [],
    "location",
    disabled=location_disabled,
    disabled_message="Set Camp Name to Custom, then choose at least one camp first.",
)
location_scope = apply_filter(settlement_scope, "location_clean", selected_locations, location_explicit)

cfs_disabled = location_disabled or not location_explicit or not selected_locations or location_scope.empty
selected_cfs, cfs_explicit = multi_choice_selector(
    "4. CFS / site",
    location_scope["cfs_clean"] if not location_scope.empty else [],
    "cfs",
    disabled=cfs_disabled,
    disabled_message="Choose at least one Specific camp location first.",
)
cfs_scope = apply_filter(location_scope, "cfs_clean", selected_cfs, cfs_explicit)

st.sidebar.markdown("<div class='filter-group-title'>Operational Filters</div>", unsafe_allow_html=True)
downstream_disabled = cfs_disabled or not cfs_explicit or not selected_cfs or cfs_scope.empty
selected_staff, staff_explicit = multi_choice_selector(
    "5. Staff / CPV",
    cfs_scope["staff_clean"] if not cfs_scope.empty else [],
    "staff",
    disabled=downstream_disabled,
    disabled_message="Choose at least one CFS / site first.",
)
staff_scope = apply_filter(cfs_scope, "staff_clean", selected_staff, staff_explicit)

st.sidebar.markdown("<div class='filter-group-title'>Beneficiary Filters</div>", unsafe_allow_html=True)
selected_gender, gender_explicit = multi_choice_selector(
    "6. Gender",
    staff_scope["gender_clean"] if not staff_scope.empty else [],
    "gender",
    disabled=downstream_disabled,
    disabled_message="Choose at least one CFS / site first.",
    order=GENDER_ORDER,
)
gender_scope = apply_filter(staff_scope, "gender_clean", selected_gender, gender_explicit)

selected_age, age_explicit = multi_choice_selector(
    "7. Age group",
    gender_scope["age_group"].astype("string") if not gender_scope.empty else [],
    "age_group",
    disabled=downstream_disabled,
    disabled_message="Choose at least one CFS / site first.",
    order=AGE_GROUP_ORDER,
)
age_scope = apply_filter(gender_scope.assign(age_group=gender_scope["age_group"].astype("string")), "age_group", selected_age, age_explicit)

selected_disability, disability_explicit = multi_choice_selector(
    "8. Living with disability",
    age_scope["disability_status_clean"].astype("string") if not age_scope.empty else [],
    "disability",
    disabled=downstream_disabled,
    disabled_message="Choose at least one CFS / site first.",
    order=YES_NO_ORDER,
)
filtered = apply_filter(
    age_scope.assign(disability_status_clean=age_scope["disability_status_clean"].astype("string")),
    "disability_status_clean",
    selected_disability,
    disability_explicit,
)
filtered["age_group"] = pd.Categorical(filtered["age_group"], categories=AGE_GROUP_ORDER, ordered=True)
filtered["gender_clean"] = pd.Categorical(filtered["gender_clean"], categories=GENDER_ORDER, ordered=True)
filtered["first_visit_clean"] = pd.Categorical(filtered["first_visit_clean"], categories=YES_NO_ORDER, ordered=True)
filtered["disability_status_clean"] = pd.Categorical(filtered["disability_status_clean"], categories=YES_NO_ORDER, ordered=True)

context_columns = ["record_id", "settlement_clean", "location_clean", "cfs_clean", "staff_clean", "gender_clean", "age_group"]
issue_context = issue_long.merge(filtered[context_columns], on="record_id", how="inner")
support_context = support_long.merge(filtered[context_columns], on="record_id", how="inner")
game_context = game_long.merge(filtered[context_columns], on="record_id", how="inner")

source_modified = time.strftime("%d %b %Y %H:%M", time.localtime(modified_time))
st.caption(f"Last modified: {source_modified}")

if start_date is not None and end_date is not None:
    start_label = start_date.strftime("%d %b %Y")
    end_label = end_date.strftime("%d %b %Y")
    date_selected_label = start_label if start_label == end_label else f"{start_label} to {end_label}"
else:
    date_selected_label = "All available dates"
record_status_slot.markdown(
    f"<div class='record-status'>Showing {len(filtered):,} of {len(df):,} records | Date selected: {date_selected_label}</div>",
    unsafe_allow_html=True,
)

render_filter_chips(
    [
        ("Camp", format_selection(selected_settlements, settlement_explicit, "All camps"), settlement_explicit),
        ("Specific location", format_selection(selected_locations, location_explicit, "All locations"), location_explicit),
        ("CFS / site", format_selection(selected_cfs, cfs_explicit, "All sites"), cfs_explicit),
        ("Staff / CPV", format_selection(selected_staff, staff_explicit, "All staff"), staff_explicit),
        ("Gender", format_selection(selected_gender, gender_explicit, "All genders"), gender_explicit),
        ("Age group", format_selection(selected_age, age_explicit, "All age groups"), age_explicit),
        ("Disability", format_selection(selected_disability, disability_explicit, "All statuses"), disability_explicit),
    ]
)

summary_bits = [
    f"Camp Name: {format_selection(selected_settlements, settlement_explicit, 'All camps')}",
    f"Specific camp location: {format_selection(selected_locations, location_explicit, 'All specific camp locations')}",
    f"CFS / site: {format_selection(selected_cfs, cfs_explicit, 'All CFS / sites')}",
    f"Staff: {format_selection(selected_staff, staff_explicit, 'All staff')}",
    f"Gender: {format_selection(selected_gender, gender_explicit, 'All genders')}",
    f"Age group: {format_selection(selected_age, age_explicit, 'All age groups')}",
    f"Disability: {format_selection(selected_disability, disability_explicit, 'All disability statuses')}",
]
st.sidebar.markdown(
    f"<div class='filter-summary'><strong>Current filter path:</strong><br>{'<br>'.join(summary_bits)}</div>",
    unsafe_allow_html=True,
)

export_cols = [
    "record_id", "date", "month", "consent_clean", "staff_clean", "child_name",
    "gender_clean", "age_clean", "age_group", "disability_status_clean",
    "disability_type_clean", "disability_type_display", "settlement_clean",
    "location_clean", "cfs_clean", "games_played_clean", "issues_combined",
    "support_combined", "first_visit_clean", "referral_made_clean",
    "referral_destination_grouped", "external_referral_agency_clean",
]
download_df = filtered.copy()
for col in export_cols:
    if col in download_df.columns and str(download_df[col].dtype) == "category":
        download_df[col] = download_df[col].astype("string")


referral_rate = filtered["referral_made_clean"].astype("string").eq("Yes").mean()
disability_rate = filtered["disability_status_clean"].astype("string").eq("Yes").mean()
first_visit_rate = filtered["first_visit_clean"].astype("string").eq("Yes").mean()
avg_age = filtered["age_clean"].mean()

k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    metric_card("Total CFS visits", f"{len(filtered):,}", "Filtered records")
with k2:
    metric_card("Referral rate", f"{referral_rate:.1%}", "Referral marked Yes")
with k3:
    metric_card("Disability prevalence", f"{disability_rate:.1%}", "Children recorded as Yes")
with k4:
    metric_card("First visit rate", f"{first_visit_rate:.1%}", "First visit marked Yes")
with k5:
    metric_card("Average child age", f"{avg_age:.1f}" if pd.notna(avg_age) else "N/A", "Years")

top_camp, top_camp_count = top_category(filtered, "settlement_clean", exclude=[MISSING])
top_site, top_site_count = top_category(filtered, "cfs_clean", exclude=[MISSING])
top_gender, top_gender_count = top_category(filtered, "gender_clean", exclude=[MISSING])
top_age, top_age_count = top_category(filtered, "age_group", exclude=[MISSING])
top_issue, top_issue_count = top_category(issue_context, "issue_clean", exclude=[MISSING])
top_support, top_support_count = top_category(support_context, "support_clean", exclude=[MISSING])
top_game, top_game_count = top_category(filtered, "games_played_clean", exclude=[MISSING])
top_referral_dest, top_referral_dest_count = top_category(
    filtered[filtered["referral_made_clean"].astype("string").eq("Yes")],
    "referral_destination_grouped",
    exclude=[MISSING],
)

st.markdown("<div class='section-heading'>Quick Insights</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='section-subtitle'>Fast outlook for the currently filtered dataset.</div>",
    unsafe_allow_html=True,
)
try:
    quick_panel = st.container(border=True)
except TypeError:
    quick_panel = st.container()
with quick_panel:
    qi1, qi2, qi3, qi4 = st.tabs(["Coverage", "Demographics", "Protection & Support", "Engagement & Referrals"])
    with qi1:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            insight_card("Camp with most records", top_camp, f"{top_camp_count:,} records")
        with c2:
            insight_card("Leading CFS / site", top_site, f"{top_site_count:,} records")
        with c3:
            insight_card("CFS / sites represented", f"{filtered['cfs_clean'].nunique():,}", "Filtered data")
        with c4:
            insight_card("Staff represented", f"{filtered['staff_clean'].nunique():,}", "Filtered data")
    with qi2:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            insight_card("Largest gender group", top_gender, f"{top_gender_count:,} records")
        with c2:
            insight_card("Largest age group", top_age, f"{top_age_count:,} records")
        with c3:
            insight_card("Children with disability", f"{yes_count(filtered, 'disability_status_clean'):,}", f"{disability_rate:.1%} of filtered records")
        with c4:
            insight_card("First-time visitors", f"{yes_count(filtered, 'first_visit_clean'):,}", f"{first_visit_rate:.1%} of filtered records")
    with qi3:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            insight_card("Most reported issue", top_issue, f"{top_issue_count:,} mentions")
        with c2:
            insight_card("Most common support", top_support, f"{top_support_count:,} mentions")
        with c3:
            insight_card("Issue mentions", f"{len(issue_context):,}", "Multi-response count")
        with c4:
            insight_card("Support mentions", f"{len(support_context):,}", "Multi-response count")
    with qi4:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            insight_card("Top activity / game", top_game, f"{top_game_count:,} records")
        with c2:
            insight_card("Referrals made", f"{yes_count(filtered, 'referral_made_clean'):,}", f"{referral_rate:.1%} of filtered records")
        with c3:
            insight_card("Top referral destination", top_referral_dest, f"{top_referral_dest_count:,} records")
        with c4:
            insight_card("External agency records", f"{filtered['external_referral_agency_clean'].astype('string').ne(MISSING).sum():,}", "Where agency is specified")

st.markdown("<div class='section-heading dashboard-section-heading'>Dashboard Section</div>", unsafe_allow_html=True)
st.markdown("<div class='section-subtitle'>Select the analytical view you want to explore.</div>", unsafe_allow_html=True)
section_options = [
    "Overview",
    "CPVs KPIs",
    "Demographics",
    "Games & Activities",
    "Protection & Support",
    "Referrals",
    "Data Quality",
]
section = dashboard_nav(section_options)

if section == "Overview":
    st.subheader("Overview")
    st.caption("A focused landing view for the current filter path, with each topic separated into Chart, Table, and Notes.")
    o1, o2, o3, o4 = st.columns(4)
    with o1:
        metric_card("CFS visits", f"{len(filtered):,}", "Current filter")
    with o2:
        metric_card("Issue mentions", f"{len(issue_context):,}", "Multi-response count")
    with o3:
        metric_card("Support mentions", f"{len(support_context):,}", "Multi-response count")
    with o4:
        metric_card("Referral rate", f"{referral_rate:.1%}", "Referral marked Yes")

    overview_camp_table = table_with_total(filtered, ["settlement_clean"], ["gender_clean"])
    overview_camp_chart = count_table(filtered, ["settlement_clean", "gender_clean"])
    render_analysis_block(
        "Coverage: Camp Records by Gender",
        overview_camp_table,
        lambda: sliced_bar_chart(overview_camp_chart, "settlement_clean", "overview_camps", "Count", "settlement_clean", "gender_clean", "Overview camp records by gender", horizontal=True, height=420),
        category_share_insight(overview_camp_chart, "settlement_clean", "camp"),
        "overview_camp_records",
        badges=[("Filtered", "good"), ("Top 5 default", "harmonized")],
    )

    overview_age_table = table_with_total(filtered, ["age_group"], ["gender_clean"])
    overview_age_chart = count_table(filtered, ["age_group", "gender_clean"], order_col="age_group")
    render_analysis_block(
        "Beneficiary Profile: Age Group by Gender",
        overview_age_table,
        lambda: bar_chart(overview_age_chart, "age_group", "Count", "gender_clean", "Overview age group by gender", category_orders={"age_group": AGE_GROUP_ORDER, "gender_clean": GENDER_ORDER}),
        category_share_insight(overview_age_chart, "age_group", "age group"),
        "overview_age_group",
        badges=[("Age order locked", "good"), ("Gender categories aligned", "harmonized")],
    )

    overview_issue_table = table_with_total(issue_context, ["issue_clean"], ["gender_clean"])
    overview_issue_chart = count_table(issue_context, ["issue_clean", "gender_clean"])
    render_analysis_block(
        "Protection: Top Reported Issues",
        overview_issue_table,
        lambda: sliced_bar_chart(overview_issue_chart, "issue_clean", "overview_issues", "Count", "issue_clean", "gender_clean", "Overview top reported issues", horizontal=True, height=520),
        category_share_insight(overview_issue_chart, "issue_clean", "reported issue"),
        "overview_issue_records",
        badges=[("Multi-response", "review"), ("Harmonized specify values", "harmonized")],
    )

    overview_game_table = table_with_total(game_context, ["game_clean"], ["gender_clean"])
    overview_game_chart = count_table(game_context, ["game_clean", "gender_clean"])
    render_analysis_block(
        "Activities & Referrals: Top Games / Activities",
        overview_game_table,
        lambda: sliced_bar_chart(overview_game_chart, "game_clean", "overview_games", "Count", "game_clean", "gender_clean", "Overview top games and activities", horizontal=True, height=520),
        category_share_insight(overview_game_chart, "game_clean", "game or activity"),
        "overview_game_records",
        badges=[("Take 5 harmonized", "harmonized"), ("Chart limited only", "good")],
    )

if section == "CPVs KPIs":
    staff_table = table_with_total(filtered, ["staff_clean"], ["gender_clean"])
    staff_table = add_interview_date_columns(staff_table, filtered, "staff_clean", "date")
    staff_chart = count_table(filtered, ["staff_clean", "gender_clean"])
    render_analysis_block(
        "CPV / Staff Data Submission",
        staff_table,
        lambda: sliced_bar_chart(staff_chart, "staff_clean", "staff", "Count", "staff_clean", "gender_clean", "Top staff filling forms by gender", horizontal=True, height=680, default=15),
        "Staff names are harmonized dynamically using exact mappings and fuzzy similarity so spelling variations are grouped together. First and latest interview dates are shown in the table after Total.",
        "staff_submission",
        badges=[("First/latest dates", "good"), ("Staff harmonized", "harmonized")],
    )

    st.divider()
    site_table = table_with_total(filtered, ["settlement_clean", "location_clean", "cfs_clean"], ["gender_clean"])
    site_chart = count_table(filtered, ["cfs_clean", "gender_clean"])
    render_analysis_block(
        "Camp, Location & CFS Distribution",
        site_table,
        lambda: sliced_bar_chart(site_chart, "cfs_clean", "site_distribution", "Count", "cfs_clean", "gender_clean", "CFS / site records by gender", horizontal=True, height=600, default=15),
        category_share_insight(site_chart, "cfs_clean", "CFS / site"),
        "site_distribution",
        badges=[("Linked location filters", "good"), ("Chart limited only", "harmonized")],
    )

if section == "Demographics":
    first_visit_records = filtered[filtered["first_visit_clean"].astype("string").isin(["Yes", "No"])].copy()
    first_visit_records["first_visit_clean"] = pd.Categorical(first_visit_records["first_visit_clean"].astype("string"), categories=["Yes", "No"], ordered=True)
    first_visit_pie = count_table(first_visit_records, ["first_visit_clean"])
    first_visit_overall_table = add_grand_total_row(first_visit_pie, "first_visit_clean")
    render_analysis_block(
        "First Visit to CFS: Overall Split",
        first_visit_overall_table,
        lambda: pie_chart(first_visit_pie, "first_visit_clean", "Count", "Overall first visit split"),
        category_share_insight(first_visit_pie, "first_visit_clean", "first-visit response"),
        "first_visit_overall",
        simple=True,
        badges=[("Yes/No view", "good")],
    )

    st.divider()
    first_visit_gender_table = table_with_total(first_visit_records, ["first_visit_clean"], ["gender_clean"])
    first_visit_gender_chart = count_table(first_visit_records, ["first_visit_clean", "gender_clean"])
    render_analysis_block(
        "First Visit to CFS by Gender",
        first_visit_gender_table,
        lambda: bar_chart(
            first_visit_gender_chart,
            "Count",
            "first_visit_clean",
            "gender_clean",
            "First visit to CFS by gender",
            horizontal=True,
            height=360,
            category_orders={"gender_clean": GENDER_ORDER},
        ),
        "This view keeps the Yes and No first-visit responses and splits each by the standardized gender categories.",
        "first_visit_gender",
        badges=[("Gender aligned", "harmonized")],
    )

    st.divider()
    first_visit_table = table_with_total(first_visit_records, ["cfs_clean"], ["first_visit_clean", "gender_clean"])
    first_visit_chart = count_table(first_visit_records, ["cfs_clean", "first_visit_clean"])
    render_analysis_block(
        "First Visit to CFS by Site",
        first_visit_table,
        lambda: sliced_bar_chart(
            first_visit_chart,
            "cfs_clean",
            "first_visit_site",
            "Count",
            "cfs_clean",
            "first_visit_clean",
            "First visit to CFS by site",
            horizontal=True,
            height=620,
            category_orders={"first_visit_clean": ["Yes", "No"]},
            default=15,
        ),
        "The site table keeps CFS / site as the row field, with Yes and No as the top-level columns. Missing first-visit responses are excluded from this specific Yes/No view.",
        "first_visit_site",
        badges=[("CFS site rows", "good"), ("Missing excluded", "review")],
    )

    st.divider()
    gender_cfs_table = table_with_total(filtered, ["cfs_clean"], ["gender_clean"])
    gender_cfs_chart = count_table(filtered, ["cfs_clean", "gender_clean"])
    render_analysis_block(
        "Gender by CFS",
        gender_cfs_table,
        lambda: sliced_bar_chart(gender_cfs_chart, "cfs_clean", "gender_cfs", "Count", "cfs_clean", "gender_clean", "Gender by CFS / site", horizontal=True, height=640, default=15),
        category_share_insight(gender_cfs_chart, "cfs_clean", "CFS / site"),
        "gender_by_cfs",
        badges=[("Girls / Boys / Transgender", "harmonized")],
    )

    st.divider()
    age_table = table_with_total(filtered, ["age_group"], ["gender_clean"])
    age_chart = count_table(filtered, ["age_group", "gender_clean"], order_col="age_group")
    render_analysis_block(
        "Age Group Breakdown by Gender",
        age_table,
        lambda: bar_chart(age_chart, "age_group", "Count", "gender_clean", "Overall age group distribution by gender", category_orders={"age_group": AGE_GROUP_ORDER, "gender_clean": GENDER_ORDER}),
        "Age groups are fixed in child-development order: 0-5 years, 6-12 years, then 13-17 years.",
        "age_group_overall",
        badges=[("Age order locked", "good")],
    )

    st.divider()
    age_cfs_table = table_with_total(filtered, ["cfs_clean"], ["age_group", "gender_clean"])
    age_cfs_chart = count_table(filtered, ["cfs_clean", "age_group"])
    render_analysis_block(
        "Age Group Breakdown by CFS & Gender",
        age_cfs_table,
        lambda: sliced_bar_chart(age_cfs_chart, "cfs_clean", "age_group_cfs_chart", "Count", "cfs_clean", "age_group", "Age group by CFS / site", horizontal=True, height=640, category_orders={"age_group": AGE_GROUP_ORDER}, default=15),
        "This detailed table shows the CFS-level age and gender breakdown. The chart summarizes age group distribution by CFS while the table carries the full gender detail.",
        "age_group_cfs_gender",
        badges=[("Detailed table", "good")],
    )

if section == "Protection & Support":
    disability_table = table_with_total(filtered, ["disability_status_clean"], ["gender_clean"])
    disability_chart = count_table(filtered, ["disability_status_clean", "gender_clean"])
    render_analysis_block(
        "Disability Prevalence",
        disability_table,
        lambda: bar_chart(disability_chart, "Count", "disability_status_clean", "gender_clean", "Disability prevalence by gender", horizontal=True, height=360),
        category_share_insight(disability_chart, "disability_status_clean", "disability status"),
        "disability_prevalence",
        badges=[("Yes/No view", "good")],
    )

    st.divider()
    disability_yes = filtered[filtered["disability_status_clean"].astype("string").eq("Yes")].copy()
    disability_type_table = table_with_total(disability_yes[~disability_yes["disability_type_display"].isin([MISSING])], ["disability_type_display"], ["gender_clean"])
    disability_chart_source = disability_yes[~disability_yes["disability_type_display"].isin([MISSING])].copy()
    disability_chart_source["disability_type_display"] = disability_chart_source["disability_type_display"].map(shorten_disability_type)
    disability_types = count_table(disability_chart_source, ["disability_type_display", "gender_clean"])
    render_analysis_block(
        "Disability Types",
        disability_type_table,
        lambda: sliced_bar_chart(disability_types, "disability_type_display", "disability_types", "Count", "disability_type_display", "gender_clean", "Disability types by gender", horizontal=True, height=480, default=10),
        "Long disability type descriptions are shortened only for display. The original workbook values remain available in the exported cleaned data.",
        "disability_types",
        badges=[("Readable chart labels", "harmonized")],
    )

    st.divider()
    issue_table = table_with_total(issue_context, ["issue_clean"], ["gender_clean"])
    issue_chart = count_table(issue_context, ["issue_clean", "gender_clean"])
    render_analysis_block(
        "Nature of Issues Reported",
        issue_table,
        lambda: sliced_bar_chart(issue_chart, "issue_clean", "issues", "Count", "issue_clean", "gender_clean", "Issues reported by gender", horizontal=True, height=720, default=15),
        category_share_insight(issue_chart, "issue_clean", "reported issue") + " Other/specify issue entries are harmonized into categories such as Take 5, Play / child engagement, Medical concern and GBV.",
        "nature_issues",
        badges=[("Multi-response", "review"), ("Specify harmonized", "harmonized")],
    )
    issue_distribution, issue_combinations = multi_response_tables(issue_context, "issue_clean", "Issue")
    st.markdown("#### Multiple issues per child / record")
    issue_dist_col, issue_combo_col = st.columns([1, 1.8])
    with issue_dist_col:
        render_table(issue_distribution, "issue_selection_distribution", "Issue Selection Count per Child / Record", simple=True)
    with issue_combo_col:
        render_table(issue_combinations, "issue_combinations", "Most Common Issue Combinations", simple=True)
    st.markdown("#### Issues Captured by CPV / Staff")
    issue_staff_table = table_with_total(issue_context, ["staff_clean"], ["issue_clean"])
    render_table(issue_staff_table, "issues_by_cpv_staff", "Issues Captured by CPV / Staff")

    st.divider()
    support_table = table_with_total(support_context, ["support_clean"], ["gender_clean"])
    support_chart = count_table(support_context, ["support_clean", "gender_clean"])
    render_analysis_block(
        "Support Offered",
        support_table,
        lambda: bar_chart(support_chart, "Count", "support_clean", "gender_clean", "Support offered by gender", horizontal=True, height=420),
        category_share_insight(support_chart, "support_clean", "support type"),
        "support_offered",
        badges=[("Multi-response", "review")],
    )
    support_distribution, support_combinations = multi_response_tables(support_context, "support_clean", "Support")
    st.markdown("#### Multiple support types per child / record")
    support_dist_col, support_combo_col = st.columns([1, 1.8])
    with support_dist_col:
        render_table(support_distribution, "support_selection_distribution", "Support Selection Count per Child / Record", simple=True)
    with support_combo_col:
        render_table(support_combinations, "support_combinations", "Most Common Support Combinations", simple=True)

if section == "Games & Activities":
    games_table = table_with_total(game_context, ["game_clean"], ["gender_clean"])
    games_chart = count_table(game_context, ["game_clean", "gender_clean"])
    render_analysis_block(
        "Games / Activities Engagement",
        games_table,
        lambda: sliced_bar_chart(games_chart, "game_clean", "games", "Count", "game_clean", "gender_clean", "Games / activities by gender", horizontal=True, height=680, default=15),
        category_share_insight(games_chart, "game_clean", "game or activity") + " Take5 variants and Take 5 routine values are grouped under Take 5.",
        "games_activities",
        badges=[("Take 5 harmonized", "harmonized"), ("Chart limited only", "good")],
    )
    game_distribution, game_combinations = multi_response_tables(game_context, "game_clean", "Game / activity")
    st.markdown("#### Multiple games / activities per child / record")
    game_dist_col, game_combo_col = st.columns([1, 1.8])
    with game_dist_col:
        render_table(game_distribution, "game_selection_distribution", "Game / Activity Selection Count per Child / Record", simple=True)
    with game_combo_col:
        render_table(game_combinations, "game_combinations", "Most Common Game / Activity Combinations", simple=True)

if section == "Referrals":
    referrals = filtered[filtered["referral_made_clean"].astype("string").eq("Yes")].copy()
    referral_dest = referrals[referrals["referral_destination_grouped"].isin(["PSS", "Case Management", "Empowerment", "External referrals"])]
    referral_table = table_with_total(referral_dest, ["referral_destination_grouped"], ["gender_clean"])
    referral_chart = count_table(referral_dest, ["referral_destination_grouped", "gender_clean"])
    render_analysis_block(
        "Referral Destination If Yes",
        referral_table,
        lambda: bar_chart(
            referral_chart,
            "Count",
            "referral_destination_grouped",
            "gender_clean",
            "Referral destination by gender",
            horizontal=True,
            height=420,
        ),
        category_share_insight(referral_chart, "referral_destination_grouped", "referral destination") + " External referral partners are intentionally excluded here and shown in the agency breakdown.",
        "referral_destination",
        badges=[("Four destination groups", "good")],
    )

    st.divider()
    external = referrals[
        referrals["referral_destination_grouped"].eq("External referrals")
        & ~referrals["external_referral_agency_clean"].isin([MISSING, "Unknown", REVIEW])
    ].copy()
    external_table = table_with_total(external, ["external_referral_agency_clean"], ["gender_clean"])
    external_chart = count_table(external, ["external_referral_agency_clean", "gender_clean"])
    render_analysis_block(
        "External Referral Agency Breakdown",
        external_table,
        lambda: sliced_bar_chart(external_chart, "external_referral_agency_clean", "external_agency", "Count", "external_referral_agency_clean", "gender_clean", "External referral agencies by gender", horizontal=True, height=600, default=15),
        category_share_insight(external_chart, "external_referral_agency_clean", "external agency"),
        "external_referral_agencies",
        badges=[("Agency names harmonized", "harmonized"), ("Unknown removed", "review")],
    )

if section == "Data Quality":
    st.subheader("Data Quality & Harmonization Review")
    q1, q2, q3, q4 = st.columns(4)
    with q1:
        metric_card("Missing location", f"{filtered['location_clean'].eq(MISSING).sum():,}")
    with q2:
        metric_card("Missing CFS / site", f"{filtered['cfs_clean'].eq(MISSING).sum():,}")
    with q3:
        metric_card("Missing age", f"{filtered['age_clean'].isna().sum():,}")
    with q4:
        metric_card("Issue Other harmonized", f"{filtered['issue_other_specify'].notna().sum():,}")

    staff_review = (
        filtered.groupby(["staff_filling_form", "staff_clean"], dropna=False, observed=False)
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )
    staff_review = top_n_with_total(staff_review, 80, "staff_filling_form")
    staff_review_chart = staff_review[staff_review["staff_filling_form"].astype("string").ne("Grand Total")].head(20).copy()
    render_analysis_block(
        "Harmonized Staff Name Review",
        staff_review,
        lambda: bar_chart(staff_review_chart, "Count", "staff_clean", title="Top harmonized staff names", horizontal=True, height=560),
        "This review shows the raw staff entry beside the cleaned staff name, making spelling corrections auditable.",
        "staff_harmonization_review",
        simple=True,
        badges=[("Harmonized", "harmonized")],
    )

    st.divider()
    issue_review = (
        filtered[filtered["issue_other_specify"].notna()]
        .groupby(["issue_other_specify", "issue_other_specify_clean"], dropna=False, observed=False)
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )
    issue_review = top_n_with_total(issue_review, 80, "issue_other_specify")
    issue_review_chart = issue_review[issue_review["issue_other_specify"].astype("string").ne("Grand Total")].head(20).copy()
    render_analysis_block(
        "Other Issue Specification Review",
        issue_review,
        lambda: bar_chart(issue_review_chart, "Count", "issue_other_specify_clean", title="Top harmonized issue specifications", horizontal=True, height=560),
        "This review makes the Other/specify harmonization transparent, including Take 5 and protection-related recoding.",
        "issue_other_review",
        simple=True,
        badges=[("Specify harmonized", "harmonized")],
    )

    st.divider()
    disability_review = (
        filtered[filtered["disability_type"].notna()]
        .groupby(["disability_type", "disability_type_display"], dropna=False, observed=False)
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )
    disability_review = add_grand_total_row(disability_review, "disability_type")
    disability_review_chart = disability_review[disability_review["disability_type"].astype("string").ne("Grand Total")].head(20).copy()
    render_analysis_block(
        "Disability Display Label Review",
        disability_review,
        lambda: bar_chart(disability_review_chart, "Count", "disability_type_display", title="Top disability display labels", horizontal=True, height=520),
        "This review confirms long disability descriptions are shortened for readability without changing the underlying workbook values.",
        "disability_label_review",
        simple=True,
        badges=[("Readable labels", "harmonized")],
    )
