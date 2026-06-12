from pathlib import Path
from difflib import SequenceMatcher
import re
import time

import pandas as pd
import plotly.express as px
import streamlit as st


DATA_PATH = Path(r"./data/CFS_QUESTIONNAIRE_Tdh_Kenya_T1.xlsx")
SHEET_NAME = "Transformed Data"
PREPARED_CACHE_PATH = Path(__file__).with_name(".cfs_dashboard_prepared_cache.pkl")
CACHE_VERSION = "cfs-dashboard-prepared-v5"

MISSING = "Missing / unspecified"
REVIEW = "Needs review"
AGE_GROUP_ORDER = ["0-5 years", "6-12 years", "13-17 years", MISSING]
YES_NO_ORDER = ["Yes", "No", MISSING]
GENDER_ORDER = ["Boy", "Girl", MISSING]
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
    upper_tokens = {"cfs", "tdh", "unhcr", "drs", "lwf", "hi", "wfp", "nrc", "drc", "rck", "krcs", "pss", "gbv"}
    return " ".join(token.upper() if token.lower() in upper_tokens else token.capitalize() for token in text.split())


def yes_no(value):
    key = norm_text(value)
    if key in {"yes", "y", "1", "true"}:
        return "Yes"
    if key in {"no", "n", "0", "false"}:
        return "No"
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


def clean_game(row):
    raw = row.get("games_played", None)
    other = row.get("game_other_specify", None)
    raw_key = norm_text(raw)
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
    return fuzzy_harmonize(value, ISSUE_OTHER_MAP, cutoff=0.84)


def ordered_unique(series, order=None):
    values = [v for v in series.dropna().unique().tolist() if str(v).strip() != ""]
    values = [v for v in values if v != MISSING] + ([MISSING] if MISSING in values else [])
    if order:
        order_map = {v: i for i, v in enumerate(order)}
        return sorted(values, key=lambda v: (order_map.get(v, 999), str(v)))
    return sorted(values, key=lambda v: str(v))


@st.cache_data(show_spinner=False, ttl=300)
def load_and_prepare(path_text, modified_time):
    path = Path(path_text)
    try:
        if PREPARED_CACHE_PATH.exists():
            payload = pd.read_pickle(PREPARED_CACHE_PATH)
            if (
                payload.get("cache_version") == CACHE_VERSION
                and payload.get("source_path") == str(path)
                and payload.get("modified_time") == modified_time
            ):
                return payload["df"], payload["issue_long"], payload["support_long"]
    except Exception:
        pass

    df = pd.read_excel(path, sheet_name=SHEET_NAME)
    df = df.copy()
    df.insert(0, "record_id", range(1, len(df) + 1))

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["month"] = df["date"].dt.to_period("M").astype("string")
    df["consent_clean"] = df["consent"].map(yes_no)
    df["staff_clean"] = df["staff_filling_form"].map(clean_staff)
    df["gender_clean"] = df["child_gender"].map(lambda v: harmonize_from_map(v, {}, fallback_title=True))
    df["gender_clean"] = df["gender_clean"].replace({"Male": "Boy", "Female": "Girl"})
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

    support_frames = []
    for col, label in SUPPORT_COLUMNS.items():
        if col in df.columns:
            mask = df[col].map(is_yes)
            tmp = df.loc[mask, ["record_id"]].copy()
            tmp["support_clean"] = label
            support_frames.append(tmp)
    support_long = pd.concat(support_frames, ignore_index=True) if support_frames else pd.DataFrame(columns=["record_id", "support_clean"])

    issues_combined = issue_long.groupby("record_id")["issue_clean"].apply(lambda x: "; ".join(sorted(set(x)))).rename("issues_combined")
    support_combined = support_long.groupby("record_id")["support_clean"].apply(lambda x: "; ".join(sorted(set(x)))).rename("support_combined")
    df = df.merge(issues_combined, on="record_id", how="left").merge(support_combined, on="record_id", how="left")
    df["issues_combined"] = df["issues_combined"].fillna(MISSING)
    df["support_combined"] = df["support_combined"].fillna(MISSING)

    try:
        pd.to_pickle(
            {
                "cache_version": CACHE_VERSION,
                "source_path": str(path),
                "modified_time": modified_time,
                "df": df,
                "issue_long": issue_long,
                "support_long": support_long,
            },
            PREPARED_CACHE_PATH,
        )
    except Exception:
        pass

    return df, issue_long, support_long


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
    default_index = TOP_N_OPTIONS.index(default) if default in TOP_N_OPTIONS else 1
    mode_col, count_col = st.columns([1, 1.4])
    with mode_col:
        mode = st.radio(
            label,
            ["Highest", "Lowest"],
            horizontal=True,
            key=f"chart_rank_mode_{key}",
            help="This limits the chart only. The table above remains complete.",
        )
    with count_col:
        selected_n = st.radio(
            "Number of categories",
            TOP_N_OPTIONS,
            index=default_index,
            horizontal=True,
            key=f"chart_top_n_{key}",
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
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False, "responsive": True})


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
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False, "responsive": True})


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

st.markdown(
    """
    <div class="dashboard-title">
        <h1>Tdh Kenya Child Friendly Spaces Dashboard</h1>
    </div>
    """,
    unsafe_allow_html=True,
)
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
df, issue_long, support_long = load_and_prepare(str(DATA_PATH), modified_time)

st.sidebar.header("Filters")
st.sidebar.caption("Use each level in order. Later filters unlock after the location path is defined.")
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

downstream_disabled = cfs_disabled or not cfs_explicit or not selected_cfs or cfs_scope.empty
selected_staff, staff_explicit = multi_choice_selector(
    "5. Staff / CPV",
    cfs_scope["staff_clean"] if not cfs_scope.empty else [],
    "staff",
    disabled=downstream_disabled,
    disabled_message="Choose at least one CFS / site first.",
)
staff_scope = apply_filter(cfs_scope, "staff_clean", selected_staff, staff_explicit)

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

context_columns = ["record_id", "settlement_clean", "location_clean", "cfs_clean", "gender_clean", "age_group"]
issue_context = issue_long.merge(filtered[context_columns], on="record_id", how="inner")
support_context = support_long.merge(filtered[context_columns], on="record_id", how="inner")

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
    st.markdown("<div class='quick-insights'>", unsafe_allow_html=True)
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
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div class='section-heading dashboard-section-heading'>Dashboard Section</div>", unsafe_allow_html=True)
st.markdown("<div class='section-subtitle'>Select the analytical view you want to explore.</div>", unsafe_allow_html=True)
section = st.radio(
    "Dashboard section",
    [
        "CPVs Data Distribution",
        "CFS Beneficiary Demographics",
        "Protection & Support",
        "Games, Activities & Referrals",
        "Data Quality & Harmonization",
    ],
    horizontal=True,
    label_visibility="collapsed",
)

if section == "CPVs Data Distribution":
    st.subheader("CPV / Staff Data Submission")
    staff_table = table_with_total(filtered, ["staff_clean"], ["gender_clean"])
    staff_table = add_interview_date_columns(staff_table, filtered, "staff_clean", "date")
    st.dataframe(style_total(staff_table), use_container_width=True)
    staff_chart = count_table(filtered, ["staff_clean", "gender_clean"])
    staff_chart = top_n_chart_control(staff_chart, "staff_clean", "staff", default=15)
    bar_chart(staff_chart, "Count", "staff_clean", "gender_clean", "Top staff filling forms by gender", horizontal=True, height=680)
    section_note("Staff names are harmonized dynamically using exact mappings and fuzzy similarity so spelling variations are grouped together.")

    st.divider()
    st.subheader("Settlement, Location & CFS Distribution")
    site_table = table_with_total(filtered, ["settlement_clean", "location_clean", "cfs_clean"], ["gender_clean"])
    st.dataframe(style_total(site_table), use_container_width=True)
    site_chart = count_table(filtered, ["cfs_clean", "gender_clean"])
    site_chart = top_n_chart_control(site_chart, "cfs_clean", "site_distribution", default=15)
    bar_chart(site_chart, "Count", "cfs_clean", "gender_clean", "CFS / site records by gender", horizontal=True, height=600)

if section == "CFS Beneficiary Demographics":
    st.subheader("First Visit to CFS")
    first_visit_records = filtered[filtered["first_visit_clean"].astype("string").isin(["Yes", "No"])].copy()
    first_visit_records["first_visit_clean"] = pd.Categorical(first_visit_records["first_visit_clean"].astype("string"), categories=["Yes", "No"], ordered=True)
    first_visit_pie = count_table(first_visit_records, ["first_visit_clean"])
    pie_chart(first_visit_pie, "first_visit_clean", "Count", "Overall first visit split")

    st.markdown("#### First visit to CFS by Gender")
    first_visit_gender_table = table_with_total(first_visit_records, ["first_visit_clean"], ["gender_clean"])
    st.dataframe(style_total(first_visit_gender_table), use_container_width=True)
    first_visit_gender_chart = count_table(first_visit_records, ["first_visit_clean", "gender_clean"])
    bar_chart(
        first_visit_gender_chart,
        "Count",
        "first_visit_clean",
        "gender_clean",
        "First visit to CFS by gender",
        horizontal=True,
        height=360,
        category_orders={"gender_clean": GENDER_ORDER},
    )

    st.markdown("#### First visit to CFS by Site")
    first_visit_table = table_with_total(first_visit_records, ["cfs_clean"], ["first_visit_clean", "gender_clean"])
    st.dataframe(style_total(first_visit_table), use_container_width=True)
    first_visit_chart = count_table(first_visit_records, ["cfs_clean", "first_visit_clean"])
    first_visit_chart = top_n_chart_control(first_visit_chart, "cfs_clean", "first_visit_site", default=15)
    bar_chart(
        first_visit_chart,
        "Count",
        "cfs_clean",
        "first_visit_clean",
        "First visit to CFS by site",
        horizontal=True,
        height=620,
        category_orders={"first_visit_clean": ["Yes", "No"]},
    )
    section_note("The site table keeps CFS / site as the row field, with Yes and No as the top-level columns. Missing first-visit responses are excluded from this specific Yes/No view.")

    st.divider()
    st.subheader("Gender by CFS")
    gender_cfs_table = table_with_total(filtered, ["cfs_clean"], ["gender_clean"])
    st.dataframe(style_total(gender_cfs_table), use_container_width=True)
    gender_cfs_chart = count_table(filtered, ["cfs_clean", "gender_clean"])
    gender_cfs_chart = top_n_chart_control(gender_cfs_chart, "cfs_clean", "gender_cfs", default=15)
    bar_chart(gender_cfs_chart, "Count", "cfs_clean", "gender_clean", "Gender by CFS / site", horizontal=True, height=640)

    st.divider()
    st.subheader("Age Group Breakdown by CFS & Gender")
    age_table = table_with_total(filtered, ["age_group"], ["gender_clean"])
    st.dataframe(style_total(age_table), use_container_width=True)
    age_cfs_table = table_with_total(filtered, ["cfs_clean"], ["age_group", "gender_clean"])
    st.dataframe(style_total(age_cfs_table), use_container_width=True)
    age_chart = count_table(filtered, ["age_group", "gender_clean"], order_col="age_group")
    bar_chart(age_chart, "age_group", "Count", "gender_clean", "Overall age group distribution by gender", category_orders={"age_group": AGE_GROUP_ORDER, "gender_clean": GENDER_ORDER})

if section == "Protection & Support":
    st.subheader("Disability Prevalence & Disability Types")
    disability_table = table_with_total(filtered, ["disability_status_clean"], ["gender_clean"])
    st.dataframe(style_total(disability_table), use_container_width=True)
    disability_chart = count_table(filtered, ["disability_status_clean", "gender_clean"])
    bar_chart(disability_chart, "Count", "disability_status_clean", "gender_clean", "Disability prevalence by gender", horizontal=True, height=360)

    st.markdown("#### Disability types")
    disability_yes = filtered[filtered["disability_status_clean"].astype("string").eq("Yes")].copy()
    disability_type_table = table_with_total(disability_yes[~disability_yes["disability_type_display"].isin([MISSING])], ["disability_type_display"], ["gender_clean"])
    st.dataframe(style_total(disability_type_table), use_container_width=True)
    disability_chart_source = disability_yes[~disability_yes["disability_type_display"].isin([MISSING])].copy()
    disability_chart_source["disability_type_display"] = disability_chart_source["disability_type_display"].map(shorten_disability_type)
    disability_types = count_table(disability_chart_source, ["disability_type_display", "gender_clean"])
    disability_types = top_n_chart_control(disability_types, "disability_type_display", "disability_types", default=10)
    bar_chart(disability_types, "Count", "disability_type_display", "gender_clean", "Disability types by gender", horizontal=True, height=480)
    section_note("Long disability type descriptions are shortened only for display. The original workbook values remain available in the exported cleaned data.")

    st.divider()
    st.subheader("Nature of Issues Reported")
    issue_table = table_with_total(issue_context, ["issue_clean"], ["gender_clean"])
    st.dataframe(style_total(issue_table), use_container_width=True)
    issue_chart = count_table(issue_context, ["issue_clean", "gender_clean"])
    issue_chart = top_n_chart_control(issue_chart, "issue_clean", "issues", default=15)
    bar_chart(issue_chart, "Count", "issue_clean", "gender_clean", "Issues reported by gender", horizontal=True, height=720)
    section_note("The Other/specify issue entries are harmonized into readable categories such as Take 5, Play / child engagement, Medical concern and GBV.")

    st.divider()
    st.subheader("Support Offered")
    support_table = table_with_total(support_context, ["support_clean"], ["gender_clean"])
    st.dataframe(style_total(support_table), use_container_width=True)
    support_chart = count_table(support_context, ["support_clean", "gender_clean"])
    bar_chart(support_chart, "Count", "support_clean", "gender_clean", "Support offered by gender", horizontal=True, height=420)

if section == "Games, Activities & Referrals":
    st.subheader("Games / Activities Engagement")
    games_table = table_with_total(filtered, ["games_played_clean"], ["gender_clean"])
    st.dataframe(style_total(games_table), use_container_width=True)
    games_chart = count_table(filtered, ["games_played_clean", "gender_clean"])
    games_chart = top_n_chart_control(games_chart, "games_played_clean", "games", default=15)
    bar_chart(games_chart, "Count", "games_played_clean", "gender_clean", "Games / activities by gender", horizontal=True, height=680)

    st.divider()
    st.subheader("Referral Destination If Yes")
    referrals = filtered[filtered["referral_made_clean"].astype("string").eq("Yes")].copy()
    referral_dest = referrals[referrals["referral_destination_grouped"].isin(["PSS", "Case Management", "Empowerment", "External referrals"])]
    referral_table = table_with_total(referral_dest, ["referral_destination_grouped"], ["gender_clean"])
    st.dataframe(style_total(referral_table), use_container_width=True)
    referral_chart = count_table(referral_dest, ["referral_destination_grouped", "gender_clean"])
    bar_chart(
        referral_chart,
        "Count",
        "referral_destination_grouped",
        "gender_clean",
        "Referral destination by gender",
        horizontal=True,
        height=420,
    )
    section_note("External referral partners are intentionally excluded from this table and shown separately below.")

    st.divider()
    st.subheader("External Referral Agency Breakdown")
    external = referrals[
        referrals["referral_destination_grouped"].eq("External referrals")
        & ~referrals["external_referral_agency_clean"].isin([MISSING, "Unknown", REVIEW])
    ].copy()
    external_table = table_with_total(external, ["external_referral_agency_clean"], ["gender_clean"])
    st.dataframe(style_total(external_table), use_container_width=True)
    external_chart = count_table(external, ["external_referral_agency_clean", "gender_clean"])
    external_chart = top_n_chart_control(external_chart, "external_referral_agency_clean", "external_agency", default=15)
    bar_chart(external_chart, "Count", "external_referral_agency_clean", "gender_clean", "External referral agencies by gender", horizontal=True, height=600)

if section == "Data Quality & Harmonization":
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

    st.markdown("#### Harmonized staff name review")
    staff_review = (
        filtered.groupby(["staff_filling_form", "staff_clean"], dropna=False, observed=False)
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )
    staff_review = top_n_with_total(staff_review, 80, "staff_filling_form")
    st.dataframe(style_simple_total(staff_review), use_container_width=True)

    st.markdown("#### Other issue specification review")
    issue_review = (
        filtered[filtered["issue_other_specify"].notna()]
        .groupby(["issue_other_specify", "issue_other_specify_clean"], dropna=False, observed=False)
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )
    issue_review = top_n_with_total(issue_review, 80, "issue_other_specify")
    st.dataframe(style_simple_total(issue_review), use_container_width=True)

    st.markdown("#### Disability display label review")
    disability_review = (
        filtered[filtered["disability_type"].notna()]
        .groupby(["disability_type", "disability_type_display"], dropna=False, observed=False)
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )
    disability_review = add_grand_total_row(disability_review, "disability_type")
    st.dataframe(style_simple_total(disability_review), use_container_width=True)
