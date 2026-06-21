# Tdh Kenya CFS Dashboard

Professional Streamlit dashboard for analysing TDH Kenya Child Friendly Spaces data.

## Files

```text
app.py                  Main Streamlit dashboard
assets/styles.css       Professional CSS theme for app, buttons, menus, cards, charts, and tables
ui_formatting_patch.py  Optional reusable UI helper functions
requirements.txt        Python dependencies
README.md               Setup and usage notes
data/.gitkeep           Placeholder for the data folder
```

## How to run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Data options

Use one of the following server-side options:

1. Put your Excel file here:

```text
data/CFS_QUESTIONNAIRE_Tdh_Kenya_T1.xlsx
```

2. Set an environment variable:

```bash
export CFS_DATA_PATH="/path/to/CFS_QUESTIONNAIRE_Tdh_Kenya_T1.xlsx"
streamlit run app.py
```


## Expected data columns

The app is defensive and will still open if some columns are missing. It works best with columns such as:

- `date`
- `consent`
- `staff_filling_form`
- `child_gender`
- `child_age`
- `child_living_with_disability`
- `disability_type`
- `first_visit_tdh_cfs`
- `referral_made`
- `referral_destination`
- `external_referral_agency`
- `camp_of_information_seeking`
- `specific_camp_location`
- `camp_location_alt`
- `exact_registered_location`
- `child_friendly_space_visited`
- `cfs_visited`
- `games_played`
- `game_other_specify`
- issue columns such as `issue_basic_needs`, `issue_education`, etc.
- support columns such as `support_psychological_first_aid`, etc.

## Notes

- Records where `consent` is clearly `No` are excluded from the dashboard.
- Missing columns are created automatically with blank/default values so the app does not crash.
- The styling is contained in `assets/styles.css` and is loaded automatically by `app.py`.
