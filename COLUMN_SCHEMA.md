# CFS Dashboard Analysis Column Schema

The dashboard now standardises incoming data to the transformed `analysis_column` schema before analysis. It supports both raw survey question labels and transformed column names.

Key analysis columns include:

- consent
- staff_filling_form
- date
- child_name
- child_age
- child_gender
- child_living_with_disability
- disability_type
- disability_type_other
- camp_of_information_seeking
- specific_camp_location
- section_block_residence
- camp_location_alt
- exact_registered_location
- child_friendly_space_visited
- cfs_visited
- games_played
- game_other_specify
- take5_activities_integrated
- first_visit_tdh_cfs
- nature_issues_reported_text
- issue_new_arrival_lack_of_card
- issue_disability
- issue_basic_needs
- issue_deceased_parent
- issue_education
- issue_psychosocial_support
- issue_neglected
- issue_parents_separated
- issue_child_out_of_wedlock
- issue_food
- issue_clothing
- issue_shelter
- issue_reporting_protection_concern
- issue_need_profiling_registration_unhcr
- issue_none
- issue_other
- issue_other_specify
- support_offered_text
- support_psychological_first_aid
- support_play_art_therapy
- support_psychoeducation
- support_none
- referral_made
- referral_destination
- external_referral_agency

Operational metadata columns are also supported where present: today, username, deviceid, phonenumber, id, uuid, submission_time, validation_status, notes, status, submitted_by, version, tags, root_uuid, index.
