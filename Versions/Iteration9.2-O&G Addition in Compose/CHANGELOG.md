# Iteration1 -> Iteration2

Added files: 2026_02_03/AALLC Survey Details and Distribution Lists.xlsx
Removed files: None
Changed files: 2026_02_03/2025FullSDDL.xlsx, 2026_02_03/Client-Site-URL-Email.xlsx

# Iteration2 -> Iteration3

Added files: Email Templates/First Site (smclaren@archaerial.com).eml, Email Templates/First-Photo&Video (smclaren@archaerial.com).eml, Email Templates/GFT (smclaren@archaerial.com).eml, Email Templates/Leetsdale (smclaren@archaerial.com).eml, Email Templates/Main (smclaren@archaerial.com).eml, Email Templates/Main Bottom (smclaren@archaerial.com).eml, Email Templates/PhotoVideo2DMap (smclaren@archaerial.com).eml, Email Templates/Reply (smclaren@archaerial.com).eml
Removed files: None
Changed files: 2026_02_03/2025FullSDDL.xlsx, 2026_02_03/Client-Site-URL-Email.xlsx, __pycache__/email_template_gui.cpython-313.pyc, email_template_gui.py
Key logic changes (email_template_gui.py):

- email_templates_dir: False -> True
- force_sync: False -> True
- uses_contacts_file: True -> False
- uses_dropbox_file: True -> False
- uses_master_file: False -> True

# Iteration3 -> Iteration4

Added files: None
Removed files: None
Changed files: email_template_gui.py, Versions/CHANGELOG.md
Key logic changes (email_template_gui.py):

- reload_buttons_refresh_ui: False -> True
- reload_resets_compose_defaults: False -> True

# Iteration4 -> Iteration5

Added files: None
Removed files: None
Changed files: email_template_gui.py, Versions/CHANGELOG.md
Key logic changes (email_template_gui.py):

- outlook_recipients_resolve_all: False -> True
- outlook_recipient_add_by_list: False -> True

# Iteration5 -> Iteration6

Added files: None
Removed files: None
Changed files: email_template_gui.py, Versions/CHANGELOG.md
Key logic changes (email_template_gui.py):

- master_tab_scrollbar_grid_layout: False -> True

# Iteration8-EmailImages -> Iteration9-Doug Addition

Added files: Email Templates/Doug/First Site.eml, Email Templates/Doug/First-Photo&Video.eml, Email Templates/Doug/GFT.eml, Email Templates/Doug/GFTEncroachemtn&Video.eml, Email Templates/Doug/Main.eml, Email Templates/Steven/First Site.eml, Email Templates/Steven/First-Photo&Video.eml, Email Templates/Steven/GFT.eml, Email Templates/Steven/GFTEncroachemtn&Video.eml, Email Templates/Steven/Main.eml
Removed files: Email Templates/First Site (smclaren@archaerial.com).eml, Email Templates/First-Photo&Video (smclaren@archaerial.com).eml, Email Templates/GFT (smclaren@archaerial.com).eml, Email Templates/Leetsdale (smclaren@archaerial.com).eml, Email Templates/Main (smclaren@archaerial.com).eml, Email Templates/Main Bottom (smclaren@archaerial.com).eml, Email Templates/New/First Site-Photo&Video.eml, Email Templates/New/First Site-Photo.eml, Email Templates/New/Main.eml, Email Templates/PhotoVideo2DMap (smclaren@archaerial.com).eml, Email Templates/Reply (smclaren@archaerial.com).eml
Changed files: 2026_02_03/Client-Site-URL-Email.xlsx, __pycache__/email_template_gui.cpython-313.pyc, email_template_gui.py, Versions/CHANGELOG.md
Key logic changes (email_template_gui.py):

- compose_template_source_folders_detected_from_Email_Templates: False -> True
- compose_template_folder_selector_next_to_bcc: False -> True
- compose_only_one_template_folder_selectable_at_a_time: False -> True
- compose_template_dropdown_populates_from_selected_folder: False -> True
- compose_default_prefers_Main_template_name: False -> True

# Iteration9-Doug Addition -> Iteration9.1-Copyable

Added files: None
Removed files: None
Changed files: __pycache__/email_template_gui.cpython-313.pyc, email_template_gui.py, Versions/CHANGELOG.md
Key logic changes (email_template_gui.py):

- app_depends_on_Outlook_signature_files_in_user_profile: True -> False
- compose_template_html_rehydrates_cid_images_from_AppData_signatures: True -> False
- compose_falls_back_to_AppData_signature_html_when_template_missing_html: True -> False
- body_dropdown_includes_user_profile_signature_names: True -> False

# Iteration9.1-Copyable -> Iteration9.2-O&G Addition

Added files: .gitignore, Email Templates/Doug/O&G.eml, Email Templates/Steven/O&G.eml
Removed files: None
Changed files: 2026_02_03/Client-Site-URL-Email.xlsx, 2026_02_03/FULL_SDDL_Cleaned.xlsx, Email Templates/Doug/First Site.eml, Email Templates/Doug/First-Photo&Video.eml, Email Templates/Doug/GFT.eml, Email Templates/Doug/GFTEncroachemtn&Video.eml, Email Templates/Doug/Main.eml, Email Templates/Steven/First Site.eml, Email Templates/Steven/First-Photo&Video.eml, Email Templates/Steven/Main.eml, __pycache__/email_template_gui.cpython-313.pyc, email_template_gui.py, Versions/CHANGELOG.md
Key logic changes (email_template_gui.py):

- client_sites_tab_renamed_from_Client_Sites_to_Construction: False -> True
- oil_and_gas_client_sites_tab_added_with_matching_CRUD_workflow: False -> True
- client_sites_workbook_sheet_resolution_uses_explicit_sheet_names_instead_of_active_sheet_fallback: False -> True
- construction_sheet_accepts_legacy_ABCD_alias: False -> True
- oil_and_gas_sheet_can_be_created_from_the_UI_if_missing: False -> True
- compose_client_and_site_dropdowns_follow_selected_CC_list_sheet_context: False -> True
- compose_defaults_to_Commercial_Construction_but_can_switch_to_Oil_&_Gas_data: False -> True
- media_type_text_embeds_progress_for_standard_modes_and_aerial_for_ROW_documentation: False -> True
- Doug_signature_spacing_matches_Steven_spacing_across_active_templates: False -> True
