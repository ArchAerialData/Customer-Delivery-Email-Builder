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
