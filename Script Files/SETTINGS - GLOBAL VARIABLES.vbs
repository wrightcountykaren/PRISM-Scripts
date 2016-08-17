'COUNTY CUSTOM VARIABLES----------------------------------------------------------------------------------------------------
'The following variables are dynamically added via the installer. They can be modified manually to make changes without re-running the installer, but doing so should not be undertaken lightly.

'Run locally: if this is set to "True", the scripts will run locally and bypass GitHub entirely. This is great for debugging or developing scripts. Only scriptwriters should do it. An agency should always be set to "false".
run_locally = true

'Default directory: used by the script to determine if we're scriptwriters or not (scriptwriters use a default directory traditionally).
'	This is modified by the installer, which will determine if this is a scriptwriter or a production user.
default_directory = "C:\PRISM-Scripts\Script Files\"

'This is used for determining whether script_end_procedure will also log usage info in an Access table.
collecting_statistics = False

'This is the file path for the statistics Access database.
stats_database_path = "C:\PRISM-Scripts\Databases for script usage\usage statistics.accdb"

'This is the folder path for county-specific Word documents. Modify this with your shared-drive location for Word documents.
word_documents_folder_path = "C:\PRISM-Scripts\Word files for script usage\"

'This is used by scripts which tell the worker where to find a doc to send to a client (ie "Send form using Compass Pilot")
EDMS_choice = "Compass Pilot"

'This is used for MEMO scripts, such as appointment letter
'Replace "Anoka" with your county name below. "Anoka County" just demonstrates the format for County Name.
county_name = "Anoka County"

'This is the county code on the CALI screen.
county_cali_code = "###"

'Creates a double array of county offices, first by office (using the ~), then by address line (using the |). Dynamically added with the installer.
'Address below is an example.  Replace with your county office address.
county_office_array = split("2100 3rd Ave Suite 400|Anoka, MN 55303", "~")

'This is a variable which signifies the agency uses the master branch or the RELEASE branch. Set to true if you're a scriptwriter agency and all users are going to be on the master branch. Otherwise, set to false.
use_master_branch = true

'This allows a "beta user" group to have access to master branch scripts, while everyone else uses release. This is helpful for counties that want to maintain a small test group.
'Here is the list of agency super users. These users will have access to the test scripts. Enter the list of users' log-in IDs in the quotes below, comma separated
beta_users = ""

'An array of county attorneys. "Select one:" should ALWAYS be in there, and ALWAYS be first. Replace "County Attorney #" with your agency's county attorney names.
county_attorney_array = array("Select one:", "County Attorney 1", "County Attorney 2", "County Attorney 3", "County Attorney 4", "County Attorney 5")

'An array of child support magistrates. "Select one:" should ALWAYS be in there, and ALWAYS be first.  Replace "Magistrate # with your agency's child support magistrate names.
child_support_magistrates_array = array("Select one:", "Magistrate 1", "Magistrate 2", "Magistrate 3", "Magistrate 4", "Magistrate 5")

'An array of judges. "Select one:" should ALWAYS be in there, and ALWAYS be first.  Replace "Judge #" with your agency's judges names.
county_judge_array = array("Select one:", "Judge 1", "Judge 2", "Judge 3", "Judge 4", "Judge 5")

'ACTIONS TAKEN BASED ON COUNTY CUSTOM VARIABLES------------------------------------------------------------------------------

'Making a list of offices to be used in various scripts
For each office in county_office_array
	new_office_array = split(office, "|")									'Assigned earlier in the FUNCTIONS FILE script. Splits into an array, containing each line of the address.
	comma_location_in_address_line_02 = instr(new_office_array(1), ",")				'Finds the location of the first comma in the second line of the address (because everything before this is the city)
	city_for_array = left(new_office_array(1), comma_location_in_address_line_02 - 1)		'Pops this city into a variable
	county_office_list = county_office_list & chr(9) & city_for_array					'Adds the city to the variable called "county_office_list", which also contains a new line, so that it works correctly in dialogs.
Next

is_county_collecting_stats = collecting_statistics	'IT DOES THIS BECAUSE THE SETUP SCRIPT WILL OVERWRITE LINES BELOW WHICH DEPEND ON THIS, BY SEPARATING THE VARIABLES WE PREVENT ISSUES

'This loads the user ID for use in determining beta users. May also be used elsewhere in scripts.
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName

'This will assign beta users to the master branch.
If InStr(beta_users, UCASE(windows_user_ID)) <> 0 then use_master_branch = true

'This is the URL of our script repository, and should only change if the agency is a scriptwriting agency. Scriptwriters can elect to use the master branch, allowing them to test new tools, etc.
IF use_master_branch = TRUE THEN		'scriptwriters typically use the master branch
	script_repository = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Script Files/"
ELSE							'Everyone else (who isn't a scriptwriter) typically uses the release branch
	script_repository = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Script Files/"
END IF

'If run locally is set to "True", the scripts will totally bypass GitHub and run locally.
IF run_locally = TRUE THEN script_repository = "C:\PRISM-Scripts\Script Files\"
