'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = ""
start_time = timer
'
''LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Child Support Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

BeginDialog salutation_dialog, 0, 0, 191, 105, "Enter your salutation"
  Text 10, 5, 60, 10, "Case number"
  EditBox 10, 15, 80, 15, case_number
  Text 10, 40, 60, 10, "Salutation"
  EditBox 10, 50, 80, 15, hello
  ButtonGroup ButtonPressed
    OkButton 135, 55, 50, 15
    CancelButton 135, 75, 50, 15
EndDialog


'Connecting to BlueZone
EMConnect ""

'Display dialog
Dialog salutation_dialog
If buttonpressed = 0 then stopscript

'Going to CAAD note
call navigate_to_PRISM_screen("CAAD")

'Entering case number
EMWriteScreen case_number, 20, 8


PF5					'Did this because you have to add a new note
EMSetCursor 16, 4			'Because the cursor does not default to this location
call write_editbox_in_PRISM_case_note("Salutation", hello, 6)