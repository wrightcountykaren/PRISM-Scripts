'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NAV - PRISM Obscure Screens"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Child Support Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

BeginDialog PRISM_obscure_screens_dialog, 0, 0, 161, 185, "PRISM Obscure screens"
  ButtonGroup ButtonPressed
    OkButton 45, 160, 50, 15
    CancelButton 100, 160, 50, 15
    PushButton 115, 5, 35, 10, "ACSD", ACSD_button
    PushButton 115, 20, 35, 10, "CPRE", CPRE_button
    PushButton 115, 35, 35, 10, "GCSC", GCSC_button
    PushButton 115, 50, 35, 10, "NCLD", NCLD_button
    PushButton 115, 65, 35, 10, "NCLL", NCLL_button
    PushButton 115, 80, 35, 10, "NCSL", NCSL_button
    PushButton 115, 95, 35, 10, "PALI", PALI_button
    PushButton 115, 110, 35, 10, "REID", REID_button
    PushButton 115, 125, 35, 10, "SEPD", SEPD_button
    PushButton 115, 140, 35, 10, "WEDL", WEDL_button
  Text 10, 5, 70, 10, "Account status detail"
  Text 10, 20, 70, 10, "CP/NCP Relationship"
  Text 10, 35, 70, 10, "Good Cause Screen"
  Text 10, 50, 80, 10, "NCP license data detail"
  Text 10, 65, 70, 10, "NCP license data list"
  Text 10, 80, 70, 10, "NCP alias detail"
  Text 10, 95, 70, 10, "Payment listing"
  Text 10, 110, 90, 10, "Re-employment ins detail"
  Text 10, 125, 90, 10, "Service of Process detail"
  Text 10, 140, 95, 10, "On-line employer reporting"
EndDialog

'The Script--------------------------------------------------------------------------------------------------

Dialog PRISM_obscure_screens_dialog

'Connect to BlueZone
EMConnect ""

PRISM_check_function			'this checks whether in PRISM or timed out

'Not adding case number function as PRISM will go to screen without case number or PRISM populates with last case number used.

'Naviagting to any of the buttons chosen
If buttonpressed = ACSD_button then call navigate_to_PRISM_screen("ACSD")   
If buttonpressed = CPRE_button then call navigate_to_PRISM_screen("CPRE")
If buttonpressed = GCSC_button then call navigate_to_PRISM_screen("GCSC")   
If buttonpressed = NCLD_button then call navigate_to_PRISM_screen("NCLD")
If buttonpressed = NCLL_button then call navigate_to_PRISM_screen("NCLL")   
If buttonpressed = NCSL_button then call navigate_to_PRISM_screen("NCSL")
If buttonpressed = PALI_button then call navigate_to_PRISM_screen("PALI")   
If buttonpressed = REID_button then call navigate_to_PRISM_screen("REID")
If buttonpressed = SEPD_button then call navigate_to_PRISM_screen("SEPD")   
If buttonpressed = WEDL_button then call navigate_to_PRISM_screen("WEDL")

script_end_procedure("")

