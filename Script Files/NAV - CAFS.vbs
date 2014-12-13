'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NAV - CS - CAFS"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Child Support Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


EMConnect ""

PRISM_check_function


call navigate_to_PRISM_screen("CAFS")


script_end_procedure("")
