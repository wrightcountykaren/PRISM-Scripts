'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NAV - Log into PRISM training region"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Child Support Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'VARIABLES AND CALCULATIONS----------------------------------------------------------------------------------------------------
'PRISM training uses the current month as part of the password. This figures out what it needs to be.
date_for_PW = datepart("m", date) 
If len(date_for_PW) = 1 then date_for_PW = "0" & date_for_PW

'Connects to BlueZone
EMConnect ""

EMReadScreen ADMNET_check, 6, 1, 2
If ADMNET_check <> "ADMNET" then script_end_procedure("You are not in ADMNET (main STATE OF MN screen). The script will now stop.")

EMWriteScreen "cicsdt4", 12, 61
transmit
EMWaitReady 0, 0 'waits as the script might hang
EMWriteScreen "pwcst05", 12, 21
EMWriteScreen "Train#" & date_for_PW, 13, 21
transmit
EMSendKey "QQT4"
transmit
