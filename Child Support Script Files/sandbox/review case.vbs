
 'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = ""
start_time = timer

 'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'FUNCTIONS (MERGE INTO MAIN FUNCTIONS BEFORE GO-LIVE)----------------------------------------------------------------------------------------------------
Function convert_Proc_Type_to_PAYMENT_type(PAYMENT_type_code, variable)

	If Proc_type = "01" then variable = APP
	If Proc_type = "02" then variable = BND
	If Proc_type = "03" then variable = FAO
	If Proc_type = "04" then variable = FIN
	If Proc_type = "05" then variable = FTJ
	If Proc_type = "06" then variable = FTS
	If Proc_type = "07" then variable = IFC
	If Proc_type = "08" then variable = INW
	If Proc_type = "09" then variable = NOC
	If Proc_type = "10" then variable = ORE
	If Proc_type = "11" then variable = OST
	If Proc_type = "12" then variable = OWA
	If Proc_type = "13" then variable = PCA
	If Proc_type = "14" then variable = PIF
	If Proc_type = "15" then variable = REG
	If Proc_type = "16" then variable = REI
	If Proc_type = "17" then variable = REO
	If Proc_type = "18" then variable = STJ
	If Proc_type = "19" then variable = STS
	If Proc_type = "20" then variable = WOC

End function
 
BeginDialog Review, 0, 0, 201, 275, "Review"
  Text 5, 10, 50, 15, "Total Due"
  EditBox 5, 30, 70, 20, Total_due
  Text 5, 70, 60, 15, "Last Payment (PALC)"
  EditBox 5, 100, 70, 25, Last_payment
  Text 5, 135, 45, 10, "Pymnt Type"
  Text 105, 20, 65, 10, "Arrears"
  CheckBox 90, 30, 105, 15, "Greater than 1 month current", Arrears
  Text 105, 75, 70, 10, "Date of last payment"
  EditBox 100, 100, 75, 25, Edit5
  CheckBox 140, 165, 55, 10, "CCPA", CCPA
  CheckBox 140, 180, 55, 15, "120 percent in place", Current_plus_20_percent
  ButtonGroup ButtonPressed
    CancelButton 85, 255, 50, 15
    OkButton 140, 255, 50, 15
  EditBox 0, 190, 120, 15, Enforcement
  Text 10, 180, 100, 10, "Additional Narrative"
  EditBox 50, 135, 65, 15, payment_type
  CheckBox 10, 215, 70, 20, "Is NCP compliant?", NCP_compliant
EndDialog




'The Script--------------------------------------------------------------------------------------------------------------------------------------------------


'Connects to BlueZone
EMConnect ""

'Searches for the case number.
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if

'<<<<A TEMPORARY MSGBOX TO CHECK THE ACCURACY OF THE PRISM CASE NUMBER FINDER. IF THIS WORKS CREATE A CUSTOM FUNCTION OUT OF THE ABOVE CODE
If PRISM_case_number <> "" then MsgBox "A case number was automatically found on this screen! It is indicated as: " & PRISM_case_number & ". If this case number is incorrect, please take a screenshot of PRISM and send a description of what's wrong to Veronica Cary."

'Shows case number dialog
Do
	Do
		Dialog PRISM_case_number_dialog
		If buttonpressed = 0 then stopscript
		call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	Loop until case_number_valid = True
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"


'Autofilling from PRISM--------------------------------------------------------------------------------------------------------------------------------------------------

'Getting info from CAFS'

	call navigate_to_PRISM_screen("CAFS")
transmit
	EMSetcursor 14, 30
	EMReadScreen total_due, 9, 14, 30
	EMsendKey replace (total_due, "-" , "")











