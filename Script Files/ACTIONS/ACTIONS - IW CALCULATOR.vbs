
'Option Explicit 'this has to be on the top, always
'Option Explicit

'this is a function document
'DIM beta_agency 'remember to add

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO                                                                          'Declares variables to be good to option explicit users
If beta_agency = "" then                                              'For scriptwriters only
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then                 'For beta agencies and testers
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else                                                                                                                        'For most users
                url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")                                                               'Creates an object to get a URL
req.open "GET", url, False                                                                                                                                            'Attempts to open the URL
req.send                                                                                                                                                                                                              'Sends request
If req.Status = 200 Then                                                                                                                                                '200 means great success
                Set fso = CreateObject("Scripting.FileSystemObject")    'Creates an FSO
                Execute req.responseText                                                                                                                          'Executes the script code
ELSE                                                                                                                                                                                                                       'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
                MsgBox                "Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
                                                vbCr & _
                                                "Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
                                                vbCr & _
                                                "If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
                                                vbTab & "- The name of the script you are running." & vbCr &_
                                                vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
                                                vbTab & "- The name and email for an employee from your IT department," & vbCr & _
                                                vbTab & vbTab & "responsible for network issues." & vbCr &_
                                                vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
                                                vbCr & _
                                                "Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
                                                vbCr &_
                                                "URL: " & url
                                                StopScript
END IF
'this is where the copy and paste from functions library ended



'DIALOGS---------------------------------------------------------------------------
'DIM IW_CALC_Dialog, PRISM_case_number, Current_Support, Percent, Manual, Other_Amount, err_msg, ButtonPressed, case_number_is_valid, MoTotal

BeginDialog IW_CALC_Dialog, 0, 0, 177, 156, "IW CALC Dialog"
  EditBox 60, 0, 100, 20, PRISM_case_number
  EditBox 100, 20, 60, 20, Current_Support
  CheckBox 10, 60, 60, 10, "20 Percent", Percent
  CheckBox 10, 80, 60, 10, "Other Amount", Other_Amount
  CheckBox 10, 100, 140, 10, "Add 30-day FREE worklist?", cawd_check
  EditBox 70, 70, 50, 20, Manual
  ButtonGroup ButtonPressed
    OkButton 30, 130, 50, 20
    CancelButton 100, 130, 50, 20
  Text 0, 10, 50, 10, "Case Number"
  Text 0, 30, 90, 10, "Current Monthly Obligation "
  Text 0, 50, 80, 10, "Arrears Collection Rate"
EndDialog


'THE SCRIPT-----------------------------------

'Connecting to BlueZone
EMConnect ""


CALL navigate_to_PRISM_screen ("CAFS")

'variable name in edit box that i want autofilled
EMReadScreen PRISM_case_number, 13, 4, 8
EMReadScreen Current_Support, 10, 9, 29

'adding a loop
Do
	err_msg = ""
	Dialog IW_CALC_Dialog				'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		IF PRISM_case_number = "" THEN err_msg = err_msg & vbNewline & "Prism case number must be completed."
		IF Current_Support = "" THEN err_msg = err_msg & vbNewline & "Current Support must be completed"
		'IF CP = 0 AND NCP = 0 THEN err_msg = vbNewline & "Either CP or NCP must be selected."

		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & "Please resolve for the script to continue."
		END IF
LOOP UNTIL err_msg = ""

'Converting accrual amts to number from string 
Current_Support = CDbl(Current_Support)

IF Percent = checked THEN MoTotal = Current_Support * 1.2
IF Other_Amount = checked THEN MoTotal = Current_Support + Manual
IF Percent = 0 AND Other_Amount = 0 THEN MoTotal = Current_Support

'Calculating pay period amounts
Dim WeekPay, BiWeekPay, SemiMoPay

WeekPay = MoTotal * 12/52
WeekPay = FormatNumber(WeekPay, 2)
BiWeekPay = MoTotal * 12/26
BiWeekPay = FormatNumber(BiWeekPay, 2)
SemiMoPay = MoTotal/2
SemiMoPay = FormatNumber(SemiMoPay, 2)

IF cawd_check = checked THEN
CALL navigate_to_PRISM_screen ("CAWD")
PF5
EMWriteScreen "A", 3, 30
EMWriteScreen "FREE", 4, 37
EMWriteScreen "30", 17, 52

EMSetCursor 10, 4
CALL write_variable_in_CAAD ("Monthly: " & FormatCurrency(MoTotal))
CALL write_variable_in_CAAD ("Bi-Weekly: " & FormatCurrency(BiWeekPay))
CALL write_variable_in_CAAD ("Semi-Monthly: " & FormatCurrency(SemiMoPay))
CALL write_variable_in_CAAD ("Weekly: " & FormatCurrency(WeekPay))


ELSE
 
'msgbox needed to show calculations, weekly, biweekly, semi monthly, and monthly
MsgBox ("Monthly: $" & MoTotal & VbNewline & VbNewline & _
	"Weekly: $" & WeekPay & VbNewline & VbNewline & _
	"Bi-Weekly: $" & BiWeekPay & VbNewline & VbNewline & _
	"Semi-Monthly: $" & SemiMoPay)
END IF
script_end_procedure("")
