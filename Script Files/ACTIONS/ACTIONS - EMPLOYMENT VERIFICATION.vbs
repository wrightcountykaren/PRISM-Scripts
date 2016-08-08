'Gathering stats
name_of_script = "Action - Employment Verification.vbs"
start_time = timer
STATS_Counter = 1
STATS_manualtime = 300
STATS_denomination = "C"
'End of stats block 

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Robert Kalb, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Robert Kalb and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Robert will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

MsgBox 	"You must be on PANEL ONE of NCID or CPID with the employer you want updated."

BeginDialog Employment_Verification_dialog, 0, 0, 191, 365, "Employment Verification"
  EditBox 65, 10, 45, 15, Income_Type
  EditBox 65, 40, 75, 15, Begin_Date
  EditBox 65, 75, 100, 15, Occupation
  EditBox 65, 105, 75, 15, Verification_Date
  EditBox 70, 145, 40, 15, Verification_Source
  EditBox 70, 180, 70, 15, Wage
  EditBox 70, 200, 45, 15, Frequency
  EditBox 70, 220, 45, 15, Hours_Per_Period
  EditBox 70, 240, 45, 15, Wage_Type
  EditBox 70, 260, 40, 15, Income_Source
  DropListBox 70, 280, 60, 15, "Select one..."+chr(9)+"Y"+chr(9)+"N", Med_cov_dropdown
  DropListBox 70, 300, 60, 15, "Select one..."+chr(9)+"Y"+chr(9)+"N", Den_cov_dropdown
  EditBox 75, 325, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 75, 345, 50, 15
    CancelButton 135, 345, 50, 15
  Text 110, 150, 45, 10, "3 letter code"
  Text 110, 20, 45, 10, "3 letter code"
  Text 5, 110, 55, 10, "Verification Date"
  Text 5, 185, 20, 10, "Wage"
  Text 5, 80, 40, 10, "Occupation:"
  Text 5, 205, 35, 10, "Frequency"
  Text 115, 245, 45, 10, "3 letter code"
  Text 5, 20, 50, 10, "Income Type:"
  Text 5, 225, 60, 10, "Hours Per Period"
  Text 15, 55, 140, 10, "Date must be formated 00/00/0000"
  Text 5, 245, 45, 10, "Wage Type"
  Text 15, 120, 140, 10, "Date must be formated 00/00/0000"
  Text 110, 265, 45, 10, "3 letter code"
  GroupBox 0, 5, 185, 165, "1st Panel"
  GroupBox 0, 170, 185, 150, "2nd Panel"
  Text 5, 265, 55, 10, "Income Source"
  Text 5, 325, 65, 10, "Worker Signature:"
  Text 115, 205, 45, 10, "3 letter code"
  Text 5, 280, 50, 10, "Med Cov Avail"
  Text 5, 45, 40, 10, "Begin Date:"
  Text 5, 300, 55, 10, "Den Cov Avail"
  Text 5, 150, 65, 10, "Verification Source"
EndDialog
EMconnect ""


DO
	err_msg = ""
	Dialog Employment_Verification_dialog
	IF ButtonPressed = 0 THEN StopScript
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "You must sign your CAAD note!" 'If worker sig is blank, message box pops saying you must sign caad note
	If err_msg <> "" THEN msgbox "***NOTICE***" & vbcr & err_msg & vbNewline & vbNewline & "Please resolve for this script to continue."
LOOP UNTIL err_msg = ""

	'IF either of these are "select one" then we are blanking them out so they do not write in the CAAD note 
If Med_cov_dropdown = "Select one..." then Med_cov_dropdown = ""
If Den_cov_dropdown = "Select one..." then Den_cov_dropdown = ""

	'Enters "M" to modify
EMwritescreen "M", 3, 29
	'completes 1st screen of income verification
EMwritescreen Income_Type, 7, 15			'this writes the the type of income they are receiving
EMwritescreen Begin_Date, 10, 14			'this writes the day employment began
EMwritescreen Occupation, 12, 14			'this writes their occupation
EMwritescreen Verification_Date, 20, 7		'this writes the day the information was verified
EMwritescreen Verification_Source, 20, 36 	'this writes who verified the information

EMreadscreen Employer, 30, 9, 17			'this reads the name of the employer so it can be added to the CAAD note

PF11
	'Completes 2nd page of income verification
EMwritescreen Wage, 11, 8				'this writes the amount of pay per pay period
EMwritescreen Frequency, 11, 27			'this is the frequency of the pay periods
EMwritescreen Hours_Per_Period, 12, 20		'this is the amount of hours in a pay period
EMwritescreen Wage_Type, 12, 35 			'this is the type of way they are paid their wages
EMwritescreen Verification_Date, 13, 7		'this is the date the information was verified
Emwritescreen Income_Source, 13, 38			'this is who is paying 
Emwritescreen Med_cov_dropdown, 16, 18		'this is if Med coverage is avail. If not answered it is coded above to leave blank
EMwritescreen Den_cov_dropdown, 17, 18		'this is if Den coverage is avail. If not answered it is coded above to leave blank

If Med_cov_dropdown = "N" then Emwritescreen Verification_Date, 16, 37
If Den_cov_dropdown = "N" then Emwritescreen Verification_Date, 17, 37

Transmit 
	
EmReadscreen success_msg, 21, 24, 24
	IF success_msg <> "modified successfully" THEN script_end_procedure("Problem updating this screen, please verify information!")

PF3

EMwritescreen "CAAD", 21, 18

transmit

EMwritescreen "M", 8, 5

transmit

emsetcursor 19, 4
	'updateds CAAD with information for the dialog box

call write_bullet_and_variable_in_CAAD("Employer", Employer)						'writes the employer name 
call write_bullet_and_variable_in_CAAD("Income Type", Income_Type)					'writes the income type
call write_bullet_and_variable_in_CAAD ("Begin Date", Begin_Date)						'writes the date employment began
call write_bullet_and_variable_in_CAAD ("Verification Date", Verification_Date)			'writes the date the information was verified
call write_bullet_and_variable_in_CAAD ("Verification Source", Verification_Source)			'writes who verified the income
call write_bullet_and_variable_in_CAAD ("Wage", "$" & Wage)							'writes the wage per pay period
call write_bullet_and_variable_in_CAAD ("Frequency", Frequency)						'writes how often they are paid
call write_bullet_and_variable_in_CAAD ("Income Source", Income_Source)					'writes where the income is coming from
call write_bullet_and_variable_in_CAAD ("Medical Coverage", Med_cov_dropdown)				'writes if med coverage is avail, and the date if the med ins is not available
call write_bullet_and_variable_in_CAAD ("Dental Coverage", Den_cov_dropdown)				'writes if den coverage is avail, and the date if the med ins is not available

CALL write_variable_in_CAAD (worker_signature) 

transmit

MsgBox 	"If medical or dental insurance is avaible be sure to update NCPD/CPPD, or request the information" 

























