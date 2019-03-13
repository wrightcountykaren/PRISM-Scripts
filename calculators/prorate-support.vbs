'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "prorate-support.vbs"
start_time = timer
'MANUAL TIME TO COMPLETE THIS SCRIPT IS NEEDED

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
Dim prorate_dialog, number_days, obligation_amt, month_to_prorate, days_in_month, leap_year, prorate_amt, ButtonPressed


'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog prorate_dialog, 0, 0, 216, 130, "Prorate Support"
  Text 5, 5, 100, 15, "Enter the monthly obligation amount to be prorated:"
  ButtonGroup ButtonPressed
    CancelButton 155, 110, 50, 15
  Text 5, 35, 110, 20, "Please select the month you would like to prorate support for:"
  DropListBox 115, 40, 65, 15, "January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", month_to_prorate
  Text 5, 75, 100, 20, "How many days is the party entitled to support?"
  EditBox 115, 80, 35, 15, number_days
  ButtonGroup ButtonPressed
    OkButton 100, 110, 50, 15
  EditBox 115, 5, 65, 15, obligation_amt
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------


Do
	dialog prorate_dialog  'Display dialog

	If buttonpressed = 0 then stopscript 'Cancel button


	'For each month of the year, set the number of days in the month
	If month_to_prorate = "January" then days_in_month = 31
	If month_to_prorate = "February" then
		'If the month selected by the user is February, find out if the calculations should use 28 days or 29 days (for leap years).
		leap_year=MsgBox("Are there 28 days in February? (It's not a leap year?)", 4, "Leap Year?")
		If leap_year = 6 then days_in_month = 28 'if leap_year = 6, user clicked Yes
		If leap_year = 7 then days_in_month = 29 'if leap_year = 7, user clicked No
	End If
	If month_to_prorate = "March" then days_in_month = 31
	If month_to_prorate = "April" then days_in_month = 30
	If month_to_prorate = "May" then days_in_month = 31
	If month_to_prorate = "June" then days_in_month = 30
	If month_to_prorate = "July" then days_in_month = 31
	If month_to_prorate = "August" then days_in_month = 31
	If month_to_prorate = "September" then days_in_month = 30
	If month_to_prorate = "October" then days_in_month = 31
	If month_to_prorate = "November" then days_in_month = 30
	If month_to_prorate = "December" then days_in_month = 31

'Validate that the number of days the user entered is a valid number.  A valid number is greater than 1, but less than the number of days in the month.  The CDBL() command allows the number_days variable
'to be compared as a number, even though it was used as a string. This code repeats until a valid number is entered.

If (days_in_month =< CDbl(number_days)) THEN MsgBox "The number of days to prorate must be less than the number of days in the month."
IF IsNumeric(number_days) = False OR (IsNumeric(number_days)= True AND CDbl(number_days) < 1) then Msgbox "Days to prorate must be a positive number!"

Loop Until days_in_month > CDbl(number_days) and isnumeric(number_days) = true and number_days <> 0


prorate_amt = (obligation_amt/days_in_month) * number_days  'prorated amount is obligation amount, divided by the number of days in the month, multiplied by the number of days to prorate
Msgbox "The prorated obligation for " & number_days & " days in " & month_to_prorate & " is " & formatCurrency(prorate_amt) & "." 'display result in a message box.
