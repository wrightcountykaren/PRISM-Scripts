'LOADING SCRIPT
url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
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

'EXCEL BLOCK
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 'Set this to False to make the Excel spreadsheet go away. This is necessary in production.
Set objWorkbook = objExcel.Workbooks.Add() 
objExcel.DisplayAlerts = True 'Set this to false to make alerts go away. This is necessary in production.

Dim CAFS_checkbox
BeginDialog CALI_to_excel_Dialog, 0, 0, 196, 70, "CALI To Excel"
  CheckBox 5, 10, 115, 10, "Check here if you want to include total arrears, monthly accrual amount, and non-accrual amount to Excel", CAFS_checkbox
  Text 15, 20, 110, 30, "total arrears, monthly accrual amount, and non-accrual amount to Excel from CAFS. "
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 10, 50, 120, 15, "(This takes more time to process)"
EndDialog

'Connects to Bluezone
EMConnect ""

					DIALOG CALI_to_excel_Dialog
     					IF ButtonPressed = 0 THEN StopScript

PF3

'Goes to CALI
EMWriteScreen "CALI", 21,18
Transmit
'Blanks out case number and brings you to the top of CALI
EMWriteScreen "              ", 20, 58
Transmit

'sets row to fill info into Excel
excel_row = 2


Do 'Loops script until the end of CALI
	prism_row = 8
	Do 'Copies Case Number, Function Type, Program Type, CP Name, and NCP Name to the Excel document
	
	
		EMReadScreen prism_case_number, 14, prism_row, 7 'Reads and copies case number
		EMReadScreen function_type, 2, prism_row, 23 'Reads and copies function type
		EMReadScreen program_type, 3, prism_row, 27 'Reads and copies program type
		EMReadScreen interstate_code, 1, prism_row, 33 'Reads and copies intersate code
		EMReadScreen CP_name, 26, prism_row, 38 'Reads and copies CP name
		pf11
		EMReadScreen NCP_name, 26, prism_row, 33 'Reads and copies NCP name
		pf10
	
		'Set rows in Excel for case number, funtion type, program type, CP name, and NCP name
		ObjExcel.Cells(excel_row, 1).Value = prism_case_number
		ObjExcel.Cells(excel_row, 2).Value = function_type
		ObjExcel.Cells(excel_row, 3).Value = program_type
		ObjExcel.Cells(excel_row, 4).Value = interstate_code
		ObjExcel.Cells(excel_row, 5).Value = CP_name
		ObjExcel.Cells(excel_row, 6).Value = NCP_name

		prism_row = prism_row + 1
		excel_row = excel_row + 1

		EmReadscreen end_of_data_check, 11, prism_row, 32
		If end_of_data_check = "End of Data" then exit do

	Loop Until prism_row = 19

	pf8

Loop Until end_of_data_check = "End of Data"

EMWriteScreen "PALC", 21, 18 
Transmit

excel_row = 2

Do
	prism_case_number = Trim(ObjExcel.Cells(excel_row, 1).Value)
	EMWriteScreen Left (prism_case_number, 10), 20, 9
	EMWriteScreen Right (prism_case_number, 2), 20, 20
	Transmit
	EMReadScreen last_payment_date, 8, 9, 59
	If last_payment_date = "        " then last_payment_date = "No Payments"
	ObjExcel.Cells(excel_row, 7).Value = last_payment_date
	excel_row = excel_row + 1

	

Loop until prism_case_number = ""

If CAFS_checkbox = checked then

EMWriteScreen "CAFS", 21, 18 
Transmit

excel_row = 2

Do
	prism_case_number = Trim(ObjExcel.Cells(excel_row, 1).Value)
	EMWriteScreen Left (prism_case_number, 10), 4, 8
	EMWriteScreen Right (prism_case_number, 2), 4, 19
	EMWriteScreen "D", 3, 29
	Transmit 
	EMReadScreen amount_of_arrears, 10, 12, 68
	ObjExcel.Cells(excel_row, 8).Value = amount_of_arrears
	EMReadScreen monthly_accrual, 7, 9, 32
	ObjExcel.Cells(excel_row, 9).Value = monthly_accrual
	EMReadScreen monthly_non_accrual, 7, 10, 32
	ObjExcel.Cells(excel_row, 10).Value = monthly_non_accrual
	excel_row = excel_row + 1
	

Loop until prism_case_number = ""

End If




ObjExcel.Cells(1, 1).Value = "Case Number"
ObjExcel.Cells(1, 2).Value = "Function"
ObjExcel.Cells(1, 3).Value = "Program"
ObjExcel.Cells(1, 4).Value = "Interstate?"
ObjExcel.Cells(1, 5).Value = "CP Name"
ObjExcel.Cells(1, 6).Value = "NCP Name"
ObjExcel.Cells(1, 7).Value = "Last Payment Date"
If CAFS_checkbox = checked then
ObjExcel.Cells(1, 8).Value = "Amount Of Arrears"
ObjExcel.Cells(1, 9).Value = "Monthly Accrual"
ObjExcel.Cells(1, 10).Value = "Monthly Non-Accrual"
End If

'Autofitting columns
For col_to_autofit = 1 to 10
	ObjExcel.columns(col_to_autofit).AutoFit()
Next
