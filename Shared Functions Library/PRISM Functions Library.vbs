'---------------------------------------------------------------------------------------------------
'HOW THIS SCRIPT WORKS:
'
'========== THIS LIBRARY IS DEPRECIATED AND SIMPLY FORWARDS TO THE UNIFIED BZS-FUNCLIB LIBRARY!!!!!
'Some functions are included below for depreciation purposes.

'Defining the script as a PRISM script just to ensure compatibility for the current ALL SCRIPTS.vbs business
PRISM_script = true

'LOADING THE STANDARD LIBRARY FROM GITHUB===================================================================================================

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

'----------------------------------------------------------------------------------------------------DEPRECIATED FUNCTIONS LEFT HERE FOR COMPATIBILITY PURPOSES
function PRISM_check_function													'DEPRECIATED 03/10/2015
	call check_for_PRISM(True)	'Defaults to True because that's how we always did it.
END function

Function save_cord_doc
    If datediff("d", #08/14/2016#, date) > 0 then MsgBox "This function (save_cord_doc) is being depreciated and removed for the September release. If you are seeing this pop-up, it's because you have a script which has this function, and requires updating. It can be replaced with the write_value_and_transmit function. This function must be replaced by September or it may become unavailable entirely. The script will continue."
    EMWriteScreen "M", 3, 29
    transmit
End function

Function send_text_to_DORD(string_to_write, recipient)
    If datediff("d", #08/14/2016#, date) > 0 then MsgBox "This function (send_text_to_DORD) is being depreciated and removed for the September release. If you are seeing this pop-up, it's because you have a script which has this function, and requires updating. It can be replaced with the write_variable_in_DORD function. This function must be replaced by September or it may become unavailable entirely. The script will continue."
    call write_variable_in_DORD(string_to_write, recipient)
End function

Function write_editbox_in_PRISM_case_note(bullet, variable, spaces_count)		'DEPRECIATED 03/10/2015
	call write_bullet_and_variable_in_CAAD(bullet, variable)
End function

Function write_new_line_in_PRISM_case_note(variable)							'DEPRECIATED 03/10/2015
	call write_variable_in_CAAD(variable)
End function

FUNCTION write_value_and_transmit(input_value, PRISM_row, PRISM_col)
	EMWriteScreen input_value, PRISM_row, PRISM_col
	transmit
END FUNCTION

Function write_variable_to_CORD_paragraph(variable)
	If trim(variable) <> "" THEN
		EMGetCursor noting_row, noting_col		'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
		noting_col = 6					'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
		IF noting_row < 11 THEN noting_row = 11	'Making sure it is writing in the paragraph.

		'Backing out of the CORD paragraph
		IF noting_row > 20 THEN
			MsgBox "The script is attempting to write in a spot that is not supported by PRISM. Please review your CORD document for accuracy and contact a scripts administrator to have this issue resolved.", vbCritical + vbSystemModal, "Critical CORD Paragraph Error!!"
			EXIT FUNCTION
		END IF

		'Splits the contents of the variable into an array of words
		variable_array = split(variable, " ")

		FOR EACH word IN variable_array

			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 75 then
				noting_row = noting_row + 1
				noting_col = 6
			End if

			'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)

			'Backing out of the CORD paragraph
			IF noting_row >= 20 THEN
				MsgBox "The script is attempting to write in a spot that is not supported by PRISM. Please review your CORD document for accuracy and a scripts administrator to have this issue resolved.", vbCritical + vbSystemModal, "Critical CORD Paragraph Error!!"
				EXIT FUNCTION
			END IF
		NEXT

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 6
	End if
End function

'>>>>> CLASSES!!!!!!!!!!!!!!!!!!!!! <<<<<
'This CLASS contains properties used to populate documents
' These properties should not be used for other applications in scripts.
' Everytime you call the property, the script will try to navigate and grab the information
CLASS doc_info
	' >>>>>>>>>>>>><<<<<<<<<<<<<
	' >>>>> CP INFORMATION <<<<<
	' >>>>>>>>>>>>><<<<<<<<<<<<<
	' CP name (last, first middle initial, suffix (if any))
	PUBLIC PROPERTY GET cp_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_name, 50, 5, 25
		cp_name = trim(cp_name)
	END PROPERTY

	' CP first name
	PUBLIC PROPERTY GET cp_first_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_first_name, 12, 8, 34
		cp_first_name = trim(replace(cp_first_name, "_", ""))
	END PROPERTY

	' CP last name
	PUBLIC PROPERTY GET cp_last_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_last_name, 17, 8, 8
		cp_last_name = trim(replace(cp_last_name, "_", ""))
	END PROPERTY

	' CP middle name
	PUBLIC PROPERTY GET cp_middle_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_middle_name, 12, 8, 56
		cp_middle_name = trim(replace(cp_middle_name, "_", ""))
	END PROPERTY

	' CP middle initial
	PUBLIC PROPERTY GET cp_middle_initial
		cp_middle_initial = left(cp_middle_name, 1)
	END PROPERTY

	' CP suffix
	PUBLIC PROPERTY GET cp_suffix
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_suffix, 3, 8, 74
		cp_suffix = trim(replace(cp_suffix, "_", ""))
	END PROPERTY

	' CP date of birth
	PUBLIC PROPERTY GET cp_dob
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_dob, 8, 6, 24
	END PROPERTY

	' CP social security number
	PUBLIC PROPERTY GET cp_ssn
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_ssn, 11, 6, 7
	END PROPERTY

	' CP MCI
	PUBLIC PROPERTY GET cp_mci
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDE" THEN CALL navigate_to_PRISM_screen("CPDE")
		EMReadScreen cp_mci, 10, 5, 7
	END PROPERTY

	' CP address
	PUBLIC PROPERTY GET cp_addr
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadScreen cp_addr1, 30, 15, 11
		EMReadScreen cp_addr2, 30, 16, 11
		cp_addr = replace(cp_addr1, "_", "") & ", " & replace(cp_addr2, "_", "")
	END PROPERTY

	' CP address city
	PUBLIC PROPERTY GET cp_city
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadScreen cp_city, 20, 17, 11
		cp_city = replace(cp_city, "_", "")
	END PROPERTY

	' CP address state
	PUBLIC PROPERTY GET cp_state
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadScreen cp_state, 2, 17, 39
	END PROPERTY

    ' CP address zip code
	PUBLIC PROPERTY GET cp_zip
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CPDD" THEN CALL navigate_to_PRISM_screen("CPDD")
		EMReadScreen cp_zip, 10, 17, 50
	END PROPERTY

	' >>>>>>>>>>>>><<<<<<<<<<<<<<
	' >>>>> NCP Information <<<<<
	' >>>>>>>>>>>>><<<<<<<<<<<<<<
	' NCP Name
	PUBLIC PROPERTY GET ncp_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_name, 50, 5, 25
		ncp_name = trim(ncp_name)
	END PROPERTY

	' NCP first name
	PUBLIC PROPERTY GET ncp_first_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_first_name, 12, 8, 34
		ncp_first_name = trim(replace(ncp_first_name, "_", ""))
	END PROPERTY

	' NCP last name
	PUBLIC PROPERTY GET ncp_last_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_last_name, 17, 8, 8
		ncp_last_name = trim(replace(ncp_last_name, "_", ""))
	END PROPERTY

	' NCP middle name
	PUBLIC PROPERTY GET ncp_middle_name
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_middle_name, 12, 8, 56
		ncp_middle_name = trim(replace(ncp_middle_name, "_", ""))
	END PROPERTY

	' NCP middle initial
	PUBLIC PROPERTY GET ncp_middle_initial
		ncp_middle_initial = left(ncp_middle_name, 1)
	END PROPERTY

	' NCP suffix
	PUBLIC PROPERTY GET ncp_suffix
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_suffix, 3, 8, 74
		ncp_suffix = trim(replace(ncp_suffix, "_", ""))
	END PROPERTY

	' NCP date of birth
	PUBLIC PROPERTY GET ncp_dob
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_dob, 8, 6, 24
	END PROPERTY

	' NCP SSN
	PUBLIC PROPERTY GET ncp_ssn
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_ssn, 11, 6, 7
	END PROPERTY

	' NCP MCI
	PUBLIC PROPERTY GET ncp_mci
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDE" THEN CALL navigate_to_PRISM_screen("NCDE")
		EMReadScreen ncp_mci, 10, 5, 7
	END PROPERTY

	' NCP street address
	PUBLIC PROPERTY GET ncp_addr
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen ncp_addr1, 30, 15, 11
		EMReadScreen ncp_addr2, 30, 16, 11
		ncp_addr = replace(ncp_addr1, "_", "") & ", " & replace(ncp_addr2, "_", "")
	END PROPERTY

	' NCP address city
	PUBLIC PROPERTY GET ncp_city
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen ncp_city, 20, 17, 11
		ncp_city = replace(ncp_city, "_", "")
	END PROPERTY

	' NCP address state
	PUBLIC PROPERTY GET ncp_state
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen ncp_state, 2, 17, 39
	END PROPERTY

	' NCP address zip code
	PUBLIC PROPERTY GET ncp_zip
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "NCDD" THEN CALL navigate_to_PRISM_screen("NCDD")
		EMReadScreen ncp_zip, 10, 17, 50
	END PROPERTY

	' >>>>>>>>>>>>>>><<<<<<<<<<<<<<
	' >>> Financial Information <<<
	' >>>>>>>>>>>>>>><<<<<<<<<<<<<<
	' monthly accrual amount
	PUBLIC PROPERTY GET monthly_accrual
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen monthly_accrual, 8, 9, 31
		monthly_accrual = trim(monthly_accrual)
	END PROPERTY

	' monthly non-accrual
	PUBLIC PROPERTY GET monthly_non_accrual
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen monthly_non_accrual, 8, 10, 31
		monthly_non_accrual = trim(monthly_non_accrual)
	END PROPERTY

	' NPA arrears
	PUBLIC PROPERTY GET npa_arrears
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen npa_arrears, 8, 9, 70
		npa_arrears = trim(npa_arrears)
	END PROPERTY

	' PA arrears
	PUBLIC PROPERTY GET pa_arrears
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen pa_arrears, 8, 10, 70
		pa_arrears = trim(pa_arrears)
	END PROPERTY

	' Total arrears
	PUBLIC PROPERTY GET ttl_arrears
		EMReadScreen at_screen, 4, 21, 75
		IF at_screen <> "CAFS" THEN CALL navigate_to_PRISM_screen("CAFS")
		EMReadScreen ttl_arrears, 8, 11, 70
		ttl_arrears = trim(ttl_arrears)
	END PROPERTY
END CLASS
