IF jude_checkbox = 1 THEN
	'CP Name											
	call navigate_to_PRISM_screen("CPDE")
	EMWriteScreen CP_MCI, 4, 7
	EMReadScreen CP_F, 12, 8, 34
	EMReadScreen CP_M, 12, 8, 56
	EMReadScreen CP_L, 17, 8, 8

	CP_name = fix_read_data(CP_F) & " " & fix_read_data(CP_M) & " " & fix_read_data(CP_L)	
	CP_name = trim(CP_Name)


	CALL navigate_to_PRISM_screen ("SUOD")
	EMWriteScreen "B", 3, 29
	transmit

	BeginDialog PRISM_INFO_Dialog, 0, 0, 266, 185, "Info needed to add Un/Un to PRISM"
	  EditBox 85, 25, 25, 15, CO_Seq
	  EditBox 50, 45, 50, 15, From_date
	  EditBox 130, 45, 50, 15, To_date
		EditBox 50, 70, 200, 15, CP_name
	  EditBox 65, 110, 40, 15, eff_date
	  EditBox 55, 135, 50, 15, beg_date
	  ButtonGroup ButtonPressed
				OkButton 145, 160, 50, 15
				CancelButton 205, 160, 50, 15
	  Text 101, 10, 65, 10, "JUDE Information"
	  Text 10, 30, 70, 10, "Court Order Seq Nbr:"
	  Text 120, 30, 35, 10, "format 01"
	  Text 10, 50, 40, 10, "Date From:"
	  Text 110, 50, 15, 10, "To:"
	  Text 190, 50, 50, 10, "xx/xx/xxxx"
	  Text 10, 75, 40, 10, "In Favor of:"
	  Text 101, 95, 65, 10, "NCOD Information"
	  Text 10, 115, 55, 10, "Effective Date:"
	  Text 130, 115, 50, 10, "xx/xxxx"
	  Text 10, 140, 40, 10, "Begin Date:"
	  Text 120, 140, 50, 10, "xx/xx/xxxx"
	EndDialog


	Do
		err_msg = ""
		Dialog PRISM_INFO_Dialog
		IF buttonpressed = 0 then stopscript
		IF Co_Seq = "" THEN err_msg = err_msg & vbNewline & "Please enter the Court order sequence number."
		IF From_date = "" THEN err_msg = err_msg & vbNewline & "Please enter FROM date."
		IF To_date = "" THEN err_msg = err_msg & vbNewline & "Please enter TO date."
		IF CP_name = "" THEN err_msg = err_msg & vbNewline & "Please enter the CP's name."
		IF eff_date = "" THEN err_msg = err_msg & vbNewline & "Please enter the effective date."
		IF beg_date = "" THEN err_msg = err_msg & vbNewline & "Please enter the begin date."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF
	Loop until err_msg = ""

	'adding jude info
	CALL navigate_to_PRISM_screen ("JUDE")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen Co_Seq, 4, 34
	EMWriteScreen "JME", 10, 6
	EMWriteScreen From_date, 10, 17
	EMWriteScreen To_date, 10, 31
	EMWriteScreen CP_name, 13, 16
	EMWriteScreen amount, 14, 17
	EMWriteScreen "JOL", 15, 20
	PF11
	EMWriteScreen "un/un expenses requested by cp", 12, 3
	transmit

	'checking bottom screen for jol success
	EMReadScreen jol_success, 18, 24, 33
	IF jol_success <> "added successfully" THEN 
		script_end_procedure ("Jude information was not added correctly, please reneter information.  Script Ended.")
	END IF

	'reading judgment sequence number to add to ncod
	EMReadScreen jdgmt_number, 2, 4, 52
		
	'adding ncod info
	CALL navigate_to_PRISM_screen ("NCOD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen "JME", 4, 34
	EMWriteScreen "  ", 4, 053
	EMWriteScreen eff_date, 9, 59 
	EMWriteScreen "npa", 12, 10
	EMWriteScreen Co_Seq, 11, 62
	EMWriteScreen "n", 13, 12
	EMWriteScreen Co_Seq, 12, 55
	EMWriteScreen jdgmt_number, 12, 74
	EMWriteScreen "y", 18, 57
	EMWriteScreen beg_date, 14, 68 
	transmit
	
	
	'reading ncod success
	EMReadScreen ncod_success, 18 , 24, 34
	IF ncod_success <> "added successfully" THEN 
		ncod_message = Msgbox ("NCOD information was not added correctly, please correct error and click OK to continue. click CANCEL to end script.", VbOKCancel)
		If ncod_message = vbCancel then stopscript
	END IF

	'adding obbd info
	CALL navigate_to_PRISM_screen ("OBBD")
	EMWriteScreen "M", 3, 29
	EMWriteScreen "           ", 18, 15
	EMWriteScreen amount, 18, 15
	PF11
	EMWriteScreen "added un/un expenses. " & worker_signature, 18, 25
	EMWriteScreen "n", 17, 72 
	transmit

	'reading modified sucess
	EMReadScreen obbd_success, 13 , 24, 68
	IF obbd_success <> "modified succ" THEN 
		Msgbox "OBBD information was not added correctly, please reneter information.  Script Ended."
		StopScript
	END IF

	CALL navigate_to_PRISM_screen ("NCOL")

END IF
