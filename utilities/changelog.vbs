'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "changelog.vbs"
start_time = timer

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

'A temporary MsgBox while we actually build the functionality...
'Script ends
'script_end_procedure("Changelog is coming soon! -Veronica and Robert")

'...SO ROBERT- here's what I came up with:

function update_changelog(date_of_change, text_of_change, scriptwriter_of_change)
	ReDim Preserve changelog(UBound(changelog) + 1)
	changelog(ubound(changelog)) = date_of_change & "|" & text_of_change & "|" & scriptwriter_of_change
end function

function display_changelog

	'Needs to determine MyDocs directory before proceeding.
	Set wshshell = CreateObject("WScript.Shell")
	user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

	'Now it determines the signature
	With (CreateObject("Scripting.FileSystemObject"))															'Creating an FSO
		If .FileExists(user_myDocs_folder & "scripts-last-used-date.txt") Then									'If the workersig.txt file exists...
			Set get_changelog = CreateObject("Scripting.FileSystemObject")										'Create another FSO
			Set changelog_command = get_changelog.OpenTextFile(user_myDocs_folder & "scripts-last-used-date.txt")			'Open the text file
			Do until changelog_command.AtEndOfStream
				strNextLine = changelog_command.ReadLine
				array_script_details = split(strNextLine, "|")
				If instr(array_script_details(1), name_of_script) <> 0 then
					last_date_script_used = array_script_details(0)
					line_to_replace = strNextLine																		'<<<<STORING THIS FOR LATER
				Else
					new_line_to_enter = date & " | " & name_of_script & " | initial use"
				End if

			Loop
		Else
			'Opens an FSO, opens workersig.txt, writes the new signature in, and exits
			SET update_changelog_fso = CreateObject("Scripting.FileSystemObject")
			SET update_changelog_command = update_changelog_fso.CreateTextFile(user_myDocs_folder & "scripts-last-used-date.txt", 2)
			update_changelog_command.Write(date & " | " & name_of_script & " | initial use")
			update_changelog_command.Close
		End if
	END WITH

	'Splitting the changelog into different variables for making things prettier
	For each changelog_entry in changelog
		date_of_change = left(changelog_entry, instr(changelog_entry, "|") - 1)
		If date_of_change > last_date_script_used and last_date_script_used <> "" then
			scriptwriter_of_change = right(changelog_entry, len(changelog_entry) - instrrev(changelog_entry, "|") )
			text_of_change = replace(replace(replace(changelog_entry, scriptwriter_of_change, ""), date_of_change, ""), "|", "")
			changelog_msgbox = changelog_msgbox & "-----" & cdate(date_of_change) & "-----" & vbNewLine & text_of_change & vbNewLine & "Completed by " & scriptwriter_of_change & vbNewLine & vbNewLine
		End if
	Next

	If changelog_msgbox <> "" then
		'Opens an FSO, opens workersig.txt, writes the new signature in, and exits
		SET update_changelog_fso = CreateObject("Scripting.FileSystemObject")
		SET update_changelog_command = update_changelog_fso.CreateTextFile(user_myDocs_folder & "scripts-last-used-date.txt", 2)
		Do until update_changelog_command.AtEndOfStream
			strNextLine = changelog_command.ReadLine
''			If strNextLine = line_to_replace then strNextLine = new_line_to_enter
		Loop

		MsgBox changelog_msgbox

	ElseIf new_line_to_enter <> "" then
		'Opens an FSO, opens workersig.txt, writes the new signature in, and exits
''		SET update_changelog_fso = CreateObject("Scripting.FileSystemObject")
''		SET update_changelog_command = update_changelog_fso.CreateTextFile(user_myDocs_folder & "scripts-last-used-date.txt", 8)

''		update_changelog_command.WriteLine new_line_to_enter
	End if



	'MsgBox line_to_replace

''	MsgBox new_line_to_enter





end function


changelog = array()

'===== CHANGELOG
call update_changelog("10/26/2016", "I did some new things.", "Veronica Cary, DHS")
call update_changelog("10/25/16", "Today a new function was added: the script now has content.", "Robert Fewins-Kalb, Anoka County")

display_changelog



'I'm thinking we could either put this on the list of scripts, or in each individual script. Obviously the former is faster to load while the latter is easier for scriptwriters.
'	Also thinking that we could use this changelog on starting the script: if changes happened in the last day or so, we could alert the worker with a dialog...
