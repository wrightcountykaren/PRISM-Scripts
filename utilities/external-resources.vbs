'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "external-resources.vbs"
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
MsgBox "External Resources is coming soon! -Veronica and Robert"

'Script ends
script_end_procedure("")

'...SO ROBERT- here's what I came up with:

function update_changelog(date_of_change, text_of_change, scriptwriter_of_change)
	ReDim Preserve changelog(UBound(changelog) + 1)
	changelog(ubound(changelog)) = date_of_change & "|" & text_of_change & "|" & scriptwriter_of_change
end function
changelog = array()

'===== CHANGELOG
call update_changelog("10/26/2016", "I did some new things.", "Veronica Cary, DHS")
call update_changelog("10/25/2016", "Today a new function was added: the script now has content.", "Veronica Cary, DHS")

For each changelog_entry in changelog
	MsgBox changelog_entry
Next

'I'm thinking we could either put this on the list of scripts, or in each individual script. Obviously the former is faster to load while the latter is easier for scriptwriters.
'	Also thinking that we could use this changelog on starting the script: if changes happened in the last day or so, we could alert the worker with a dialog...
