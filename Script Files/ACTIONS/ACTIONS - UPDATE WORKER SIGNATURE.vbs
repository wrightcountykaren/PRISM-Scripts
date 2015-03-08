'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - UPDATE WORKER SIGNATURE.vbs"
start_time = timer

worker_signature = InputBox("Please enter what you would like for your default worker signature (NOTE: this will create the signature that is auto-filled as worker signature in scripts)")
IF worker_signature = "" THEN stopscript

Set objNet = CreateObject("WScript.NetWork") 
windows_user_ID = objNet.UserName

SET update_worker_sig_fso = CreateObject("Scripting.FileSystemObject")
SET update_worker_sig_command = update_worker_sig_fso.CreateTextFile("C:\USERS\" & windows_user_ID & "\MY DOCUMENTS\workersig.txt", 2)
update_worker_sig_command.Write(worker_signature)
update_worker_sig_command.Close

script_end_procedure("")
