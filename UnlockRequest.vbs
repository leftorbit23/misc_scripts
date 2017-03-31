`````````````````````````````
`````````````````````````````
`` Set sample file name
``````````````````````````````
`````````````````````````````
strMessage =Inputbox("Enter full path and file name"& vbCrLf & vbCrLf & "Example: " & vbCrLf & " S:\Sales\report.pdf","Input Required")

Set fso = CreateObject("Scripting.FileSystemObject")
If (fso.FileExists(strMessage)) Then

  Set objFSO=CreateObject("Scripting.FileSystemObject")

 
  `````````````````````````````
  `````````````````````````````
  `` Set request file location
  ``````````````````````````````
  `````````````````````````````
  outFile="S:\UnlockRequest\UnlockRequest.txt"
  Set objFile = objFSO.CreateTextFile(outFile,True)
  objFile.Write strMessage & vbCrLf
  objFile.Close
  
  ''send help email notification

  Set wshShell = CreateObject( "WScript.Shell" )
  strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
  strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
  
  
  `````````````````````````````
  `````````````````````````````
  `` Set e-mail information
  ``````````````````````````````
  `````````````````````````````
  strSMTPFrom = "no-reply@example.org"
  strSMTPTo = "helpdesk@example.org"
  strSMTPRelay = "smtp.example.org"
  strSubject = "Share Drive File Unlock Request Processed"
  strTextBody = "User: " & strUserName & vbCrLf & "Computer Name: " & strComputerName & vbCrLf & "File: " &   strMessage 
  'strAttachment = "full UNC path of file"


  Set oMessage = CreateObject("CDO.Message")
  oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
  oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPRelay
  oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
  oMessage.Configuration.Fields.Update

  oMessage.Subject = strSubject
  oMessage.From = strSMTPFrom
  oMessage.To = strSMTPTo
  oMessage.TextBody = strTextBody
  'oMessage.AddAttachment strAttachment

  oMessage.Send


  result=Msgbox("Try reopening the file in one minute",vbInformation, "Unlocking File")

Else
  result=Msgbox("File doesn't exist",vbCritical, "Error")
End If




