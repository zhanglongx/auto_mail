' Version: 1.21

Const RECORD  = "Record.txt"
Const MAIL_TO = "lixin@dvt.dvt.com"
Const MAIL_CC = "xiyan@dvt.dvt.com"

Dim iCount
iCount = 0

Set objFSO = Createobject ("Scripting.FileSystemObject")
If (objFSO.FileExists(RECORD)) Then

    ' Read last record, empty or modify time not today will both
    ' lead to count = 0
    Set objFile = objFSO.GetFile(RECORD)

    If DateDiff("d", objFile.DateLastModified, Now) < 1 Then

        set textFile = objFSO.OpenTextFile(RECORD)
        ' FIXME:
        iCount = textFile.ReadLine
        textFile.close

    End If

    set objFile = Nothing

End If

Dim iResult 
iResult = MsgBox("Happy time? ( Already " & iCount & " Today )", vbYesNo)
If iResult = vbYes Then

    Randomize

    WScript.Sleep 60 * 1000

    Call Mail_Outlook("外出吸烟")

    Dim delay

    delay = Int(150 + Rnd * 120) * 1000

    WScript.Sleep delay

    Call Mail_Outlook("外出吸烟返回")

    iCount = iCount + 1

    Set textFile = objFSO.CreateTextFile(RECORD,True)
    textFile.Write iCount & vbCrLf
    textFile.Close

End If

set objFSO = Nothing

Wscript.Quit

Sub Mail_Outlook(ByVal theSubject)

   Set OutApp = CreateObject("Outlook.Application")
   Set OutMail = OutApp.CreateItem(0)
   
   With OutMail
       .to = MAIL_TO
       .CC = MAIL_CC
       .Subject = theSubject
       .Send
   End With

   Set OutMail = Nothing
   Set OutApp  = Nothing

End Sub
