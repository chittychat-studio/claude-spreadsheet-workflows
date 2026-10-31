Attribute VB_Name = "SendRotaEmail"
Option Explicit

' sheets30 reference macro (2026-10-31). Opens an Outlook DRAFT with this
' workbook attached. It never sends: you check the draft and press Send yourself.
'
' Needs: desktop Excel on Windows + the classic desktop Outlook app.
' The new Outlook for Windows and Outlook on the web can't be driven by a macro.
'
' To use: open the workbook, press Alt+F11, Insert > Module, paste this in,
' save as "Excel Macro-Enabled Workbook (*.xlsm)", then run SendRotaEmail.

Sub SendRotaEmail()
    Const RECIPIENT As String = "coordinator@example.org"   ' change to the real address
    Const SUBJECT_LINE As String = "Volunteer rota for next month"

    Dim olApp As Object
    Dim mail As Object

    ' 1. The workbook must be saved somewhere, so there is a file to attach.
    If ThisWorkbook.Path = "" Then
        MsgBox "Save this workbook first (as a macro-enabled .xlsm), then run the macro again.", vbExclamation
        Exit Sub
    End If

    ' 2. Save, so the attachment is the latest version.
    ThisWorkbook.Save

    ' 3. Use Outlook if it's already open, otherwise start it.
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
        MsgBox "Couldn't start desktop Outlook. This macro needs the classic Outlook app.", vbCritical
        Exit Sub
    End If

    ' 4. Build the email: recipient, subject, message, attachment.
    Set mail = olApp.CreateItem(0)   ' 0 = a new mail item
    With mail
        .To = RECIPIENT
        .Subject = SUBJECT_LINE
        .Body = "Hi," & vbCrLf & vbCrLf & _
                "Here's the volunteer rota for next month, attached." & vbCrLf & _
                "Every rule has been checked. Shout if anything looks wrong." & vbCrLf & vbCrLf & _
                "Thanks," & vbCrLf & "Ravi"
        .Attachments.Add ThisWorkbook.FullName

        ' 5. Show the draft. Never .Send: a person checks it and presses Send.
        .Display
    End With
End Sub
