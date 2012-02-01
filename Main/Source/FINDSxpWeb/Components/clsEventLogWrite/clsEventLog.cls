VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Private oMSEmail As MailSender
Const SMTP_SERVER = "172.19.0.213" 'DUBEXCHANGE.PACER.LOCAL

' Writing to the Windows 2000 Event Log is a powerful feature of the ASP+ and .NET
'   framework.  For those individuals who work in a large company and want to make
'   track application errors, writing to the event log is a must! One of the major
'   shortfalls of asp was the process of making this happen.
Public Sub EventLogWrite(str_in_ErrorInfo As String)
'  Dim oMSEmail As MailSender
  Dim sBody As String

'  Set oMSEmail = New MailSender
'  oMSEmail.Host = SMTP_SERVER
  
  'Build Body
  sBody = "FindsXP Ship-At-A-Glance Website Error:" & vbCr & vbLf
  sBody = sBody & str_in_ErrorInfo
  
  'set Email Information
'  oMSEmail.AddAddress "Michael.Katzman@pacer.com", " Michael Katzman"
  'if there is no user email address in the system send from automated FINDSxp email address
'  oMSEmail.From = "SAAG@Pacer.com"
'  oMSEmail.FromName = "FINDSxp SAAG"
'  oMSEmail.Subject = "FindsXP Ship-At-A-Glance Website Error"
'  oMSEmail.Body = "" & sBody
'  oMSEmail.Send

  Dim objFileSystem As New FileSystemObject
  Dim DebugFile As TextStream
  
  'Create text stream file
  Set objFileSystem = CreateObject("Scripting.FileSystemObject")
  
  If objFileSystem.FileExists(App.Path & "\ErrorLog.log") = False Then
      Set DebugFile = objFileSystem.CreateTextFile(App.Path & "\ErrorLog.log", False)
  Else
      Set DebugFile = objFileSystem.OpenTextFile(App.Path & "\ErrorLog.log", ForAppending, False, TristateUseDefault)
  End If
  
  DebugFile.WriteLine (sBody)

  DebugFile.Close
  
End Sub

Private Sub Class_Initialize()
  
'  Set oMSEmail = New MailSender
'  oMSEmail.Host = SMTP_SERVER

End Sub


Private Sub Class_Terminate()
  
'  Set oMSEmail = Nothing

End Sub


