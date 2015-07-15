'Imports System
'Imports System.Runtime.InteropServices
'Imports System.Drawing
'Imports System.Drawing.Imaging
'Imports System.Diagnostics

'Public Class AppAdmin
'    Private cEnviro As Enviro
'    '''Private Const cVersionNumber As String = "1.10"
'    '''Private Const cMaxDisplayRecords As Integer = 500
'    '''Private Const cExcessiveRecordAmount As Integer = 1000
'    '''Private Const cRecordMaximum As Integer = 5000

'    '''Public Enum ReportTypeEnum
'    '''    Information = 1
'    '''    InformationNoLog = 2
'    '''    [Error] = 3
'    '''    ErrorNoShutdown = 4
'    '''    Timeout = 5
'    '''End Enum

'    Public Sub New()
'        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
'        cEnviro = SessionObj("Enviro")
'    End Sub

'    ''''Public ReadOnly Property VersionNumber() As String
'    ''''    Get
'    ''''        Return cVersionNumber
'    ''''    End Get
'    ''''End Property
'    ''''Public ReadOnly Property MaxDisplayRecords() As Integer
'    ''''    Get
'    ''''        Return cMaxDisplayRecords
'    ''''    End Get
'    ''''End Property
'    ''''Public ReadOnly Property ExcessiveRecordAmount() As Integer
'    ''''    Get
'    ''''        Return cExcessiveRecordAmount
'    ''''    End Get
'    ''''End Property
'    ''''Public ReadOnly Property RecordMaximum() As Integer
'    ''''    Get
'    ''''        Return cRecordMaximum
'    ''''    End Get
'    ''''End Property

'    'Public Function Report(ByVal ReportType As ReportTypeEnum, ByVal Message As String) As String
'    '    Dim Coll As New Collection
'    '    Dim SendEmailResults As Results
'    '    Dim HeaderMessage As String

'    '    Try

'    '        Select Case ReportType
'    '            Case ReportTypeEnum.Information
'    '                WriteToLogFile(Message)

'    '            Case ReportTypeEnum.InformationNoLog

'    '            Case ReportTypeEnum.Error
'    '                WriteToLogFile(Message & " ** ERROR FORCING APPLICATION SHUTDOWN **")

'    '            Case ReportTypeEnum.ErrorNoShutdown
'    '                WriteToLogFile(Message)

'    '        End Select

'    '        Select Case ReportType
'    '            Case ReportTypeEnum.Error

'    '                ' ___ Send email to help desk, including log and screenshot
'    '                Coll.Add(cEnviro.LogFileFullPath)

'    '                'SendEmailResults = SendEmail("rbluestein@benefitvision.com", cEnviro.LoggedInUserID & "@benefitvision.com", "", "UserManagement error - " & cEnviro.LoggedInUserID & "/" & cEnviro.LoginIP, "An error occurred in the execution of UserManagement. Attached please find a copy of the application log and a screen shot.", Coll)
'    '                SendEmailResults = SendEmail("rbluestein@benefitvision.com", cEnviro.LoggedInUserID & "@benefitvision.com", "", "UserManagement error - " & cEnviro.LoggedInUserID & "/" & cEnviro.LoginIP, "An error occurred in the execution of UserManagement. Attached please find a copy of the application log and a screen shot.", Coll)
'    '                If SendEmailResults.Success Then
'    '                    HeaderMessage = "An error has occurred requiring UserManagement to shut down. You may view the details of the problem in the log. UserManagement has emailed a copy of the log to the help desk."
'    '                Else
'    '                    HeaderMessage = "An error has occurred requiring UserManagement to shut down. You may view the details of the problem in the log."
'    '                End If

'    '        End Select

'    '        Return HeaderMessage

'    '    Catch ex As Exception
'    '        Throw New Exception("Error #802: AppAdmin Report. " & ex.Message)
'    '    End Try
'    'End Function

'    '''''Public Function Report(ByVal ReportType As ReportTypeEnum, ByVal Message As String) As String
'    '''''    Dim Coll As New Collection
'    '''''    Dim SendEmailResults As Results
'    '''''    Dim HeaderMessage As String = String.Empty
'    '''''    Dim WriteToLog As Boolean
'    '''''    Dim SendEmailPlease As Boolean

'    '''''    Try

'    '''''        '___ Write to log
'    '''''        Select Case ReportType
'    '''''            Case ReportTypeEnum.Information, ReportTypeEnum.ErrorNoShutdown
'    '''''                WriteToLogFile(Message)
'    '''''                WriteToLog = True
'    '''''            Case ReportTypeEnum.Error
'    '''''                WriteToLogFile(Message & " ** ERROR FORCING APPLICATION SHUTDOWN **")
'    '''''                WriteToLog = True
'    '''''        End Select

'    '''''        ' ___ Email
'    '''''        Select Case ReportType
'    '''''            Case ReportTypeEnum.Error, ReportTypeEnum.ErrorNoShutdown
'    '''''                SendEmailPlease = True
'    '''''                Coll.Add(cEnviro.LogFileFullPath)

'    '''''                'SendEmailResults = SendEmail("rbluestein@benefitvision.com", cEnviro.LoggedInUserID & "@benefitvision.com", "", "UserManagement error - " & cEnviro.LoggedInUserID & "/" & cEnviro.LoginIP, "An error occurred in the execution of UserManagement. Attached please find a copy of the application log and a screen shot.", Coll)
'    '''''                SendEmailResults = SendEmail("HelpDesk@benefitvision.com", cEnviro.LoggedInUserID & "@benefitvision.com", "rbluestein@benefitvision.com", "UserManagement error - " & cEnviro.LoggedInUserID & "/" & cEnviro.LoginIP, "An error occurred in the execution of UserManagement. Attached please find a copy of the application log.", Coll)

'    '''''        End Select

'    '''''        ' ___ Header message shutdown
'    '''''        Select Case ReportType
'    '''''            Case ReportTypeEnum.Error, ReportTypeEnum.Timeout
'    '''''                HeaderMessage = "An error has occurred requiring UserManagement to shut down. "
'    '''''        End Select

'    '''''        ' ___ Header message log
'    '''''        Select Case ReportType
'    '''''            Case ReportTypeEnum.InformationNoLog, ReportTypeEnum.Timeout
'    '''''                ' no action
'    '''''            Case Else
'    '''''                HeaderMessage &= "You may view the details of the problem in the log. "
'    '''''        End Select

'    '''''        ' ___ Email
'    '''''        If SendEmailPlease AndAlso SendEmailResults.Success Then
'    '''''            HeaderMessage &= "UserManagement has emailed a copy of the log to the help desk."
'    '''''        End If

'    '''''        Return HeaderMessage.Trim

'    '''''    Catch ex As Exception
'    '''''        Throw New Exception("Error #802: AppAdmin Report. " & ex.Message)
'    '''''    End Try
'    '''''End Function

'    ''''Private Function ReadLogFile() As String
'    ''''    Dim StreamReader As System.IO.StreamReader
'    ''''    Dim FileText As String

'    ''''    Try
'    ''''        StreamReader = New System.IO.StreamReader(cEnviro.LogFileFullPath)
'    ''''        FileText = StreamReader.ReadToEnd

'    ''''        'Do While StreamReader.Peek() >= 0
'    ''''        '    'Console.WriteLine(StreamReader.ReadLine())

'    ''''        '    x = StreamReader.ReadLine()
'    ''''        'Loop
'    ''''        StreamReader.Close()
'    ''''        Return FileText

'    ''''    Catch
'    ''''    End Try
'    ''''End Function

'    ''''Private Sub WriteToLogFile(ByVal Message As String)
'    ''''    'Dim FileInfo As System.IO.FileInfo
'    ''''    Dim StreamWriter As System.IO.StreamWriter

'    ''''    Try

'    ''''        Message = Replace(Message, "~", "")

'    ''''        'FileInfo = New System.IO.FileInfo(cEnviro.LogFileFullPath)
'    ''''        StreamWriter = New System.IO.StreamWriter(cEnviro.LogFileFullPath, True)
'    ''''        'StreamWriter.Write("[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & "] " & Message & vbCrLf)
'    ''''        StreamWriter.Write(GetTimeStamp() & Message & vbCrLf)
'    ''''        StreamWriter.Close()
'    ''''    Catch ex As Exception
'    ''''        Try
'    ''''            StreamWriter.Close()
'    ''''        Catch
'    ''''        End Try
'    ''''    End Try
'    ''''End Sub

'    ''''Private Function GetTimeStamp() As String
'    ''''    ' Return "[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & "] "
'    ''''    Return "[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & " " & cEnviro.LoggedInUserID & "] "
'    ''''End Function

'    'Public Shared Sub CaptureScreen(ByVal FileFullPath As String)
'    '    Dim SC As New ScreenCapture
'    '    System.Threading.Thread.Sleep(1000)
'    '    SC.CaptureScreenToFile(FileFullPath, System.Drawing.Imaging.ImageFormat.Gif)
'    '    SC = Nothing
'    'End Sub

'    'Private Function SendEmail(ByVal SendTo As String, ByVal From As String, ByVal cc As String, ByVal Subject As String, ByVal TextBody As String, ByRef AttachmentColl As Collection) As Results
'    '    Dim i As Integer
'    '    Dim MyResults As New Results

'    '    Try

'    '        Dim MailMsg As New System.Web.Mail.MailMessage
'    '        Dim Attachment As MailAttachment
'    '        MailMsg.To = SendTo
'    '        MailMsg.From = From
'    '        MailMsg.Subject = Subject
'    '        MailMsg.BodyFormat = MailFormat.Text
'    '        'MailMsg.BodyFormat = MailFormat.Html
'    '        MailMsg.Body = TextBody

'    '        For i = 1 To AttachmentColl.Count
'    '            Attachment = New MailAttachment(AttachmentColl(i))
'    '            MailMsg.Attachments.Add(Attachment)
'    '        Next

'    '        SmtpMail.Send(MailMsg)

'    '        MyResults.Success = True
'    '        Return MyResults

'    '    Catch ex As Exception
'    '        MyResults.Success = False
'    '        Return MyResults
'    '    End Try
'    'End Function

'    ''''Public Shared Function SendEmail(ByVal SendTo As String, ByVal From As String, ByVal cc As String, ByVal Subject As String, ByVal TextBody As String, Optional ByRef AttachmentColl As Collection = Nothing) As Results
'    ''''    Dim schema As String
'    ''''    Dim MyResults As New Results
'    ''''    Dim i As Integer

'    ''''    Try

'    ''''        Dim CDOConfig As New CDO.Configuration
'    ''''        schema = "http://schemas.microsoft.com/cdo/configuration/"
'    ''''        CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = 2
'    ''''        ' CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value = 25
'    ''''        CDOConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value = "mail.benefitvision.com"
'    ''''        CDOConfig.Fields.Update()

'    ''''        Dim iMsg As New CDO.Message
'    ''''        iMsg.To = SendTo
'    ''''        iMsg.From = From
'    ''''        iMsg.CC = cc
'    ''''        iMsg.Subject = Subject

'    ''''        If Not AttachmentColl Is Nothing Then
'    ''''            For i = 1 To AttachmentColl.Count
'    ''''                iMsg.AddAttachment(AttachmentColl(i))
'    ''''            Next
'    ''''        End If

'    ''''        iMsg.Configuration = CDOConfig
'    ''''        iMsg.TextBody = TextBody
'    ''''        'imsg.HTMLBody = htmlbody

'    ''''        iMsg.Send()

'    ''''        ' ___ Clean up
'    ''''        iMsg.Attachments.DeleteAll()
'    ''''        CDOConfig = Nothing
'    ''''        iMsg = Nothing

'    ''''        MyResults.Success = True
'    ''''        Return MyResults

'    ''''    Catch ex As Exception
'    ''''        MyResults.Success = False
'    ''''        MyResults.Msg = "Error #803: " & ex.Message
'    ''''        Return MyResults
'    ''''    End Try
'    ''''End Function
'End Class

'''''Public Class BulkAppointmentDataPack
'''''    Private cDT As DataTable

'''''    Public Class JITBulkClass
'''''        Public Const JIT = "JIT"
'''''        Public Const ApptExpress = "ApptExpress"
'''''    End Class
'''''    Public Class SaveTypeClass
'''''        Public Const JITTrigger = "JIT trigger"
'''''        Public Const JITQualify = "JIT qualify"
'''''        Public Const ApptExpress = "Appt Express"
'''''        Public Const ApptExpressJITTrigger = "Appt Express JIT trigger"
'''''    End Class

'''''    Public Sub New()
'''''        cDT = New DataTable
'''''        cDT.Columns.Add("RecordKey", GetType(System.String))
'''''        cDT.Columns.Add("JITBulk", GetType(System.String))
'''''        cDT.Columns.Add("SaveType", GetType(System.String))
'''''        cDT.Columns.Add("State", GetType(System.String))
'''''        cDT.Columns.Add("CarrierID", GetType(System.String))
'''''        cDT.Columns.Add("Success", GetType(System.Int64))
'''''        cDT.Columns.Add("Comments", GetType(System.String))
'''''    End Sub

'''''    Public Sub Init()
'''''        cDT.Rows.Clear()
'''''    End Sub

'''''    Public Sub BulkAdd(ByVal SaveType As String, ByVal State As String, ByVal CarrierID As String, ByVal Success As Integer, ByVal Comments As String, Optional ByVal MakeFirstRow As Boolean = False)
'''''        Dim dr As DataRow

'''''        Try

'''''            dr = cDT.NewRow
'''''            dr("RecordKey") = State & "_" & CarrierID

'''''            Select Case SaveType
'''''                Case SaveTypeClass.JITTrigger, SaveTypeClass.JITQualify
'''''                    dr("JITBulk") = JITBulkClass.JIT
'''''                Case SaveTypeClass.ApptExpressJITTrigger, SaveTypeClass.ApptExpress
'''''                    dr("JITBulk") = JITBulkClass.ApptExpress
'''''            End Select

'''''            dr("SaveType") = SaveType
'''''            dr("State") = State
'''''            dr("CarrierID") = CarrierID
'''''            dr("Success") = Success
'''''            dr("Comments") = Comments

'''''            If MakeFirstRow Then
'''''                cDT.Rows.InsertAt(dr, 0)
'''''            Else
'''''                cDT.Rows.Add(dr)
'''''            End If

'''''        Catch ex As Exception
'''''            Throw New Exception("Error #803: BulkAppointmentDataPack BulkAdd. " & ex.Message)
'''''        End Try
'''''    End Sub

'''''    Public Function RecordExists(ByVal State As String, ByVal CarrierID As String) As Boolean
'''''        Dim i As Integer
'''''        Dim Expr As String

'''''        Try

'''''            Expr = State & "_" & CarrierID
'''''            For i = 0 To cDT.Rows.Count - 1
'''''                If cDT.Rows(i)("RecordKey") = Expr Then
'''''                    Return True
'''''                End If
'''''            Next
'''''            Return False

'''''        Catch ex As Exception
'''''            Throw New Exception("Error #804: BulkAppointmentDataPack RecordExists. " & ex.Message)
'''''        End Try
'''''    End Function

'''''    Public ReadOnly Property DT() As DataTable
'''''        Get
'''''            Return cDT
'''''        End Get
'''''    End Property
'''''End Class

'Public Class BrowserNav
'    Public IndexIniatallyLoaded As Boolean
'    Public UserMaintainInitiallyLoaded As Boolean
'End Class