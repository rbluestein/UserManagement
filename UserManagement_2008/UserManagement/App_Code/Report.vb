Imports System.Net.Mail

Public Class Report
    Private cEnviro As Enviro
    Private cCommon As Common

    Public Sub New()
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        cEnviro = SessionObj("Enviro")
        cCommon = New Common
    End Sub

    Public Function Report(ByVal ErrorMessage As String, ByVal WriteToLog As Boolean, ByVal SendEmailPlease As Boolean, ByVal Shutdown As Boolean) As String
        Dim Coll As New Collection
        Dim SendEmailResults As Results
        Dim HeaderMessage As String = String.Empty
        Dim EmailSuccess As Boolean
        Dim Subj As String

        Try

            ' ___ Do not send emails or make log entries for errors occurring on development machine.
            If System.Net.Dns.GetHostName().ToUpper = "WADEV" Then
                SendEmailPlease = False
                WriteToLog = False
            End If

            ' ___ Write to log
            If WriteToLog Then
                If Shutdown Then
                    WriteToLogFile(ErrorMessage & " ** ERROR FORCING APPLICATION SHUTDOWN **")
                Else
                    WriteToLogFile(ErrorMessage)
                End If
            End If

            ' ___ Send email
            If SendEmailPlease Then
                Coll.Add(cEnviro.LogFileFullPath)
                Subj = "User Management error - User: " & IIf(cEnviro.LoggedInUserID = Nothing, cEnviro.LoggedInUserID, cEnviro.LoggedInUserID) & " - Time: " & cCommon.GetServerDateTime()
                SendEmailResults = SendEmail("HelpDesk@benefitvision.com", "automail@benefitvision.com", Nothing, Subj, ErrorMessage)
                EmailSuccess = SendEmailResults.Success
            End If

            If Shutdown And WriteToLog And EmailSuccess Then
                HeaderMessage = "An error has occurred requiring User Management to shut down. User Management has emailed a notice to the help desk. You may view the details of the problem in the log."
            ElseIf Shutdown And WriteToLog And (Not EmailSuccess) Then
                HeaderMessage = "An error has occurred requiring User Management to shut down. You may view the details of the problem in the log."
            ElseIf Shutdown And (Not WriteToLog) And EmailSuccess Then
                HeaderMessage = "An error has occurred requiring User Management to shut down. User Management has emailed a notice to the help desk."
            ElseIf Shutdown And (Not WriteToLog) And (Not EmailSuccess) Then
                HeaderMessage = "An error has occurred requiring User Management to shut down. "
            End If

            Return HeaderMessage.Trim

        Catch ex As Exception
            Throw New Exception("Error #1502: Report Report. " & ex.Message, ex)
        End Try
    End Function

    Public Shared Function SendEmail(ByVal SendTo As String, ByVal From As String, ByVal cc As String, ByVal Subject As String, ByVal TextBody As String, Optional ByRef AttachmentColl As Collection = Nothing) As Results
        Dim i As Integer
        Dim MyResults As New Results
        Dim Msg As System.Net.Mail.MailMessage
        Dim Client As System.Net.Mail.SmtpClient
        Dim SendToBox() As String

        Try

            'MMsg = New System.Net.Mail.MailMessage(From, SendTo, Subject, TextBody)

            'If ConfigurationManager.AppSettings("DBHost").ToString.ToLower = "hbg-tst" Then
            '    If ConfigurationManager.AppSettings("EmailSupvSubstituteTest") <> Nothing Then
            '        SendTo = Replace(SendTo, "HBG_SUPERVISORS@benefitvision.com", ConfigurationManager.AppSettings("EmailSupvSubstituteTest"))
            '        If cc <> Nothing Then
            '            cc = Replace(cc, "HBG_SUPERVISORS@benefitvision.com", ConfigurationManager.AppSettings("EmailSupvSubstituteTest"))
            '        End If
            '    End If
            'End If


            Msg = New MailMessage
            Msg.From = New MailAddress(From)

            SendToBox = Split(SendTo, ";")
            For i = 0 To SendToBox.GetUpperBound(0)
                Msg.To.Add(New MailAddress(SendToBox(i).Trim))
            Next

            ' Msg.To.Add(New MailAddress(SendTo))

            'Msg.To.Add(New MailAddress("rbluestein@benefitvision.com"))

            If cc <> Nothing Then
                Msg.CC.Add(cc)
            End If

            Msg.Subject = Subject
            Msg.Body = TextBody
            Msg.IsBodyHtml = False

            '___ Attachments
            If AttachmentColl IsNot Nothing Then
                For i = 1 To AttachmentColl.Count
                    Msg.Attachments.Add(New System.Net.Mail.Attachment(AttachmentColl(i)))
                Next
            End If

            Client = New System.Net.Mail.SmtpClient("mail.benefitvision.com")
            Client.Send(Msg)

            MyResults.Success = True
            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Msg = "Error #1503: " & ex.Message
            Return MyResults
        End Try
    End Function

    Private Function ReadLogFile() As String
        Dim StreamReader As System.IO.StreamReader = Nothing
        Dim FileText As String

        Try
            StreamReader = New System.IO.StreamReader(cEnviro.LogFileFullPath)
            FileText = StreamReader.ReadToEnd

            'Do While StreamReader.Peek() >= 0
            '    'Console.WriteLine(StreamReader.ReadLine())

            '    x = StreamReader.ReadLine()
            'Loop
            'StreamReader.Close()
            Return FileText

        Catch ex As Exception
            Throw New Exception("Error #1504: Report ReadLogFile. " & ex.Message, ex)
        Finally
            Try
                StreamReader.Close()
            Catch
            End Try
        End Try
    End Function

    Public Sub WriteToLogFile(ByVal Message As String)
        Dim i As Integer
        'Dim FileInfo As System.IO.FileInfo
        Dim StreamWriter As System.IO.StreamWriter = Nothing

        Try

            Message = Replace(Message, "~", "")

            'FileInfo = New System.IO.FileInfo(cEnviro.LogFileFullPath)

            Try
                StreamWriter = New System.IO.StreamWriter(cEnviro.LogFileFullPath, True)
            Catch
                Dim procList() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                For i = 0 To procList.GetUpperBound(0)
                    If procList(i).ProcessName = "notepad" Then
                        procList(i).Kill()
                    End If
                    'System.Diagnostics.Debug.WriteLine(procList(k).ProcessName)
                Next
                StreamWriter = New System.IO.StreamWriter(cEnviro.LogFileFullPath, True)
            End Try



            'StreamWriter.Write("[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & "] " & Message & vbCrLf)
            StreamWriter.Write(GetTimeStamp() & Message & vbCrLf)
            ' StreamWriter.Close()
        Catch ex As Exception
            'Throw New Exception("Error #1505: Report WriteToLogFile. " & ex.Message, ex)
        Finally
            Try
                StreamWriter.Close()
            Catch
            End Try
        End Try
    End Sub

    Private Function GetTimeStamp() As String
        Return "[" & Date.Now.ToUniversalTime.AddHours(-5).ToString & " " & cEnviro.LoggedInUserID & "] "
    End Function
End Class
