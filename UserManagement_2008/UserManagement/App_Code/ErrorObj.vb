Public Class ErrorObj

    ' //////////////////////////////////////////////////
    ' Why the try try you ask.
    ' An exception includes a message property. In
    ' order to include the error message with a number and 
    ' method name, the next higher exception in the hierarchy 
    ' prefixes this information.
    '
    ' If an error occurs in the top-most exception level
    ' in the hierarchy, the only way of including both
    ' the exception message and the prefix message is
    ' by throwing an additional exception for the 
    ' prefix message.
    '
    ' Microsoft bug alert
    ' Error handling uses a session object, called
    ' ErrorArgs. If the error handling is written such
    ' that Global.Application_Error is used at all,
    ' it causes the loss of the session objects. This
    ' application processes errors in the ErrorObj 
    ' as an alternative.
    ' //////////////////////////////////////////////////

#Region " Declarations "
    ' Private cErrorType As ErrorTypeEnum
    ' Private cErrorMessage As String
#End Region

#Region " Enums "
    'Public Enum ErrorTypeEnum
    '    [Error] = 1
    '    Cookie = 2
    '    Timeout = 3
    '    Connection = 4
    '    ProductRegistration = 5
    'End Enum
#End Region

#Region " Properties "
    'Public ReadOnly Property ErrorType() As ErrorTypeEnum
    '    Get
    '        Return cErrorType
    '    End Get
    'End Property
    'Public ReadOnly Property ErrorMessage() As String
    '    Get
    '        Return cErrorMessage
    '    End Get
    'End Property
#End Region

#Region " Methods "
    Public Sub New(ByRef ex As Exception, ByVal ErrorTopLevel As String)
        Dim CurException As Exception
        Dim Coll As New CollX
        Dim ErrorMessage As String
        Dim Refined As String()

        Try

            '' ___ Extract the error
            'If Exception.InnerException Is Nothing Then
            '    ExceptionMessage = Exception.Message
            'Else
            '    CurException = Exception
            '    While Not (CurException Is Nothing)
            '        Coll.Add(CurException.Message)
            '        CurException = CurException.InnerException
            '    End While
            '    ExceptionMessage = Coll(Coll.Count - 1)
            'End If

            'If InStr(ExceptionMessage, "Thread was being aborted.") > 0 Then
            '    Exit Sub
            'End If

            ' ___ Extract the error




            If ex.InnerException Is Nothing Then
                ErrorMessage = ex.Message
            Else
                CurException = ex
                While Not (ex Is Nothing)
                    Coll.Assign(ex.Message)
                    CurException = ex.InnerException
                End While
                ErrorMessage = Coll(Coll.Count)
            End If

            Refined = Split(ErrorMessage, "Error #")
            ErrorMessage = Refined(Refined.GetUpperBound(0))

            If InStr(ErrorMessage, "Thread was being aborted.") > 0 Then
                Exit Sub
            Else
                HandleError("Error #" & ErrorMessage, True)
            End If

        Catch
        End Try
    End Sub

    Public Sub HandleError(ByVal RawMessage As String, ByVal Shutdown As Boolean)
        Dim Enviro As Enviro
        Dim HeaderMessage As String
        Dim ErrorMessage As String
        Dim WriteToLog As Boolean
        Dim SendEmail As Boolean
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        Dim Report As New Report

        Try

            ' ___ Get the response object
            Dim Response As System.Web.HttpResponse
            Response = HttpContext.Current.Response

            ' ___ Get Enviro from session
            Enviro = SessionObj("Enviro")

            ' ___ Get the ErrorType and ErrorMessage
            If InStr(RawMessage, "UnableToConnect", CompareMethod.Text) > 0 Then
                ErrorMessage = "User Management is unable to establish a connection with " & Enviro.DBHost & ". Please report this problem to your supervisor or IT support person."
                WriteToLog = True
                SendEmail = False
                Shutdown = True
            ElseIf InStr(RawMessage, "timeout1", CompareMethod.Text) > 0 Then
                ErrorMessage = "User Management has either timed out or you have tried to launch from an incorrect page. Please close the application and attempt to launch again."
                WriteToLog = False
                SendEmail = False
                Shutdown = True
            ElseIf InStr(RawMessage, "timeout2", CompareMethod.Text) > 0 Then
                ErrorMessage = "User Management has timed out. Please close the application and attempt to launch again."
                WriteToLog = False
                SendEmail = False
                Shutdown = True
            Else
                ErrorMessage = RawMessage
                WriteToLog = True
                SendEmail = True
                Shutdown = True
            End If

            ' ___ Send the email
            HeaderMessage = Report.Report(ErrorMessage, WriteToLog, SendEmail, Shutdown)

            ' ___ Handle possible redirect to error page and application shutdown.
            If Shutdown Then
                ErrorMessage = Replace(ErrorMessage, "#", "[sharp]")
                ErrorMessage = Replace(ErrorMessage, vbCrLf, "~")
                ErrorMessage = Replace(ErrorMessage, Chr(10), "~")
                If ErrorMessage.Length + HeaderMessage.Length > 2000 Then
                    ErrorMessage = ErrorMessage.Substring(0, (2000 - HeaderMessage.Length))
                End If

                Response.Redirect("ErrorPage.aspx?ErrorMessage=" & ErrorMessage & "&HeaderMessage=" & HeaderMessage)
            End If

        Catch ex As Exception
            'Throw New Exception("Error #2303: ErrorObj HandleError. " & ex.Message, ex)
        End Try
    End Sub

    'Private Function GetLink() As String
    '    Dim Results As String = Nothing
    '    Select Case ConfigurationManager.AppSettings("DBHost").ToString.ToLower
    '        Case "hbg-sql"
    '            Results = "http://netserver.benefitvision.com/Callback/"
    '        Case "hbg-tst"
    '            Results = "http://test.benefitvision.com/Callback/"
    '        Case "training"
    '            Results = "http://train.benefitvision.com/Callback/"
    '    End Select
    '    Return Results
    'End Function
#End Region
End Class

