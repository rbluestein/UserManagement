Public Class Enviro
#Region " Constants "
    'Private Const cVersionNumber As String = "1.37"
    Private Const cVersionNumber As String = "1.38" ' Added encryption of connectionstring.
    'Private Const cMaxDisplayRecords As Integer = 500
    Private Const cExcessiveRecordAmount As Integer = 1000
    Private Const cRecordMaximum As Integer = 5000
    Private Const cDBTimeout As Integer = 10
    Private Const cAppTimeout As Integer = 1800   ' 30 minutes
    Private Const cDefaultDatabase As String = "UserManagement"
#End Region

#Region " Declarations "
    Private cInit As Boolean
    Private cLastPageLoad As DateTime
    Private cLoggedInUserID As String
    Private cApplicationPath As String
    Private cLoginIP As String
    Private cLogFileFullPath As String
    Private cQueryDownloadDir As String
    Private cDBHost As String
    Private cDBConnStringTemplate As String = Simple3Des.DecryptData("buffalo", "r1sUZTMYRyXq21OKSqReHEoLQB/nu6t20dYmeJ0Xc1Op0kk1vLxics08F9Inm9fMEDH8maoYGUlqiI+Lp5H0wtQwI/QGTmOgVUOKxIqhlV0st2LJckdYQ4NAdWUvAGv8O0gPI6Zm1mIYkvH1mKpbAMdug+3ErknMsQQbtdaUJGkc/eaD7kMJPQ==")
    'Private cAppTimedOut As Boolean
    Private cSessionID As String
#End Region

#Region " Properties "
    Public WriteOnly Property SessionID() As String
        Set(ByVal Value As String)
            cSessionID = Value
        End Set
    End Property
    Public ReadOnly Property VersionNumber() As String
        Get
            Return cVersionNumber
        End Get
    End Property
    'Public ReadOnly Property MaxDisplayRecords() As Integer
    '    Get
    '        Return cMaxDisplayRecords
    '    End Get
    'End Property
    Public ReadOnly Property ExcessiveRecordAmount() As Integer
        Get
            Return cExcessiveRecordAmount
        End Get
    End Property
    Public ReadOnly Property RecordMaximum() As Integer
        Get
            Return cRecordMaximum
        End Get
    End Property

    'Public ReadOnly Property AppTimeout() As Integer
    '    Get
    '        Return cAppTimeout
    '    End Get
    'End Property

    'Public Property AppTimedOut() As Boolean
    '    Get
    '        Return cAppTimedOut
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        cAppTimedOut = Value
    '    End Set
    'End Property

    Public ReadOnly Property DefaultDatabase() As String
        Get
            Return cDefaultDatabase
        End Get
    End Property

    Public Property DBHost() As String
        Get
            Return cDBHost
        End Get
        Set(ByVal Value As String)
            cDBHost = Value
        End Set
    End Property
    Public ReadOnly Property DBConnStringTemplate() As String
        Get
            Return cDBConnStringTemplate
        End Get
    End Property

    Public Property Init() As Boolean
        Get
            Return cInit
        End Get
        Set(ByVal Value As Boolean)
            cInit = Value
        End Set
    End Property
    Public Property LastPageLoad() As DateTime
        Get
            Return cLastPageLoad
        End Get
        Set(ByVal Value As DateTime)
            cLastPageLoad = Value
        End Set
    End Property
    Public Property LoggedInUserID() As String
        Get
            Return cLoggedInUserID
        End Get
        Set(ByVal Value As String)
            cLoggedInUserID = Value
        End Set
    End Property

    Public Property LoginIP() As String
        Get
            Return cLoginIP
        End Get
        Set(ByVal Value As String)
            cLoginIP = Value
        End Set
    End Property

    Public Property ApplicationPath() As String
        Get
            Return cApplicationPath
        End Get
        Set(ByVal Value As String)
            cApplicationPath = Value
        End Set
    End Property

    Public Property LogFileFullPath() As String
        Get
            Return cLogFileFullPath
        End Get
        Set(ByVal Value As String)
            cLogFileFullPath = Value
        End Set
    End Property

    Public Property QueryDownloadDir() As String
        Get
            Return cQueryDownloadDir
        End Get
        Set(ByVal Value As String)
            cQueryDownloadDir = Value
        End Set
    End Property
#End Region

#Region " Methods "
    Public Function GetConnectionString(ByVal DBHost As String, ByVal Database As String) As String
        Return Replace(cDBConnStringTemplate, "|", Database) & DBHost
    End Function

    Public Function GetConnectionString() As String
        Return Replace(cDBConnStringTemplate, "|", cDefaultDatabase) & cDBHost
    End Function

    'Public Sub AuthenticateRequest_Timeout(ByRef Page As System.Web.UI.Page)
    '    Try

    '        ' ___ Check for time out
    '        ValidateAppTimeout()

    '        ' ___ Validate cookie
    '        If IsNothing(Page.Request.Cookies.Item("BVIUM")) Then
    '            Throw New Exception("No Cookie Present.")
    '        Else
    '            If Page.Request.Cookies.Item("BVIUM").Value <> cLoggedInUserID Then
    '                Throw New Exception("No Cookie Present.")
    '            End If
    '        End If

    '    Catch ex As Exception
    '        Throw New Exception("Error #820: Enviro AuthenticateRequest_Timeout. " & ex.Message)
    '    End Try
    'End Sub
    'Public Sub ValidateAppTimeout()
    '    Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
    '    Dim Enviro As Enviro
    '    Dim tp As TimeSpan

    '    Try

    '        If cAppTimedOut Then
    '            Throw New Exception("posttimeout")
    '        End If

    '        Enviro = SessionObj("Enviro")
    '        tp = Date.Now.Subtract(Enviro.LastPageLoad)

    '        If tp.TotalSeconds > cAppTimeout Then
    '            Throw New Exception("Application Timeout.")
    '        Else
    '            Enviro.LastPageLoad = Date.Now
    '        End If

    '    Catch ex As Exception
    '        Throw New Exception("Error #821: Enviro ValidateAppTimeout. " & ex.Message)
    '    End Try
    'End Sub

    'Public Sub MakeCookie(ByVal Page As Page)
    '    Dim SessionCookie As HttpCookie
    '    SessionCookie = New HttpCookie("BVIUM", cLoggedInUserID)
    '    Page.Response.Cookies.Add(SessionCookie)
    'End Sub

    Public Sub MakeCookie(ByVal Page As Page)
        Dim SessionID As String
        Dim SessionCookie As HttpCookie
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session

        SessionID = cSessionID
        SessionCookie = New HttpCookie(SessionID)
        Page.Response.Cookies.Add(SessionCookie)
    End Sub

    Public Sub AuthenticateRequest(ByRef Page As System.Web.UI.Page)
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        Dim tp As TimeSpan

        Try

            ' ___ Page entry
            If cSessionID = Nothing Then
                Throw New Exception("You have attempted to launch UserManagement from an incorrect page. Please close the application and attempt to launch using this link: http://netserver.benefitvision.com/UserManagement/")
            End If

            ' ___ Cookie
            If IsNothing(Page.Request.Cookies.Item(cSessionID)) Then
                Throw New Exception("No Cookie Present.")
            End If

            ' ___ Timeout
            tp = Date.Now.Subtract(cLastPageLoad)
            If tp.TotalSeconds > cAppTimeout Then
                Throw New Exception("Application Timeout.")
            End If


            ' ___ If we made it this far, we're good!
            cLastPageLoad = Date.Now

        Catch ex As Exception
            Throw New Exception("Error #820: Enviro AuthenticateRequest. " & ex.Message)
        End Try
    End Sub
#End Region
End Class
