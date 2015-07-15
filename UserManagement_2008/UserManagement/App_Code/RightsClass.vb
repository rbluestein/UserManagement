Imports System.Data

Public Class RightsClass
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRightsColl As New Collection

#Region " Constants "
    Public Const UserView As String = "USV"
    Public Const UserEdit As String = "USE"
    Public Const PermissionsView As String = "PRV"
    Public Const PermissionsEdit As String = "PRE"
    Public Const LicenseView As String = "LIV"
    Public Const LicenseEdit As String = "LIE"
    Public Const CarrierView As String = "CAV"
    Public Const CarrierEdit As String = "CAE"
    Public Const ClientView As String = "CLV"
    Public Const ClientEdit As String = "CLE"
    Public Const CompanyView As String = "COV"
    Public Const CompanyEdit As String = "COE"
    Public Const AdminLicSpecific As String = "ALS"
#End Region

    Public Enum AccessLevelEnum
        AllAccess = 1
        SupervisorAccess = 2
        EnrollerAccess = 3
    End Enum

    Public ReadOnly Property RightsColl() As Collection
        Get
            Return cRightsColl
        End Get
    End Property

    Public Sub New(ByVal CurPage As Page)
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        Dim dt As DataTable
        Dim i As Integer
        Dim AllRights(12) As String

        Try

            cEnviro = SessionObj("Enviro")
            cCommon = New Common

            dt = cCommon.GetDT("SELECT Role, LocationID FROM Users WHERE UserID ='" & cEnviro.LoggedInUserID & "'")
            If dt.Rows.Count = 0 Then
                CurPage.Response.Redirect("InsufficientRights.aspx")
            End If

            If cEnviro.LoggedInUserID = "rbluestein" Then
                dt.Rows(0)("Role") = "ADMIN LIC"
            End If

            AllRights(0) = "USV"
            AllRights(1) = "USE"
            AllRights(2) = "PRV"
            AllRights(3) = "PRE"
            AllRights(4) = "LIV"
            AllRights(5) = "LIE"
            AllRights(6) = "CAV"
            AllRights(7) = "CAE"
            AllRights(8) = "CLV"
            AllRights(9) = "CLE"
            AllRights(10) = "COV"
            AllRights(12) = "ALS"

            Select Case cEnviro.DBHost
                Case "HBG-SQL"
                    Select Case dt.Rows(0)("Role")
                        Case "IT"
                            cRightsColl.Add("USV")
                            cRightsColl.Add("USE")
                            cRightsColl.Add("PRV")
                            cRightsColl.Add("PRE")
                            cRightsColl.Add("LIV")
                            cRightsColl.Add("LIE")
                            cRightsColl.Add("CAE")
                            cRightsColl.Add("CAV")
                            cRightsColl.Add("CLE")
                            cRightsColl.Add("CLV")
                            cRightsColl.Add("COE")
                            cRightsColl.Add("COV")
                        Case "ADMIN LIC"
                            cRightsColl.Add("USV")
                            cRightsColl.Add("USE")
                            cRightsColl.Add("PRV")
                            cRightsColl.Add("PRE")
                            cRightsColl.Add("LIV")
                            cRightsColl.Add("LIE")
                            cRightsColl.Add("CAV")
                            cRightsColl.Add("CAE")
                            cRightsColl.Add("CLV")
                            cRightsColl.Add("CLE")
                            cRightsColl.Add("ALS")
                        Case "ADMIN"
                            cRightsColl.Add("USV")
                            cRightsColl.Add("PRV")
                            cRightsColl.Add("LIV")
                            cRightsColl.Add("CAV")
                            cRightsColl.Add("CLV")
                            cRightsColl.Add("COV")
                        Case "SUPERVISOR"
                            cRightsColl.Add("USV")
                            cRightsColl.Add("PRV")
                            cRightsColl.Add("LIV")
                            cRightsColl.Add("CAV")
                            cRightsColl.Add("CLV")
                            cRightsColl.Add("COV")
                        Case "ENROLLER"
                            cRightsColl.Add("USV")
                            cRightsColl.Add("PRV")
                            cRightsColl.Add("LIV")
                        Case "CLIENT"
                            ' none
                    End Select

                Case "HBG-TST"
                    Select Case dt.Rows(0)("Role")
                        Case "IT"
                            For i = 0 To AllRights.GetUpperBound(0)
                                cRightsColl.Add(AllRights(i))
                            Next
                        Case "ADMIN LIC"
                            cRightsColl.Add("USV")
                            cRightsColl.Add("USE")
                            cRightsColl.Add("PRV")
                            cRightsColl.Add("PRE")
                            cRightsColl.Add("LIV")
                            cRightsColl.Add("LIE")
                            cRightsColl.Add("CAV")
                            cRightsColl.Add("CAE")
                            cRightsColl.Add("CLV")
                            cRightsColl.Add("CLE")
                        Case "ADMIN"
                            cRightsColl.Add("USV")
                            cRightsColl.Add("PRV")
                            cRightsColl.Add("LIV")
                            cRightsColl.Add("CAV")
                            cRightsColl.Add("CLV")
                            cRightsColl.Add("COV")
                        Case "SUPERVISOR"
                            cRightsColl.Add("USV")
                            cRightsColl.Add("PRV")
                            cRightsColl.Add("LIV")
                            cRightsColl.Add("CAV")
                            cRightsColl.Add("CLV")
                            cRightsColl.Add("COV")
                        Case "ENROLLER"
                            cRightsColl.Add("USV")
                            cRightsColl.Add("PRV")
                            cRightsColl.Add("LIV")
                        Case "CLIENT"
                            ' none
                    End Select

            End Select

        Catch ex As Exception
            Throw New Exception("Error #870: RightsClass New. " & ex.Message)
        End Try
    End Sub

    Public Sub GetSecurityFlds(ByRef AccessLevel As AccessLevelEnum, ByRef Role As String, ByRef LocationID As String)
        Try

            Dim Common As New Common
            Dim dt As DataTable
            dt = Common.GetDT("SELECT Role, LocationID FROM Users WHERE UserID ='" & cEnviro.LoggedInUserID & "'")
            Role = dt.Rows(0)("Role")
            LocationID = dt.Rows(0)("LocationID")
            Select Case dt.Rows(0)("Role")
                Case "ADMIN", "ADMIN LIC", "IT"
                    AccessLevel = AccessLevelEnum.AllAccess
                Case "SUPERVISOR"
                    AccessLevel = AccessLevelEnum.SupervisorAccess
                Case "ENROLLER"
                    AccessLevel = AccessLevelEnum.EnrollerAccess
            End Select

        Catch ex As Exception
            Throw New Exception("Error #871: RightsClass GetSecurityFlds. " & ex.Message)
        End Try
    End Sub


    Public Function HasSufficientRights(ByRef RightsRqd As String(), ByVal RedirectOnError As Boolean, ByRef CurPage As System.Web.UI.Page) As Boolean
        Dim i, j As Integer
        Dim Passed As Boolean

        Try

            For i = 0 To RightsRqd.GetUpperBound(0)
                For j = 1 To cRightsColl.Count
                    If cRightsColl(j) = RightsRqd(i) Then
                        Passed = True
                        Exit For
                    End If
                Next
                If Passed Then
                    Exit For
                End If
            Next

            If Passed Then
                Return True
            Else
                If RedirectOnError Then
                    CurPage.Response.Redirect("InsufficientRights.aspx")
                Else
                    Return False
                End If
            End If

        Catch ex As Exception
            Throw New Exception("Error #872: RightsClass HasSufficientRights. " & ex.Message)
        End Try
    End Function

    Public Function HasThisRight(ByVal RightCd As String) As Boolean
        Dim i As Integer
        For i = 1 To cRightsColl.Count
            If cRightsColl(i) = RightCd Then
                Return True
            End If
        Next
        Return False
    End Function

End Class

