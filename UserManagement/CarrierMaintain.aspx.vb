Imports System.Data.SqlClient

Public Class CarrierMaintain
    Inherits System.Web.UI.Page

#Region " Declarations "
    Protected PageCaption As HtmlGenericControl
    Protected WithEvents litMsg As System.Web.UI.WebControls.Literal
    Protected WithEvents litHeading As System.Web.UI.WebControls.Literal
    Protected WithEvents litResponseAction As System.Web.UI.WebControls.Literal
    Protected WithEvents litUpdate As System.Web.UI.WebControls.Literal
    Protected WithEvents txtCarrierID As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblCurrentRights As System.Web.UI.WebControls.Label
    Protected WithEvents txtDescription As System.Web.UI.WebControls.TextBox
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Protected WithEvents litEnviro As System.Web.UI.WebControls.Literal
    Private cSess As CarrierWorklistSession
#End Region

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Results As Results
        Dim RequestAction As RequestAction
        Dim ResponseAction As ResponseAction

        Try

            ' ___ Instantiate objects
            cEnviro = Session("Enviro")
            cCommon = New Common

            ' ___ Restore  session
            cEnviro.AuthenticateRequest(Me)

            ' ___ Right Check
            Dim RightsRqd(0) As String
            cRights = New RightsClass(Page)
            RightsRqd.SetValue(cRights.CarrierView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the session object
            cSess = Session("CarrierWorklistSession")

            ' ___ Get RequestAction
            RequestAction = cCommon.GetRequestAction(Page)

            ' ___ Execute the RequestAction
            Results = ExecuteRequestAction(RequestAction)
            ResponseAction = Results.ResponseAction

            ' ___ Execute the ResponseAction
            If ResponseAction = ResponseAction.ReturnToCallingPage Then
                cSess.PageReturnOnLoadMessage = Results.Msg
                Response.Redirect("CarrierWorklist.aspx?CalledBy=Child")

            Else

                ' ___ Execute the ResponseAction
                DisplayPage(ResponseAction)

                ' ___ Build message if present
                If Not Results.Msg = Nothing Then
                    If Results.ObtainConfirm Then
                        litMsg.Text = "<script language='javascript'>" & Results.Msg & "</script>"
                    Else
                        litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"
                    End If
                End If

                litResponseAction.Text = "<input type='hidden' name='hdResponseAction' value = '" & ResponseAction.ToString & "'>"
            End If

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'>"

        Catch ex As Exception
            Throw New Exception("Error #250: CarrierMaintain Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Function ExecuteRequestAction(ByVal RequestAction As RequestAction) As Results
        Dim ValidationResults As New Results
        Dim SaveResults As New Results
        Dim MyResults As New Results

        Try

            Select Case RequestAction
                Case RequestAction.ReturnToParent
                    MyResults.ResponseAction = ResponseAction.ReturnToCallingPage
                    MyResults.Success = True

                Case RequestAction.CreateNew
                    MyResults.ResponseAction = ResponseAction.DisplayBlank

                Case RequestAction.SaveNew
                    ValidationResults = IsDataValid(RequestAction)
                    If ValidationResults.Success Then
                        SaveResults = PerformSave(RequestAction)
                        If SaveResults.Success Then
                            MyResults.ResponseAction = ResponseAction.ReturnToCallingPage
                            MyResults.Msg = SaveResults.Msg
                        Else
                            MyResults.ResponseAction = ResponseAction.DisplayUserInputNew
                            MyResults.Msg = SaveResults.Msg
                        End If
                    ElseIf ValidationResults.ObtainConfirm Then
                        MyResults.ResponseAction = ResponseAction.DisplayUserInputNew
                        MyResults.ObtainConfirm = True
                        MyResults.Msg = ValidationResults.Msg
                    Else
                        MyResults.ResponseAction = ResponseAction.DisplayUserInputNew
                        MyResults.Msg = ValidationResults.Msg
                    End If

                Case RequestAction.LoadExisting
                    MyResults.ResponseAction = ResponseAction.DisplayExisting

                Case RequestAction.SaveExisting
                    ValidationResults = IsDataValid(RequestAction)
                    If ValidationResults.Success Then
                        SaveResults = PerformSave(RequestAction)
                        If SaveResults.Success Then
                            MyResults.ResponseAction = ResponseAction.ReturnToCallingPage
                            MyResults.Msg = SaveResults.Msg
                        Else
                            MyResults.ResponseAction = ResponseAction.DisplayUserInputExisting
                            MyResults.Msg = SaveResults.Msg
                        End If
                    ElseIf ValidationResults.ObtainConfirm Then
                        MyResults.ResponseAction = ResponseAction.DisplayUserInputExisting
                        MyResults.ObtainConfirm = True
                        MyResults.Msg = ValidationResults.Msg
                    Else
                        'MyResults.ResponseAction = ResponseAction.DisplayExisting
                        MyResults.ResponseAction = ResponseAction.DisplayUserInputExisting
                        MyResults.Msg = ValidationResults.Msg
                    End If

                Case RequestAction.NoSaveNew
                    MyResults.ResponseAction = ResponseAction.DisplayUserInputNew

                Case RequestAction.NoSaveExisting
                    MyResults.ResponseAction = ResponseAction.DisplayUserInputExisting

                Case RequestAction.Other
                    MyResults.ResponseAction = cCommon.GetResponseActionFromRequestActionOther(Page)
            End Select

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #251: CarrierMaintain ExecuteRequestAction. " & ex.Message)
        End Try
    End Function

    Private Function IsDataValid(ByVal RequestAction As RequestAction) As Object
        Dim Results As New Results
        Dim ErrColl As New Collection
        Dim dt As DataTable

        Try

            ' ___ Trim the textbox input
            txtDescription.Text = Trim(txtDescription.Text)

            cCommon.ValidateStringField(ErrColl, txtCarrierID.Text, 1, "carrierid not provided")
            cCommon.ValidateStringField(ErrColl, txtDescription.Text, 1, "description not provided")
            cCommon.ValidateApostrophe(ErrColl, txtCarrierID.Text, "carrierid apostrophe not permitted")
            cCommon.ValidateApostrophe(ErrColl, txtDescription.Text, "description apostrophe not permitted")

            ' ___ For a new carrier is the carrier id already in use?
            If RequestAction = RequestAction.SaveNew Then
                '  dt = cCommon.GetDT("Select Count (*) FROM Codes_CarrierID WHERE lower(CarrierID) = '" & txtCarrierID.Text.ToLower & "'")
                dt = cCommon.GetDT("Select Count (*) FROM Codes_CarrierID WHERE lower(CarrierID) = " & cCommon.StrOutHandler(txtCarrierID.Text.ToLower, False, StringTreatEnum.SideQts_SecApost))
                If dt.Rows(0)(0) > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "carrier id is already  in use")
                End If
            End If

            ' ___ For an existing record are we changing to an existing carrier id?
            If RequestAction = RequestAction.SaveExisting AndAlso txtCarrierID.Text.ToLower <> cSess.CarrierID.ToLower Then
                'dt = cCommon.GetDT("Select Count (*) FROM Codes_CarrierID WHERE lower(CarrierID) = '" & txtCarrierID.Text.ToLower & "'")
                dt = cCommon.GetDT("Select Count (*) FROM Codes_CarrierID WHERE lower(CarrierID) =  " & cCommon.StrOutHandler(txtCarrierID.Text.ToLower, False, StringTreatEnum.SideQts_SecApost))
                If dt.Rows(0)(0) > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "carrier id is already  in use")
                End If
            End If

            If ErrColl.Count = 0 Then
                Results.Success = True
            Else
                Results.Success = False
                Results.Msg = cCommon.ErrCollToString(ErrColl, "Not saved. Please correct the following:")
            End If

            ' ___ For an existing record, if we change the carrier id, are there user records that would be orphaned?
            If Results.Success Then
                If RequestAction = RequestAction.SaveExisting AndAlso txtCarrierID.Text <> cSess.CarrierID Then
                    Dim dtUserCount As DataTable
                    ' dtUserCount = cCommon.GetDT("SELECT Count (*) FROM UserAppointments WHERE CarrierID = '" & cSess.CarrierID & "'")
                    dtUserCount = cCommon.GetDT("SELECT Count (*) FROM UserAppointments WHERE CarrierID = " & cCommon.StrOutHandler(cSess.CarrierID, False, StringTreatEnum.SideQts_SecApost))
                    If dtUserCount.Rows(0)(0) > 0 Then
                        Results.Success = False
                        'Results.ObtainConfirm = True
                        'Results.Msg = "GetUserConfirmation();"
                        Results.Msg = "There are users currently associated with this carrier. CarrierID cannot be changed."
                    End If
                End If
            End If

            Return Results

        Catch ex As Exception
            Throw New Exception("Error #252: CarrierMaintain IsDataValid. " & ex.Message)
        End Try
    End Function

    Private Function PerformSave(ByVal RequestAction As RequestAction) As Results
        Dim Sql As New System.Text.StringBuilder
        Dim MyResults As New Results

        Try

            If RequestAction = RequestAction.SaveNew Then
                Sql.Append("INSERT INTO Codes_CarrierID (CarrierID, Description, LogicalDelete, AddDate, ChangeDate)")
                Sql.Append(" Values ")
                Sql.Append("(" & cCommon.StrOutHandler(txtCarrierID.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtDescription.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("0, ")
                Sql.Append("'" & cCommon.GetServerDateTime & "', ")
                Sql.Append("'" & cCommon.GetServerDateTime & "') ")

            Else
                Sql.Append("UPDATE Codes_CarrierID Set ")
                Sql.Append("CarrierID = " & cCommon.StrOutHandler(txtCarrierID.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("Description = " & cCommon.StrOutHandler(txtDescription.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("LogicalDelete = 0, ")
                Sql.Append("ChangeDate = '" & cCommon.GetServerDateTime & "' ")
                '  Sql.Append(" WHERE CarrierID = '" & cSess.CarrierID & "'")
                Sql.Append(" WHERE CarrierID = " & cCommon.StrOutHandler(cSess.CarrierID, False, StringTreatEnum.SideQts_SecApost))
            End If

            cCommon.ExecuteNonQuery(Sql.ToString)

            If Request.Form("hdConfirm") = "yes" Then
                Sql.Length = 0
                'Sql.Append("UPDATE UserAppointments SET CarrierID = '" & txtCarrierID.Text & "' WHERE CarrierID = '" & cSess.CarrierID & "'")
                Sql.Append("UPDATE UserAppointments SET CarrierID = " & cCommon.StrOutHandler(txtCarrierID.Text, False, StringTreatEnum.SideQts_SecApost & " WHERE CarrierID = " & cCommon.StrOutHandler(cSess.CarrierID, False, StringTreatEnum.SideQts_SecApost)))
                cCommon.ExecuteNonQuery(Sql.ToString)
            End If

            cSess.CarrierID = txtCarrierID.Text

            MyResults.Success = True
            MyResults.Msg = "Update complete."

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Msg = "Unable to save record."
        End Try
        Return MyResults
    End Function

    Private Sub DisplayPage(ByVal ResponseAction As ResponseAction)
        Dim i As Integer
        Dim dt As DataTable

        Try

            ' ___ Set the attributes for controls
            txtDescription.Attributes.Add("onkeyup", "return legallength(this, 100)")

            ' ___ Heading/CarrierID
            Select Case ResponseAction
                Case ResponseAction.DisplayBlank, ResponseAction.DisplayUserInputNew
                    litHeading.Text = "New Carrier"
                Case ResponseAction.DisplayExisting, ResponseAction.DisplayUserInputExisting
                    If cRights.HasThisRight(cRights.CarrierEdit) Then
                        litHeading.Text = "Edit Carrier"
                    Else
                        litHeading.Text = "View Carrier"
                    End If
            End Select

            FormatControls()

            If ResponseAction = ResponseAction.DisplayExisting Then

                ' ___ Get the data
                '  dt = cCommon.GetDT("Select * From Codes_CarrierID Where CarrierID = '" & cSess.CarrierID & "'")
                dt = cCommon.GetDT("Select * From Codes_CarrierID Where CarrierID = " & cCommon.StrOutHandler(cSess.CarrierID, False, StringTreatEnum.SideQts_SecApost))
                txtCarrierID.Text = cCommon.StrInHandler(dt.Rows(0)("CarrierID"))
                txtDescription.Text = cCommon.StrInHandler(dt.Rows(0)("Description"))
            End If

        Catch ex As Exception
            Throw New Exception("Error #254: CarrierMaintain DisplayPage. " & ex.Message)
        End Try
    End Sub

    Private Sub FormatControls()
        Try

            If cRights.HasThisRight(cRights.CarrierEdit) Then
                litUpdate.Text = "<input onclick='Update()' type='button' value='Update'>"
                Style.AddStyle(txtCarrierID, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtDescription, Style.StyleType.NormalEditable, 300, True)
            Else
                Style.AddStyle(txtCarrierID, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtDescription, Style.StyleType.NoneditableGrayed, 300, True)
            End If

        Catch ex As Exception
            Throw New Exception("Error #255: CarrierMaintain FormatControls. " & ex.Message)
        End Try
    End Sub
End Class
