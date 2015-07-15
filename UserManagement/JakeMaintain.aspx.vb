Imports System.Data.SqlClient

Public Class JakeMaintain
    Inherits System.Web.UI.Page
    Protected WithEvents litMsg As System.Web.UI.WebControls.Literal
    Protected WithEvents litHeading As System.Web.UI.WebControls.Literal
    Protected WithEvents lblCurrentRights As System.Web.UI.WebControls.Label
    Protected WithEvents litResponseAction As System.Web.UI.WebControls.Literal
    Protected WithEvents Textbox4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents ddQAStatus As System.Web.UI.WebControls.DropDownList
    Protected WithEvents ddTechStatus As System.Web.UI.WebControls.DropDownList
    Protected WithEvents txtTechComments As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtQAComments As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtIssueSummary As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtIssueDescription As System.Web.UI.WebControls.TextBox
    Private Common As New Common
    Private AppSession As New AppSession()
    Protected WithEvents litUpdate As System.Web.UI.WebControls.Literal
    Private Rights As RightsClass
    Private cLoggedInUserID As String
    Protected WithEvents txtOriginator As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtTechStatus As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtQAStatus As System.Web.UI.WebControls.TextBox
    Protected WithEvents litHiddens As System.Web.UI.WebControls.Literal
    Protected WithEvents txtTicketNum As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtOpenDt As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCloseDt As System.Web.UI.WebControls.TextBox
    Private cJakeId As String

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
        Dim hdResponseAction As String

        Try

            ' ___Return to Company Worklist
            If Request.Form("hdAction") = "return" Then
                Response.Redirect("JakeWorklist.aspx")
            End If

            ' ___ Save the JakeID
            If Not Page.IsPostBack Then
                If Request.QueryString("JakeId") = "" Or Request.QueryString("JakeId") = Nothing Then
                    cJakeId = String.Empty
                Else
                    cJakeId = Request.QueryString("JakeId")
                End If
            Else
                cJakeId = Request.Form("hdJakeId")
            End If

            ' ___ Restore  session
            cLoggedInUserID = AppSession.RestoreSession(Page)

            ' ___ Right check
            Rights = New RightsClass(cLoggedInUserID, Page)
            Dim RightsRqd(0) As String
            RightsRqd.SetValue(Rights.JakeEdit, 0)
            Rights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = Common.GetCurRightsHidden(Rights.RightsColl)

            ' ___ Determine the RequestAction.
            If Page.IsPostBack And Request.Form("hdAction") = "update" Then
                hdResponseAction = Request.Form("hdResponseAction")
                If hdResponseAction = "DisplayBlank" Or hdResponseAction = "DisplayUserInputNew" Then
                    RequestAction = RequestAction.SaveNew
                ElseIf hdResponseAction = "DisplayExisting" Or hdResponseAction = "DisplayUserInputExisting" Then
                    RequestAction = RequestAction.SaveExisting
                End If

            Else
                If cJakeId = String.Empty Then
                    RequestAction = RequestAction.CreateNew
                Else
                    RequestAction = RequestAction.LoadExisting
                End If
            End If


            ' __ Execute the Request Action.
            Select Case RequestAction

                Case RequestAction.CreateNew
                    ResponseAction = ResponseAction.DisplayBlank
                    DisplayPage(ResponseAction)

                Case RequestAction.SaveNew
                    Results = IsDataValid()
                    If Results.Success Then
                        PerformSave(RequestAction)
                        ResponseAction = ResponseAction.DisplayExisting
                        DisplayPage(ResponseAction)
                        litMsg.Text = "<script language='javascript'>alert('Update complete.')</script>"
                    Else
                        ResponseAction = ResponseAction.DisplayUserInputNew
                        DisplayPage(ResponseAction)
                        litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"
                    End If

                Case RequestAction.LoadExisting
                    ResponseAction = ResponseAction.DisplayExisting
                    DisplayPage(ResponseAction)

                Case RequestAction.SaveExisting
                    Results = IsDataValid()
                    If Results.Success Then
                        PerformSave(RequestAction)
                        ResponseAction = ResponseAction.DisplayExisting
                        DisplayPage(ResponseAction)
                        litMsg.Text = "<script language='javascript'>alert('Update complete.')</script>"
                    ElseIf Results.ObtainConfirm Then
                        ResponseAction = ResponseAction.DisplayUserInputExisting
                        DisplayPage(ResponseAction)
                        litMsg.Text = "<script language='javascript'>" & Results.Msg & "</script>"
                    Else
                        ResponseAction = ResponseAction.DisplayExisting
                        DisplayPage(ResponseAction)
                        litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"
                    End If

                Case RequestAction.ConfirmResults
                    hdResponseAction = Request.Form("hdResponseAction")
                    If hdResponseAction = "DisplayBlank" Or hdResponseAction = "DisplayUserInputNew" Then
                        If Request.Form("hdConfirm") = "yes" Then
                            RequestAction = RequestAction.SaveNew
                            PerformSave(RequestAction)
                            ResponseAction = ResponseAction.DisplayExisting
                            DisplayPage(ResponseAction)
                            litMsg.Text = "<script language='javascript'>alert('Update complete.')</script>"
                        Else
                            ResponseAction = ResponseAction.DisplayUserInputNew
                            DisplayPage(ResponseAction)
                        End If
                    ElseIf hdResponseAction = "DisplayExisting" Or hdResponseAction = "DisplayUserInputExisting" Then
                        If Request.Form("hdConfirm") = "yes" Then
                            RequestAction = RequestAction.SaveExisting
                            PerformSave(RequestAction)
                            ResponseAction = ResponseAction.DisplayExisting
                            DisplayPage(ResponseAction)
                            litMsg.Text = "<script language='javascript'>alert('Update complete.')</script>"
                        Else
                            ResponseAction = ResponseAction.DisplayUserInputExisting
                            DisplayPage(ResponseAction)
                        End If
                    End If
            End Select

            litResponseAction.Text = "<input type='hidden' name='hdResponseAction' value = '" & ResponseAction.ToString & "'>"

        Catch ex As Exception
            Throw New Exception("Error", ex)
        End Try
    End Sub

    Private Function IsDataValid() As Results
        Dim Results As New Results
        Dim ErrColl As New Collection

        Common.ValidateStringField(ErrColl, txtIssueSummary.Text, 1, "issue summary not provided")
        Common.ValidateStringField(ErrColl, txtIssueDescription.Text, 1, "issue description not provided")

        If ErrColl.Count = 0 Then
            Results.Success = True
        Else
            Results.Success = False
            Results.Msg = Common.ErrCollToString(ErrColl, "Not saved. Please correct the following:")
        End If

        Return Results
    End Function

    Private Function PerformSave(ByVal RequestAction As RequestAction) As Results
        Dim MyResults As New Results
        Dim CmdAsst As New CmdAsst(CommandType.StoredProcedure, "JakeUpd")
        Dim QueryPack As CmdAsst.QueryPack

        If RequestAction = RequestAction.SaveNew Then
            CmdAsst.AddBit("Action", 1)
            CmdAsst.AddInt("JakeId", 0)
            CmdAsst.AddVarChar("Originator", cLoggedInUserID, 50)
        Else
            CmdAsst.AddBit("Action", 2)
            CmdAsst.AddInt("JakeId", cJakeId)
            CmdAsst.AddVarChar("Originator", txtOriginator.Text, 50)
        End If
        CmdAsst.AddVarChar("IssueSummary", txtIssueSummary.Text, 200)
        CmdAsst.AddVarChar("IssueDescription", txtIssueDescription.Text, 100)
        CmdAsst.AddInt("TechStatus", ddTechStatus.SelectedItem.Value)
        CmdAsst.AddVarChar("TechComments", txtTechComments.Text, 100)
        CmdAsst.AddInt("QAStatus", ddQAStatus.SelectedItem.Value)
        CmdAsst.AddVarChar("QAComments", txtQAComments.Text, 100)

        QueryPack = CmdAsst.Execute

        MyResults.Success = QueryPack.Success
        If QueryPack.Success Then
            MyResults.Msg = "Update complete."
            cJakeId = QueryPack.dt.rows(0)("JakeId")
        Else
            MyResults.Msg = "Unable to update record."
        End If
        Return MyResults
    End Function

    Private Sub DisplayPage(ByVal ResponseAction As ResponseAction)
        Dim JakeId As Integer
        Dim dt As Data.DataTable
        Dim i, j As Integer

        ' ___ Set the attributes for controls
        txtIssueSummary.Attributes.Add("onkeyup", "return legallength(this, 200)")
        txtIssueDescription.Attributes.Add("onkeyup", "return legallength(this, 1000)")
        txtTechComments.Attributes.Add("onkeyup", "return legallength(this, 1000)")
        txtQAComments.Attributes.Add("onkeyup", "return legallength(this, 1000)")
        ' txtIssueDescription.Attributes.Add("onkeyup", "return IsValidChar(this)")

        '' ___ ASP.NET does not render these attributes correctly in Firefox
        'txtIssueSummary.Attributes.Add("style", "width:445px")
        'txtIssueDescription.Attributes.Add("style", "width:445px; FONT: 10pt Arial, Helvetica, sans-serif;")
        'txtTechComments.Attributes.Add("style", "width:445px; FONT: 10pt Arial, Helvetica, sans-serif;")
        'txtQAComments.Attributes.Add("style", "width:445px; FONT: 10pt Arial, Helvetica, sans-serif;")

        ' ___ Heading/JakeId
        Select Case ResponseAction
            Case ResponseAction.DisplayBlank, ResponseAction.DisplayUserInputNew
                litHeading.Text = "New Jake Item"
            Case ResponseAction.DisplayExisting, ResponseAction.DisplayUserInputExisting
                If Rights.HasThisRight(Rights.JakeEdit) Then
                    litHeading.Text = "Edit Jake Item"
                Else
                    litHeading.Text = "View Jake Item"
                End If
        End Select

        ' ___ View/Edit
        txtTicketNum.ReadOnly = True
        txtOpenDt.ReadOnly = True
        txtCloseDt.ReadOnly = True
        txtOriginator.ReadOnly = True
        txtTicketNum.Attributes.Add("style", "width:400;FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;")
        txtOpenDt.Attributes.Add("style", "width:400;FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;")
        txtCloseDt.Attributes.Add("style", "width:400;FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;")
        txtOriginator.Attributes.Add("style", "width:400;FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;background: #eeeedd;")

        If Rights.HasThisRight(Rights.JakeEdit) Then
            litUpdate.Text = "<input onclick='Update()' type='button' value='Update'>"

            txtIssueSummary.ReadOnly = False
            txtIssueDescription.ReadOnly = False

            ddTechStatus.Visible = True
            txtTechStatus.Visible = False

            txtTechComments.ReadOnly = False

            ddQAStatus.Visible = True
            txtQAStatus.Visible = False

            txtQAComments.Visible = True
            txtQAComments.ReadOnly = False

            txtIssueDescription.Attributes.Add("style", "width:400;FONT: 10pt Arial, Helvetica, sans-serif;")
            txtTechComments.Attributes.Add("style", "width:400;FONT: 10pt Arial, Helvetica, sans-serif;")
            txtQAComments.Attributes.Add("style", "width:400;FONT: 10pt Arial, Helvetica, sans-serif;")
        Else


            txtIssueSummary.ReadOnly = True
            txtIssueDescription.ReadOnly = True

            ddTechStatus.Visible = False
            txtTechStatus.Visible = True
            txtTechStatus.ReadOnly = True

            txtTechComments.ReadOnly = True

            ddQAStatus.Visible = False
            txtQAStatus.Visible = True
            txtQAStatus.ReadOnly = True
            txtTechComments.ReadOnly = True

            txtQAComments.Visible = True
            txtQAComments.ReadOnly = True
            txtIssueDescription.Attributes.Add("style", "width:400;FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;;background: #eeeedd;overflow:hidden")
            txtTechComments.Attributes.Add("style", "width:400;FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;;background: #eeeedd;overflow:hidden")
            txtQAComments.Attributes.Add("style", "width:400;FONT: 10pt Arial, Helvetica, sans-serif;border-width:1px;;background: #eeeedd;overflow:hidden")
        End If


        ' ___ Get the data for ResponseAction.DisplayExisting only
        If ResponseAction = ResponseAction.DisplayExisting Then

            ' ___ Get the data
            dt = Common.GetDT("SELECT * From Jake WHERE JakeId = '" & cJakeId & "'")

            ' ___ Populate the controls
            txtTicketNum.Text = Common.StrInHandler(dt.Rows(0)("TicketNum"))
            txtOpenDt.Text = Common.DateInHandler(dt.Rows(0)("OpenDt"))
            txtCloseDt.Text = Common.DateInHandler(dt.Rows(0)("CloseDt"))

            txtOriginator.Text = Common.StrInHandler(dt.Rows(0)("Originator"))
            txtIssueSummary.Text = Common.StrInHandler(dt.Rows(0)("IssueSummary"))
            txtIssueDescription.Text = Common.StrInHandler(dt.Rows(0)("IssueDescription"))


            ddTechStatus.Items.FindByValue(dt.Rows(0)("TechStatus")).Selected = True
            txtTechStatus.Text = Common.StrInHandler(dt.Rows(0)("TechStatus"))
            txtTechComments.Text = Common.StrInHandler(dt.Rows(0)("TechComments"))

            ddQAStatus.Items.FindByValue(dt.Rows(0)("QAStatus")).Selected = True
            txtQAStatus.Text = Common.StrInHandler(dt.Rows(0)("QAStatus"))

            txtQAComments.Text = Common.StrInHandler(dt.Rows(0)("QAComments"))
            txtQAComments.Text = Common.StrInHandler(dt.Rows(0)("QAComments"))

        End If

        litHiddens.Text = "<input type='hidden' name='hdJakeId' value='" & cJakeId & "'>"

    End Sub

End Class

