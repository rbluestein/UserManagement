Imports System.Data.SqlClient

Public Class OrigDG

#Region " Declarations "
    Public Event ChildDTRequest(ByRef ChildText As String, ByVal DataFldName As String, ByVal Value As String)
    'Public Event ChildDTRequest2(ByRef ChildText As String, ByVal DataFldName As String, ByVal Value As String, ByVal Parm2 As String)
    'Public Event ChildDTRequest3(ByRef ChildText As String, ByVal DataFldName As String, ByVal Value As String, ByVal Parm2 As String, ByVal Parm3 As String)
    'Public Event ChildDTRequest4(ByRef ChildText As String, ByVal DataFldName As String, ByVal Value As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
    Private cChildTables As ChildTablesClass
    Private cFilter As DG.Filter
    Private cMenu As DG.Menu
    Private cColumnColl As New Collection
    Private cSortDataColl As New Collection
    Private cTemplateColl As New Collection
    Private Common As New Common
    Private Rights As RightsClass
    Private cKeyFieldName As String
    Private cTableDef As String
    Private cUseDefaultTableDef As Boolean
    Private cDefaultColl As New Collection
    Private cCheckboxToggleColl As Collection
    Private cHiddenColumnColl As Collection
    Private cLeftPadding As String
    Private cDefaultSortField As String
    Private cFilterOperationMode As FilterOperationModeEnum = FilterOperationMode.NoFilter
    Private cFilterInitialShowHide As FilterInitialShowHideEnum
    Private cRecordsInitialShowHide As RecordsInitialShowHideEnum = RecordsInitialShowHideEnum.RecordsInitialShow
    Private cShowFilter As Boolean
    Private cAttachNewButton As Boolean
    Private cNewButtonRight As String
    Private cFormatAsSubTable As Boolean
    Private cNewOrderByField As String
    Private cNewSortDirection As String
#End Region

#Region " Enums "
    Public Enum Justify
        none = 0
        left = 1
        center = 2
        right = 3
    End Enum

    Public Enum ColumnType
        Databound = 1
        Template = 2
        Link = 3
        CheckboxToggle = 4
        Hidden = 5
        FreeForm = 6
        ChildTableSelect = 7
        [Boolean] = 8
        [Date] = 9
    End Enum

    Public Enum OrderByType
        Initial = 1
        Recurring = 2
        Field = 3
        ReturnToPage = 4
    End Enum

    Public Enum FilterOperationModeEnum
        NoFilter = 1
        FilterAlwaysOn = 2
        FilterAlwaysOff = 3
        FilterSwitchable = 4
    End Enum

    Public Enum FilterInitialShowHideEnum
        FilterInitialShow = 1
        FilterInitalHide = 2
    End Enum

    Public Enum RecordsInitialShowHideEnum
        RecordsInitialShow = 1
        RecordsInitialHide = 2
    End Enum
#End Region

#Region " Properties "
    Public WriteOnly Property FormatAsSubTable()
        Set(ByVal Value)
            cFormatAsSubTable = Value
        End Set
    End Property
    Public Property ChildTables()
        Get
            Return cChildTables
        End Get
        Set(ByVal Value)
            cChildTables = Value
        End Set
    End Property
    Public ReadOnly Property FilterOperationMode()
        Get
            Return cFilterOperationMode
        End Get
    End Property
    Public ReadOnly Property FilterInitialShowHide()
        Get
            Return cFilterInitialShowHide
        End Get
    End Property
    Public ReadOnly Property RecordsInitialShowHide()
        Get
            Return cRecordsInitialShowHide
        End Get
    End Property
#End Region

#Region " Constructors "
    Public Sub New()
        Dim dt As New DataTable
        'Dim ChildDG As New DG("Dog", Rights, True, "", "")
        'ChildDG.GetChildText(dt, "KeyFldName", "KeyFldValue")
    End Sub

    Public Sub New(ByVal KeyFieldName As String, ByVal Rights As RightsClass, ByVal UseDefaultTableDef As Boolean, ByVal TableDef As String, ByVal DefaultSortField As String)
        cKeyFieldName = KeyFieldName
        Dim aRights(0) As String
        cUseDefaultTableDef = UseDefaultTableDef
        cTableDef = TableDef
        Me.Rights = Rights
        cDefaultColl.Add("<table class='DG' style='LEFT: 150px; POSITION: absolute; TOP: 214px' cellSpacing='0' cellPadding='0' width='650' border='0'>", "StandardTableDef")
        cDefaultColl.Add("<table class='DGEmbedded' cellSpacing='0' cellPadding='0' width='100%' border='0'>", "EmbeddedTableDef")
        'cLeftPadding = " style='padding-left:10px' "
        cLeftPadding = ""
        cDefaultSortField = DefaultSortField
    End Sub
#End Region

#Region " Generate SQL "
    Public Sub GenerateSQL(ByRef Sql As String, ByRef ShowFilter As Boolean, ByVal SecurityWhereClause As String, ByVal OrderByType As DG.OrderByType, ByVal Request As HttpRequest, Optional ByVal OmitWhere As Boolean = False)
        Dim ShowRecords As Boolean
        Dim InternalFilterWhereClause As String
        Dim WorkingSql As String = String.Empty

        If cFilter Is Nothing Then
            If SecurityWhereClause <> Nothing Then

                If OmitWhere Then
                    Sql &= " AND " & SecurityWhereClause
                Else
                    Sql &= " WHERE " & SecurityWhereClause
                End If

            End If
            Sql &= " ORDER BY " & cNewOrderByField
            If cNewSortDirection = "A" Then
                Sql &= " ASC"
            Else
                Sql &= " DESC"
            End If
        Else

            ' ___ Show/hide the internal filter. Record internal filter values
            ShowFilter = GetShowFilter(OrderByType, Request)
            ShowRecords = GetShowRecords(OrderByType)
            RecordFilterValues(ShowFilter, Request)

            ' new
            If ShowRecords And FilterOperationMode <> DG.FilterOperationModeEnum.NoFilter Then
                InternalFilterWhereClause = GetInternalFilterWhereClause()
            End If
            If Not ShowRecords Then
                InternalFilterWhereClause = "  0 = 1 "
            End If

            If SecurityWhereClause = Nothing Then
                SecurityWhereClause = String.Empty
            End If

            If SecurityWhereClause.Length = 0 Then
                If InternalFilterWhereClause.Length = 0 Then
                    ' No action
                Else
                    WorkingSql = InternalFilterWhereClause
                End If
            Else
                If InternalFilterWhereClause.Length = 0 Then
                    WorkingSql = SecurityWhereClause
                Else
                    WorkingSql = SecurityWhereClause & " AND " & InternalFilterWhereClause
                End If
            End If

            If WorkingSql.Length > 0 Then
                If OmitWhere Then
                    Sql &= " AND " & WorkingSql
                Else
                    Sql &= " WHERE " & WorkingSql
                End If
            End If

            Sql &= " ORDER BY " & cNewOrderByField
            If cNewSortDirection = "A" Then
                Sql &= " ASC"
            Else
                Sql &= " DESC"
            End If
        End If


        ''old

        '' ___ Where clause: Get the restrictions imposed by the filter
        'If ShowRecords And FilterOperationMode <> DG.FilterOperationModeEnum.NoFilter Then
        '    FilterWhereClause = GetFilterWhereClause()
        'End If

        '' ___ Block all records
        'If Not ShowRecords Then
        '    FilterWhereClause = "  0 = 1 "
        'End If

        'If ShowRecords Then
        '    If FilterWhereClause.Length = 0 Then
        '        If SecurityWhereClause = Nothing Then
        '        Else
        '            Sql &= " WHERE " & SecurityWhereClause
        '        End If
        '    Else
        '        If SecurityWhereClause = Nothing Then
        '            Sql &= FilterWhereClause
        '        Else
        '            NewSql &= FilterWhereClause & " AND " & SecurityWhereClause
        '        End If
        '    End If
        'Else
        '    Sql &= " WHERE 0 = 1 "
        'End If



    End Sub
#End Region

#Region " Handle sorting "
    Public Sub SetSortElements(ByVal OrderByField As String, ByVal OrderByType As OrderByType, ByRef NewOrderByField As String, ByRef NewSortDirection As String)
        Dim CurSortDirection As String

        If OrderByType = OrderByType.Initial Then
            NewOrderByField = cDefaultSortField
            NewSortDirection = "A"
        ElseIf OrderByType = OrderByType.Recurring Then
            NewOrderByField = GetLastFieldSorted()
            If NewOrderByField = "" Then
                NewOrderByField = cDefaultSortField
                NewSortDirection = "A"
            Else
                NewSortDirection = GetSortDirection(NewOrderByField)
            End If
        ElseIf OrderByType = OrderByType.Field Then
            NewOrderByField = OrderByField
            CurSortDirection = GetSortDirection(OrderByField)
            Select Case CurSortDirection
                Case "A"
                    NewSortDirection = "D"
                Case "N", "D"
                    NewSortDirection = "A"
            End Select
        End If
        SetLastFieldSorted(NewOrderByField)
        SetSortDirection(NewOrderByField, NewSortDirection)
        cNewOrderByField = NewOrderByField
        cNewSortDirection = NewSortDirection
    End Sub

    Private Sub AppendViewstateString(ByVal SortExpression As String)
        Dim Position As String
        If SortExpression <> Nothing AndAlso SortExpression.Length > 0 Then
            Position = CStr(101 + cSortDataColl.Count).ToString
            cSortDataColl.Add(New SortItem(SortExpression, Position), SortExpression)
        End If
    End Sub

    Public Sub UpdateViewstate(ByVal Viewstate As String)
        Dim i As Integer
        For i = 1 To cSortDataColl.Count
            cSortDataColl(i).SortDirection = Viewstate.Substring((5 * (i - 1)) + 3, 1)
            cSortDataColl(i).LastFieldSorted = Viewstate.Substring((5 * (i - 1)) + 4, 1)
        Next
    End Sub

    Public Function GetViewStateString() As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder
        For i = 1 To cSortDataColl.Count
            sb.Append(cSortDataColl(i).Position & cSortDataColl(i).SortDirection & cSortDataColl(i).LastFieldSorted)
        Next
        Return sb.ToString
    End Function

    Private Function GetSortDirection(ByVal FldName As String) As String
        If FldName = String.Empty Then
            Return String.Empty
        Else
            Return cSortDataColl(FldName).SortDirection
        End If
    End Function

    Private Sub SetSortDirection(ByVal FldName As String, ByVal Value As String)
        If cSortDataColl.Count > 0 Then
            cSortDataColl(FldName).SortDirection = Value
        End If
    End Sub

    Public Function GetLastFieldSorted() As String
        Dim i As Integer
        For i = 1 To cSortDataColl.Count
            If cSortDataColl(i).LastFieldSorted = "Y" Then
                Return cSortDataColl(i).Name
            End If
        Next
        If i = cSortDataColl.Count + 1 Then
            Return String.Empty
        End If
    End Function

    Public Sub SetLastFieldSorted(ByVal FldName As String)
        Dim i As Integer
        If cSortDataColl.Count > 0 Then
            For i = 1 To cSortDataColl.Count
                cSortDataColl(i).LastFieldSorted = "F"
            Next
            If FldName <> String.Empty Then
                cSortDataColl(FldName).LastFieldSorted = "T"
            End If
        End If
    End Sub
#End Region

#Region " Handle the menu "
    Public Function AttachMenu() As DG.Menu
        cMenu = New DG.Menu
        Return cMenu
    End Function
#End Region

#Region " Handle the filter "
    Public Function GetFilter() As DG.Filter
        Return cFilter
    End Function
    Public Function AttachFilter(ByVal FilterOperationMode As FilterOperationModeEnum, ByVal FilterInitialShowHide As FilterInitialShowHideEnum, ByVal RecordsInitialShowHide As RecordsInitialShowHideEnum) As DG.Filter
        cFilterOperationMode = FilterOperationMode
        cFilterInitialShowHide = FilterInitialShowHide
        cRecordsInitialShowHide = RecordsInitialShowHide
        cFilter = New DG.Filter
        Return cFilter
    End Function

    Public Function GetShowFilter(ByVal OrderByType As OrderByType, ByVal Request As HttpRequest) As Boolean
        Dim ShowFilter As Boolean

        Select Case cFilterOperationMode
            Case DG.FilterOperationModeEnum.NoFilter, DG.FilterOperationModeEnum.FilterAlwaysOff
                ShowFilter = False
            Case DG.FilterOperationModeEnum.FilterAlwaysOn
                ShowFilter = True
            Case DG.FilterOperationModeEnum.FilterSwitchable
                Select Case OrderByType
                    Case DG.OrderByType.Initial
                        Select Case cFilterInitialShowHide
                            Case DG.FilterInitialShowHideEnum.FilterInitialShow
                                ShowFilter = True
                            Case DG.FilterInitialShowHideEnum.FilterInitalHide
                                ShowFilter = False
                        End Select

                    Case Else

                        Select Case Request.Form("hdFilterOnResponse")
                            Case "on"
                                Select Case Request("hdFilterShowHideToggle")
                                    Case "0"
                                        ShowFilter = True
                                    Case "1"
                                        ShowFilter = False
                                End Select
                            Case "off"
                                Select Case Request("hdFilterShowHideToggle")
                                    Case "0"
                                        ShowFilter = False
                                    Case "1"
                                        ShowFilter = True
                                End Select
                        End Select
                End Select
        End Select
        cShowFilter = ShowFilter
        Return ShowFilter
    End Function

    Public Function GetShowRecords(ByVal OrderByType As DG.OrderByType) As Boolean
        Dim ShowRecords As Boolean
        Select Case OrderByType
            Case DG.OrderByType.Initial
                Select Case cRecordsInitialShowHide
                    Case DG.FilterInitialShowHideEnum.FilterInitialShow
                        ShowRecords = True
                    Case DG.FilterInitialShowHideEnum.FilterInitalHide
                        ShowRecords = False
                End Select
            Case Else
                ShowRecords = True
        End Select
        Return ShowRecords
    End Function

    Public Sub RecordFilterValues(ByVal ShowFilter As Boolean, ByRef Request As HttpRequest)
        Dim i As Integer

        ' ___ Internal filter
        Select Case Request.Form("hdFilterOnResponse")
            Case "on"
                Select Case ShowFilter
                    Case True
                        ' Write the values
                        For i = 1 To cFilter.Coll.count
                            cFilter.Coll(i).SetValue(Request.Form(cFilter.Coll(i).CtlName))
                        Next
                    Case False
                        ' Clear it out
                        For i = 1 To cFilter.Coll.count
                            cFilter.Coll(i).SetValue("")
                        Next
                End Select
            Case "off"
                ' No action
        End Select

    End Sub

    Public Function GetInternalFilterWhereClause() As String
        Dim Coll As New Collection
        Dim i As Integer
        Dim InternalFilterWhereClause As New System.Text.StringBuilder
        Dim SelectedValue As String

        For i = 1 To cFilter.Coll.count
            If cFilter.Coll(i).GetValue <> String.Empty Then
                'If WhereClause.Length = 0 Then
                '    WhereClause.Append(" Where ")
                'Else : WhereClause.Append(" and ")
                'End If
                If cFilter.Coll(i).IsTextbox Then
                    If cFilter.Coll(i).FilterField = Nothing Then
                        'FilterWhereClause.Append(Common.StrInHandler(cFilter.Coll(i).DataFldName) & " LIKE  '" & cFilter.Coll(i).GetValue & "%'")
                        Coll.Add(Common.StrInHandler(cFilter.Coll(i).DataFldName) & " LIKE  '" & cFilter.Coll(i).GetValue & "%'")
                    Else
                        'FilterWhereClause.Append(Common.StrInHandler(cFilter.Coll(i).FilterField) & " LIKE  '" & cFilter.Coll(i).GetValue & "%'")
                        Coll.Add(Common.StrInHandler(cFilter.Coll(i).FilterField) & " LIKE  '" & cFilter.Coll(i).GetValue & "%'")
                    End If
                Else
                    'FilterWhereClause.Append(cFilter.Coll(i).DataFldName & " = '" & cFilter.Coll(i).GetValue & "'")

                    If cFilter.Coll(i).IsStandard Then
                        Coll.Add(cFilter.Coll(i).DataFldName & " = '" & cFilter.Coll(i).GetValue & "'")
                    ElseIf cFilter.Coll(i).IsExtended Then
                        SelectedValue = cFilter.Coll(i).GetValue
                        Coll.Add(cFilter.Coll(i).Coll(SelectedValue).Sql)
                    End If
                End If
            End If
        Next

        For i = 1 To Coll.Count
            If i = 1 Then
                InternalFilterWhereClause.Append(Coll(i))
            Else
                InternalFilterWhereClause.Append(" and " & Coll(i))
            End If
        Next
        Return InternalFilterWhereClause.ToString
    End Function


#End Region

#Region " Attach components "
    Public Sub AddNewButton(ByVal Right As String)
        cAttachNewButton = True
        cNewButtonRight = Right
    End Sub

    Public Sub AddDataBoundColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
        cColumnColl.Add(New DataBoundColumnItems(ColumnType.Databound, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes), ItemName)
        AppendViewstateString(SortExpression)
    End Sub

    Public Sub AddDateColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
        cColumnColl.Add(New DateColumn(ColumnType.Date, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes), ItemName)
        AppendViewstateString(SortExpression)
    End Sub

    'Public Sub AddFreeFormColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String, ByVal Text As String)
    '    cColumnColl.Add(New FreeFormColumn(ColumnType.FreeForm, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes, Text), ItemName)
    '    AppendViewstateString(SortExpression)
    'End Sub

    Public Sub AddChildTableSelectColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal EnableFldName As String, ByVal Title As String, ByVal Attributes As String)
        Dim ChildTableSelectColumn As New ChildTableSelectColumn(cChildTables, ColumnType.ChildTableSelect, ItemName, DataFldName, Nothing, Nothing, True, EnableFldName, Nothing, Title, Attributes, Nothing, Nothing, Nothing)
        cColumnColl.Add(ChildTableSelectColumn, ItemName, 1)
        cChildTables.ChildTableSelectColumn = ChildTableSelectColumn
        ' cColumnColl.Add(New ChildTableSelectColumn(cChildTables, ColumnType.ChildTableSelect, ItemName, DataFldName, Nothing, Nothing, True, EnableFldName, Nothing, Title, Attributes, Nothing, Nothing, Nothing), ItemName, 1)
    End Sub

    Public Sub AddChildTableSelectColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal EnableFldName As String, ByVal Title As String, ByVal Attributes As String, ByVal Parm2 As String)
        Dim ChildTableSelectColumn As New ChildTableSelectColumn(cChildTables, ColumnType.ChildTableSelect, ItemName, DataFldName, Nothing, Nothing, True, EnableFldName, Nothing, Title, Attributes, Parm2, Nothing, Nothing)
        cColumnColl.Add(ChildTableSelectColumn, ItemName, 1)
        cChildTables.ChildTableSelectColumn = ChildTableSelectColumn
        'cColumnColl.Add(New ChildTableSelectColumn(cChildTables, ColumnType.ChildTableSelect, ItemName, DataFldName, Nothing, Nothing, True, EnableFldName, Nothing, Title, Attributes, Parm2, Nothing, Nothing), ItemName, 1)
    End Sub

    Public Sub AddChildTableSelectColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal EnableFldName As String, ByVal Title As String, ByVal Attributes As String, ByVal Parm2 As String, ByVal Parm3 As String)
        Dim ChildTableSelectColumn As New ChildTableSelectColumn(cChildTables, ColumnType.ChildTableSelect, ItemName, DataFldName, Nothing, Nothing, True, EnableFldName, Nothing, Title, Attributes, Parm2, Parm3, Nothing)
        cColumnColl.Add(ChildTableSelectColumn, ItemName, 1)
        cChildTables.ChildTableSelectColumn = ChildTableSelectColumn
        ' cColumnColl.Add(New ChildTableSelectColumn(cChildTables, ColumnType.ChildTableSelect, ItemName, DataFldName, Nothing, Nothing, True, EnableFldName, Nothing, Title, Attributes, Parm2, Parm3, Nothing), ItemName, 1)
    End Sub

    Public Sub AddChildTableSelectColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal EnableFldName As String, ByVal Title As String, ByVal Attributes As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
        Dim ChildTableSelectColumn As New ChildTableSelectColumn(cChildTables, ColumnType.ChildTableSelect, ItemName, DataFldName, Nothing, Nothing, True, EnableFldName, Nothing, Title, Attributes, Parm2, Parm3, Parm4)
        cColumnColl.Add(ChildTableSelectColumn, ItemName, 1)
        cChildTables.ChildTableSelectColumn = ChildTableSelectColumn
        ' cColumnColl.Add(New ChildTableSelectColumn(cChildTables, ColumnType.ChildTableSelect, ItemName, DataFldName, Nothing, Nothing, True, EnableFldName, Nothing, Title, Attributes, Parm2, Parm3, Parm4), ItemName, 1)
    End Sub

    Public Sub AddBooleanColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal TrueValue As String, ByVal TrueText As String, ByVal FalseText As String, ByVal Title As String, ByVal Attributes As String)
        cColumnColl.Add(New BooleanColumn(ColumnType.Boolean, ItemName, DataFldName, HeaderText, SortExpression, Visible, TrueValue, TrueText, FalseText, Title, Attributes), ItemName)
        AppendViewstateString(SortExpression)
    End Sub

    Public Sub AddFreeFormColumn(ByVal ItemName As String, ByVal CellText As String, ByVal Header As String, ByVal Title As String, ByVal Visible As Boolean, ByVal Attributes As String)
        cColumnColl.Add(New FreeFormColumn(ColumnType.FreeForm, ItemName, CellText, Header, Title, Visible, Attributes), ItemName)
    End Sub

    Public Sub AddHiddenColumn(ByVal ItemName As String, ByVal DataFldName As String)
        cColumnColl.Add(New HiddenItems(ColumnType.Hidden, ItemName, DataFldName), ItemName)
        If cHiddenColumnColl Is Nothing Then
            cHiddenColumnColl = New Collection
        End If
        cHiddenColumnColl.Add(New Collection, ItemName)
    End Sub

    Public Sub AddLinkColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HRef As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
        cColumnColl.Add(New LinkColumnItems(ColumnType.Link, HRef, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes), ItemName)
        AppendViewstateString(SortExpression)
    End Sub

    Public Sub AddCheckboxToggleColumn(ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal Right As String, ByVal TestFld As String, ByVal TrueText As String, ByVal FalseText As String, ByVal TitleFldName As String, ByVal Attributes As String)
        If Rights.HasThisRight(Right) Then
            cColumnColl.Add(New CheckboxToggleColumnItems(ColumnType.CheckboxToggle, ItemName, DataFldName, HeaderText, SortExpression, Visible, TitleFldName, Attributes, TestFld, TrueText, FalseText), ItemName)
            If cCheckboxToggleColl Is Nothing Then
                cCheckboxToggleColl = New Collection
            End If
            cCheckboxToggleColl.Add(New System.Text.StringBuilder, ItemName)
            AppendViewstateString(SortExpression)
        End If
    End Sub

    'Public Function GetTemplateColumn(ByVal ItemName As String, ByVal HeaderText As String, ByVal Wrap As Boolean, ByVal Visible As Boolean) As TemplateColumn
    '    cColumnColl.Add(New TemplateColumn(ItemName, HeaderText, Wrap, Visible), ItemName)
    '    Return cColumnColl(cColumnColl.Count)
    'End Function

    Public Sub AttachTemplateCol(ByRef TemplateCol As TemplateColumn)
        cColumnColl.Add(TemplateCol, TemplateCol.ItemName)
    End Sub

    Public Function GetCheckboxToggleColl(ByVal ItemName As String) As String
        Dim Value As String
        Value = cCheckboxToggleColl(ItemName).ToString
        Value = Value.Substring(0, Value.Length - 1)
        Return Value
    End Function

    Public Function GetColumnColl(ByVal ItemName) As String
        Dim i As Integer
        Dim Coll As Collection
        Coll = cHiddenColumnColl(ItemName)
        Dim Value As String

        For i = 1 To Coll.Count
            Value &= Coll(i) & "|"
        Next
        Return Value
        Value = Value.Substring(0, Value.Length - 1)
        Return Value
    End Function

    Public Function AttachChildTables(ByVal ItemName As String, ByVal DataFldName As String, ByVal PermissionFldName As String) As DG.ChildTablesClass
        cChildTables = New ChildTablesClass(ItemName, DataFldName, PermissionFldName)
        Return cChildTables
    End Function
#End Region

#Region " GetText et al "
    Public Function GetText(ByRef dt As Data.DataTable) As String
        Dim sb As New System.Text.StringBuilder
        Dim OddRow As Boolean = True
        Dim RowNum As Integer
        Dim ChildText As String

        If cUseDefaultTableDef Then
            sb.Append(cDefaultColl(cTableDef))
        Else
            sb.Append(cTableDef)
        End If
        AddMenu(sb)

        AddHeaderRow(sb)

        If cFilterOperationMode <> FilterOperationModeEnum.NoFilter Then
            AddFilter(sb, "filter")
        End If

        '  For RowNum = 0 To dt.Rows.Count - 1
        '  AddDataRow(dt, RowNum, OddRow, sb)
        '  OddRow = Not OddRow
        '  Next

        'For RowNum = 0 To dt.Rows.Count - 1
        '    AddDataRow(dt, RowNum, OddRow, sb)
        '    If Not cChildTables Is Nothing AndAlso dt.Rows(RowNum)(cColumnColl(1).DataFldName) = cChildTables.Value AndAlso dt.Rows(RowNum)(cColumnColl(1).EnableFldName) = 1 Then
        '        RaiseEvent ChildDTRequest(ChildText, cChildTables.DataFldName, cChildTables.Value)
        '        sb.Append("<tr><td width='20px'>&nbsp;</td><td colspan='" & cColumnColl.Count - 1 & "'>" & ChildText & "</td></tr>")
        '        sb.Append("<td><td>&nbsp;</td></tr>")
        '    Else
        '        If Not cFormatAsSubTable Then
        '            OddRow = Not OddRow
        '        End If
        '    End If
        'Next

        For RowNum = 0 To dt.Rows.Count - 1
            AddDataRow(dt, RowNum, OddRow, sb)
            If Not cChildTables Is Nothing AndAlso dt.Rows(RowNum)(cColumnColl(1).DataFldName) = cChildTables.ChildTableSelectColumn.ParmColl("DataFldName").Value AndAlso dt.Rows(RowNum)(cColumnColl(1).EnableFldName) = 1 Then
                RaiseEvent ChildDTRequest(ChildText, cChildTables.DataFldName, cChildTables.ChildTableSelectColumn.ParmColl("DataFldName").Value)
                'Select Case cChildTables.ChildTableSelectColumn.ParmColl.Count
                '    Case 1
                '        RaiseEvent ChildDTRequest(ChildText, cChildTables.DataFldName, cChildTables.ChildTableSelectColumn.ParmColl("DataFldName").Value)
                '    Case 2
                '        RaiseEvent ChildDTRequest2(ChildText, cChildTables.DataFldName, cChildTables.ChildTableSelectColumn.ParmColl("DataFldName").Value, cChildTables.ChildTableSelectColumn.ParmColl("Parm2").Value)
                '    Case 3
                '        RaiseEvent ChildDTRequest3(ChildText, cChildTables.DataFldName, cChildTables.ChildTableSelectColumn.ParmColl("DataFldName").Value, cChildTables.ChildTableSelectColumn.ParmColl("Parm2").Value, cChildTables.ChildTableSelectColumn.ParmColl("Parm3").Value)
                '    Case 4
                '        RaiseEvent ChildDTRequest4(ChildText, cChildTables.DataFldName, cChildTables.ChildTableSelectColumn.ParmColl("DataFldName").Value, cChildTables.ChildTableSelectColumn.ParmColl("Parm2").Value, cChildTables.ChildTableSelectColumn.ParmColl("Parm3").Value, cChildTables.ChildTableSelectColumn.ParmColl("Parm4").Value)
                'End Select
                sb.Append("<tr><td width='20px'>&nbsp;</td><td colspan='" & cColumnColl.Count - 1 & "'>" & ChildText & "</td></tr>")
                sb.Append("<td><td>&nbsp;</td></tr>")
            Else
                If Not cFormatAsSubTable Then
                    OddRow = Not OddRow
                End If
            End If
        Next


        sb.Append("</table>")
        Return sb.ToString
    End Function

    Private Sub AddMenu(ByRef sb As System.Text.StringBuilder)
        Dim i As Integer
        Dim ShowHideFilter As String
        Dim DisplayNewButton As Boolean
        Dim sbMenu As New System.Text.StringBuilder
        Dim FilterExists As Boolean
        Dim MenuExists As Boolean
        Dim ItemAdded As Boolean
        Dim ColNum As Integer
        Dim ColCount As Integer

        ' ___ How this works
        ' If a filter exists, the menu displays the Go link, then the Show/Hide filter link.
        ' If a shortcut new button is present, it appears next.
        ' All items listed in the menu object appear next.
        ' Whether or not any menu items appear, this methods adds a 30px row.

        ' ___ Get the column count
        For ColNum = 1 To cColumnColl.Count
            If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then
                If cColumnColl(ColNum).Visible Then
                    If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Template Then
                        ColCount += 1
                    End If
                End If
            End If
        Next

        If Not cFilter Is Nothing Then
            FilterExists = True
        End If
        If Not cMenu Is Nothing Then
            MenuExists = True
        End If

        If FilterExists Then
            If cShowFilter Then
                ShowHideFilter = "Hide Filter"
            Else
                ShowHideFilter = "Show Filter"
            End If
        End If

        If cAttachNewButton Then
            If cNewButtonRight = Nothing Then
                DisplayNewButton = True
            Else
                If Rights.HasThisRight(cNewButtonRight) Then
                    DisplayNewButton = True
                End If
            End If
        End If

        If FilterExists Then
            sbMenu.Append("<a href=""javascript:ApplyFilter()"">Go</a>&nbsp;&nbsp;&nbsp;<a href=""javascript:ToggleShowFilter()"">" & ShowHideFilter & "</a>")
        End If

        If DisplayNewButton Then
            If FilterExists Then
                sbMenu.Append("&nbsp;&nbsp;&nbsp;")
            End If
            sbMenu.Append("<a href='javascript:NewRecord()'>New</a>")
        End If

        If MenuExists Then
            If FilterExists Or DisplayNewButton Then
                sbMenu.Append("&nbsp;&nbsp;&nbsp;")
            End If

            For i = 1 To cMenu.Coll.count
                If cMenu.Coll(i).IsVisible And Rights.HasThisRight(cMenu.Coll(i).Right) Then
                    If ItemAdded Then
                        sbMenu.Append("&nbsp;&nbsp;&nbsp;")
                    End If
                    If cMenu.Coll(i).IsLink Then
                        sbMenu.Append("<a href='javascript:" & cMenu.Coll(i).OnClickMethod & "()'>" & cMenu.Coll(i).Text & "</a>")
                    ElseIf cMenu.Coll(i).IsButton Then
                        sbMenu.Append("<input onclick='" & cMenu.Coll(i).OnClickMethod & "()' type='button' value='" & cMenu.Coll(i).Text & "'>")
                    End If
                    ItemAdded = True
                End If
            Next
        End If


        'sb.Append("<tr height='30px' class='dgh'>")
        ''   sb.Append("<tr class='dgh'>")
        'If FilterExists Or MenuExists Or DisplayNewButton Then
        '    sb.Append("<td align='left' colspan='" & ColCount - 1 & "'>" & sbMenu.ToString & "</td>")
        'Else
        '    sb.Append("<td align='left' colspan='" & ColCount - 1 & "'>&nbsp;</td>")
        'End If
        'sb.Append("</tr>")

        sb.Append("<tr height='20px' class='dgh'>")
        If FilterExists Or MenuExists Or DisplayNewButton Then
            sb.Append("<td align='left' colspan='" & ColCount - 1 & "'>" & sbMenu.ToString & "</td>")
        Else
            sb.Append("<td align='left' colspan='" & ColCount - 1 & "'>&nbsp;</td>")
        End If
        sb.Append("</tr>")

    End Sub

    Private Sub AddHeaderRow(ByRef sb As System.Text.StringBuilder)
        Dim ColNum As Integer
        Dim DataFldName As String
        Dim HeaderText As String = String.Empty
        Dim Width As String
        'Dim LeftMost As Boolean

        'LeftMost = True
        sb.Append("<tr class=""dgh"">")

        For ColNum = 1 To cColumnColl.Count

            If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then

                If cColumnColl(ColNum).Visible Then
                    If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Template Then
                        DataFldName = cColumnColl(ColNum).DataFldName
                        If cColumnColl(ColNum).HeaderText = Nothing Then
                            HeaderText = "&nbsp;"
                        Else
                            HeaderText = cColumnColl(ColNum).HeaderText
                        End If

                        '' ___ td start tag
                        'sb.Append("<td " & cColumnColl(ColNum).Attributes)
                        ''If LeftMost Then
                        ''    sb.Append(cLeftPadding)
                        ''    LeftMost = False
                        ''End If
                        'sb.Append(">")

                        '' ___ Value
                        'If cColumnColl(ColNum).SortExpression = Nothing Then
                        '    sb.Append(HeaderText)
                        'Else
                        '    sb.Append("<a href=""javascript:Sort('" & cColumnColl(ColNum).SortExpression & "')"">" & HeaderText & "</a>")
                        'End If

                        ' ___ td end tag
                        'sb.Append("</td>")

                        If cColumnColl(ColNum).SortExpression = Nothing Then
                            sb.Append("<td " & cColumnColl(ColNum).Attributes & ">" & HeaderText & "</td>")
                        Else
                            sb.Append("<td " & cColumnColl(ColNum).Attributes & "><a href=""javascript:Sort('" & cColumnColl(ColNum).SortExpression & "')"">" & HeaderText & "</a></td>")
                        End If

                    Else
                        HeaderText = cColumnColl(ColNum).HeaderText
                        sb.Append("<td>" & HeaderText & "</td>")

                        'If LeftMost Then
                        '    sb.Append("<td" & cLeftPadding & ">" & HeaderText & "</td>")
                        '    LeftMost = False
                        'Else
                        '    sb.Append("<td>" & HeaderText & "</td>")
                        'End If

                    End If
                End If

            End If

        Next

        sb.Append("</tr>")

    End Sub

    Private Sub AddFilter(ByRef sb As System.Text.StringBuilder, ByVal Section As String)
        Dim i, j As Integer
        Dim ColNum As Integer
        Dim ItemName As String
        Dim CtlName As String
        Dim Width As String
        Dim ColCount As Integer
        Dim ShowHideFilter As String
        Dim ColumnHasFilter As Boolean
        Dim DropdownColl As Collection
        Dim SelectedValue As String
        Dim DisplayNewButton As Boolean
        Dim HeaderLinkText As String

        ' ___ Are there any filter items?
        If cFilter.Coll.count = 0 Then
            Exit Sub
        End If

        ' ___ Get the column count
        For ColNum = 1 To cColumnColl.Count
            If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then
                If cColumnColl(ColNum).Visible Then
                    If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Template Then
                        ColCount += 1
                    End If
                End If
            End If
        Next


        If Section = "header" Then

            If cShowFilter Then
                ShowHideFilter = "Hide Filter"
            Else
                ShowHideFilter = "Show Filter"
            End If

            'sb.Append("<tr class=""dgh"">")
            'If ColCount = 1 Then
            '    sb.Append("<td align='left'><a href=""javascript:ToggleShowFilter()"">" & ShowHideFilter & "</a>")
            '    sb.Append("&nbsp;&nbsp;")
            '    sb.Append("<a href=""javascript:ApplyFilter()""> Go </a></td>")
            'ElseIf ColCount = 2 Then
            '    sb.Append("<td align='left'><a href=""javascript:ToggleShowFilter()"">" & ShowHideFilter & "</a></td>")
            '    sb.Append("<td><a href=""javascript:ApplyFilter()""> Go </a></td>")
            'Else
            '    sb.Append("<td align='left' colspan='" & ColCount - 1 & "'><a href=""javascript:ToggleShowFilter()"">" & ShowHideFilter & "</a></td>")
            '    sb.Append("<td><a href=""javascript:ApplyFilter()""> Go </a></td>")
            'End If
            'sb.Append("</tr>")
            'sb.Append("<tr height='10px'><td>&nbsp;</td></tr>")


            If cAttachNewButton Then
                If cNewButtonRight = Nothing Then
                    DisplayNewButton = True
                Else
                    If Rights.HasThisRight(cNewButtonRight) Then
                        DisplayNewButton = True
                    End If
                End If
            End If

            HeaderLinkText = "<a href=""javascript:ApplyFilter()"">Go</a>&nbsp;&nbsp;&nbsp;<a href=""javascript:ToggleShowFilter()"">" & ShowHideFilter & "</a>"
            sb.Append("<tr height='30px' class='dgh'>")
            If DisplayNewButton Then
                sb.Append("<td align='left' colspan='" & ColCount - 1 & "'>" & HeaderLinkText & "&nbsp;&nbsp;&nbsp<a href='javascript:NewRecord()'>New</a></td>")
            Else
                sb.Append("<td align='left'>" & HeaderLinkText & "</td>")
            End If
            sb.Append("</tr>")


            'sb.Append("<tr height='30px' class='dgh'>")
            'sb.Append("<td align='left'><a href=""javascript:ApplyFilter()"">Go</a>")
            'sb.Append("&nbsp;&nbsp;&nbsp;")
            'sb.Append("<a href=""javascript:ToggleShowFilter()"">" & ShowHideFilter & "</a></td>")
            'sb.Append("</tr>")

        ElseIf Section = "filter" Then

            If cShowFilter Then
                sb.Append("<tr class=""dgh"">")
                For ColNum = 1 To cColumnColl.Count
                    If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Hidden Then
                        If cColumnColl(ColNum).Visible Then
                            If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Template Then

                                ItemName = cColumnColl(ColNum).ItemName.toupper
                                ColumnHasFilter = False
                                For i = 1 To cFilter.Coll.count
                                    If cFilter.Coll(i).ItemName.toupper = ItemName Then
                                        ColumnHasFilter = True
                                        Exit For
                                    End If
                                Next

                                If ColumnHasFilter Then
                                    If cFilter.Coll(ItemName).IsTextbox Then
                                        sb.Append("<td " & cColumnColl(ColNum).Attributes & "><input style='FONT: 8pt Arial, Helvetica, sans-serif;' type='text' name='" & cFilter.Coll(ItemName).CtlName & "' id='" & cFilter.Coll(ItemName).CtlName & "' value='" & cFilter.Coll(ItemName).GetValue() & "' maxlength='" & cFilter.Coll(ItemName).MaxLength & "' onkeypress='return SubmitOnEnterKey(event)'></td>")
                                    ElseIf cFilter.Coll(ItemName).IsDropdown Then
                                        DropdownColl = cFilter.Coll(ItemName).Coll
                                        sb.Append("<td " & cColumnColl(ColNum).Attributes & "><select style='FONT: 8pt Arial, Helvetica, sans-serif;' name='" & cFilter.Coll(ItemName).CtlName & "' id='" & cFilter.Coll(ItemName).CtlName & "' value='" & cFilter.Coll(ItemName).GetValue() & "'>")
                                        For j = 1 To DropdownColl.Count
                                            SelectedValue = cFilter.Coll(ItemName).GetValue()
                                            If DropdownColl(j).Value = SelectedValue Then
                                                sb.Append("<option selected value='" & DropdownColl(j).Value & "'>" & DropdownColl(j).Text & "</option>")
                                            Else
                                                sb.Append("<option value='" & DropdownColl(j).Value & "'>" & DropdownColl(j).Text & "</option>")
                                            End If
                                        Next
                                        sb.Append("</select>")
                                    End If

                                Else
                                    sb.Append("<td>&nbsp;</td>")
                                End If

                            End If
                        End If
                    End If
                Next
                sb.Append("</tr>")
            End If
        End If
        sb.Append("<tr class=""dgh"">")
    End Sub

    'Private Sub AddDataRow(ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByVal OddRow As Boolean, ByRef sb As System.Text.StringBuilder)
    '    Dim ColNum As Integer

    '    If OddRow Then
    '        sb.Append("<tr class=""DGOdd"">" & vbCrLf)
    '    Else
    '        sb.Append("<tr class=""DGEven"">" & vbCrLf)
    '    End If

    '    For ColNum = 1 To cColumnColl.Count

    '        If cColumnColl(ColNum).ColumnType = DG.ColumnType.Hidden Then
    '            HandleNonTemplateColumns(ColNum, dt, RowNum, sb)
    '        Else
    '            If cColumnColl(ColNum).Visible Then
    '                If cColumnColl(ColNum).ColumnType <> DG.ColumnType.Template Then
    '                    HandleNonTemplateColumns(ColNum, dt, RowNum, sb)
    '                Else
    '                    HandleTemplateColumn(ColNum, dt, RowNum, sb)
    '                End If
    '            End If
    '        End If

    '    Next
    '    sb.Append("</tr>")
    'End Sub

    Private Sub AddDataRow(ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByVal OddRow As Boolean, ByRef sb As System.Text.StringBuilder)
        Dim ColNum As Integer

        If OddRow Then
            sb.Append("<tr class=""DGOdd"">" & vbCrLf)
        Else
            sb.Append("<tr class=""DGEven"">" & vbCrLf)
        End If

        For ColNum = 1 To cColumnColl.Count

            If cColumnColl(ColNum).ColumnType = DG.ColumnType.Hidden Then
                HandleHiddenColumn(ColNum, dt, RowNum, sb)
            Else
                If cColumnColl(ColNum).Visible Then
                    Select Case cColumnColl(ColNum).ColumnType
                        Case DG.ColumnType.Databound
                            HandleDataboundColumn(ColNum, dt, RowNum, sb)
                        Case DG.ColumnType.Link
                            HandleLinkColumn(ColNum, dt, RowNum, sb)
                        Case DG.ColumnType.CheckboxToggle
                            HandleCheckboxToggleColumn(ColNum, dt, RowNum, sb)
                        Case DG.ColumnType.Template
                            HandleTemplateColumn(ColNum, dt, RowNum, sb)
                        Case DG.ColumnType.ChildTableSelect
                            HandleChildTableSelectColumn(ColNum, dt, RowNum, sb)
                        Case DG.ColumnType.Boolean
                            HandleBooleanColumn(ColNum, dt, RowNum, sb)
                        Case DG.ColumnType.Date
                            HandleDateColumn(ColNum, dt, RowNum, sb)
                        Case DG.ColumnType.FreeForm
                            HandleFreeFormColumn(ColNum, dt, RowNum, sb)
                    End Select
                End If
            End If

        Next
        sb.Append("</tr>")
    End Sub

    Private Sub HandleCheckboxToggleColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String
        Dim ColumnAttributes As String
        Dim CheckboxLabel As String

        ' __ Handle column attributes and tooltip
        ColumnAttributes = cColumnColl(ColNum).Attributes
        If cColumnColl(ColNum).TitleFldName <> Nothing Then
            ' ColumnAttributes &= " title='" & Common.StrInHandler(objDataReader(cColumnColl(i).TitleFldName)) & "' "
            ColumnAttributes &= " title='" & Common.StrInHandler(dt.Rows(RowNum)(cColumnColl(ColNum).TitleFldName)) & "' "
        End If

        If dt.Rows(RowNum)(cColumnColl(ColNum).TestFld) = "1" Then
            CheckboxLabel = cColumnColl(ColNum).TrueText
        Else
            CheckboxLabel = cColumnColl(ColNum).FalseText
        End If

        Value = "<input type='checkbox' name='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName) & "' value='on'>&nbsp;" & CheckboxLabel
        'Value &= "<input type='hidden' name='hd" & cColumnColl(i).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName) & " value='" & dt.Rows(RowNum)(cColumnColl(i).DataFldName) & "'>"

        ' ___ This is provided as an alternative way of iterating through the checkbox items.
        cCheckboxToggleColl(cColumnColl(ColNum).ItemName).append(dt.Rows(RowNum)(cKeyFieldName) & "~" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "|")
        sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")
    End Sub

    Private Sub HandleLinkColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String
        Dim ColumnAttributes As String
        Dim Box As Object
        Dim DisplayLinkAsLink As Boolean
        Dim DisplayLinkAsText As Boolean
        Dim DataFldName As String

        ' __ Handle column attributes and tooltip
        ColumnAttributes = cColumnColl(ColNum).Attributes
        If cColumnColl(ColNum).TitleFldName <> Nothing Then
            ColumnAttributes &= " title='" & Common.StrInHandler(dt.Rows(RowNum)(cColumnColl(ColNum).TitleFldName)) & "' "
        End If

        ' HRef = "javascript:fnToggle"
        ' DataFldName "ReadyForPrint|1"       Display as link
        ' DataFldName "ReadyForPrint|0"       Display as text
        ' DataFldName "ReadyForPrint"          Display as link

        If InStr(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName), "|") = 0 Then
            DisplayLinkAsText = True
            DisplayLinkAsLink = False
            DataFldName = cColumnColl(ColNum).DataFldName
        Else
            Box = Split(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName), "|")
            If Box(1) = "0" Then
                DisplayLinkAsText = True
                DisplayLinkAsLink = False
                DataFldName = Box(0)
            Else
                DisplayLinkAsText = False
                DisplayLinkAsLink = True
                DataFldName = Box(0)
            End If
        End If

        If DisplayLinkAsLink Then
            Value = cColumnColl(ColNum).Href
            ' Value = "<a href=""""" & Value & "('" & objDataReader(cKeyFieldName) & "')"""">" & Common.StrInHandler(objDataReader(DataFldName)) & "</a>"
            Value = "<a href=""""" & Value & "('" & dt.Rows(RowNum)(cKeyFieldName) & "')"""">" & Common.StrInHandler(dt.Rows(RowNum)(DataFldName)) & "</a>"
        ElseIf DisplayLinkAsText Then
            'Value = Common.StrInHandler(objDataReader(DataFldName))
            Value = Common.StrInHandler(dt.Rows(RowNum)(DataFldName))
            If Value.Length = 0 Then
                Value = "&nbsp;"
            End If
        End If
        sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")
    End Sub

    Private Sub HandleBooleanColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String
        Dim ColumnAttributes As String

        ' __ Handle column attributes and tooltip
        ColumnAttributes = cColumnColl(ColNum).Attributes
        If cColumnColl(ColNum).Title <> Nothing Then
            ColumnAttributes &= " title='" & Common.StrInHandler(cColumnColl(ColNum).Title) & "' "
        End If

        ' ___ Determine the value
        If Common.StrInHandler(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName)) = cColumnColl(ColNum).TrueValue Then
            Value = cColumnColl(ColNum).TrueText
        Else
            Value = cColumnColl(ColNum).FalseText
        End If

        sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")
    End Sub

    Private Sub HandleDateColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim OrigValue As DateTime
        Dim Value As String
        Dim ColumnAttributes As String

        ' __ Handle column attributes and tooltip
        ColumnAttributes = cColumnColl(ColNum).Attributes
        If cColumnColl(ColNum).Title <> Nothing Then
            ColumnAttributes &= " title='" & Common.StrInHandler(cColumnColl(ColNum).Title) & "' "
        End If

        ' ___ Get and format value
        Value = Common.DateInHandler(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName))
        If Value = String.Empty Then
            Value = "&nbsp;"
        Else
            OrigValue = dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName)
            If cColumnColl(ColNum).DataFormatString = Nothing Then
                Value = OrigValue.ToString
            Else
                Value = OrigValue.ToString(cColumnColl(ColNum).DataFormatString)
            End If
        End If


        'If IsDBNull(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName)) Then
        '    Value = "&nbsp;"
        'Else
        '    OrigValue = dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName)
        '    If cColumnColl(ColNum).DataFormatString = Nothing Then
        '        Value = OrigValue.ToString
        '    Else
        '        Value = OrigValue.ToString(cColumnColl(ColNum).DataFormatString)
        '    End If
        'End If

        sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")
    End Sub


    Private Sub HandleFreeFormColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String
        Dim ColumnAttributes As String

        ' __ Handle column attributes and tooltip
        ColumnAttributes = cColumnColl(ColNum).Attributes
        If cColumnColl(ColNum).Title <> Nothing Then
            ColumnAttributes &= " title='" & Common.StrInHandler(dt.Rows(RowNum)(cColumnColl(ColNum).Title)) & "' "
        End If

        ' ___ Get the value
        Value = Common.StrInHandler(cColumnColl(ColNum).CellText)

        ' ___ Format value
        If Value = Nothing Then
            Value = "&nbsp;"
        End If

        sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")
    End Sub

    Private Sub HandleDataboundColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String
        Dim ColumnAttributes As String

        ' __ Handle column attributes and tooltip
        ColumnAttributes = cColumnColl(ColNum).Attributes
        If cColumnColl(ColNum).TitleFldName <> Nothing Then
            ColumnAttributes &= " title='" & Common.StrInHandler(dt.Rows(RowNum)(cColumnColl(ColNum).TitleFldName)) & "' "
        End If

        ' ___ Get the value
        Value = Common.StrInHandler(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName))

        '' ___ Handle boolean
        'If dt.Columns(cColumnColl(ColNum).DataFldName).datatype.ToString.ToLower = "system.boolean" Then
        '    If IsDBNull(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName)) Then
        '        Value = "No"
        '    Else
        '        If dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) Then
        '            Value = "Yes"
        '        Else
        '            Value = "No"
        '        End If
        '    End If
        'End If

        ' ___ Format value
        If Value.Length = 0 Then
            Value = "&nbsp;"
        Else
            If Not cColumnColl(ColNum).DataFormatString = Nothing Then
                Value = CDbl(Value).ToString(cColumnColl(ColNum).DataFormatString)
            End If
        End If

        sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")
    End Sub

    Private Sub HandleHiddenColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String

        Value = "<input type='hidden' id='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName) & "' name='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName) & "' value='" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "'>"
        cHiddenColumnColl(cColumnColl(ColNum).ItemName).add(dt.Rows(RowNum)(cKeyFieldName) & "~" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName))
        sb.Append(Value)
    End Sub


    Private Sub HandleChildTableSelectColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String
        Dim ColumnAttributes As String

        ' __ Handle column attributes and tooltip
        ColumnAttributes = "width= '20px' "
        ColumnAttributes &= cColumnColl(ColNum).Attributes

        If dt.Rows(RowNum)(cColumnColl(ColNum).EnableFldName) = 0 Then
            Value = "&nbsp;"
        Else
            If cColumnColl(ColNum).Title <> Nothing Then
                ColumnAttributes &= " title='" & cColumnColl(ColNum).Title & "' "
            End If


            '    If dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) = cColumnColl(ColNum).ParmColl("DataFldName").Value Then
            '        Value = "<a href=""javascript:ShowHideSubTable('" & cColumnColl(ColNum).ParmColl("DataFldName").FldName & "', '')""><strong>-</strong></a>"
            '    Else
            '        Select Case cColumnColl(ColNum).ParmColl.Count
            '            Case 1
            '                Value = "<a href=""javascript:ShowHideSubTable('" & cColumnColl(ColNum).ParmColl("DataFldName").FldName & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "')""><strong>+</strong></a>"
            '            Case 2
            '                Value = "<a href=""javascript:ShowHideSubTable('" & cColumnColl(ColNum).ParmColl("DataFldName").FldName & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm2").FldName) & "')""><strong>+</strong></a>"
            '            Case 3
            '                Value = "<a href=""javascript:ShowHideSubTable('" & cColumnColl(ColNum).ParmColl("DataFldName").FldName & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm2").FldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm3").FldName) & "')""><strong>+</strong></a>"
            '            Case 4
            '                Value = "<a href=""javascript:ShowHideSubTable('" & cColumnColl(ColNum).ParmColl("DataFldName").FldName & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm2").FldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm3").FldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm4").FldName) & "')""><strong>+</strong></a>"
            '        End Select
            '    End If
            'End If


            If dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) = cColumnColl(ColNum).ParmColl("DataFldName").Value Then
                Value = "<a href=""javascript:ShowHideSubTable('0', '')""><strong>-</strong></a>"
            Else
                Select Case cColumnColl(ColNum).ParmColl.Count
                    Case 1
                        Value = "<a href=""javascript:ShowHideSubTable('1', '" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "')""><strong>+</strong></a>"
                    Case 2
                        Value = "<a href=""javascript:ShowHideSubTable('1', '" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm2").FldName) & "')""><strong>+</strong></a>"
                    Case 3
                        Value = "<a href=""javascript:ShowHideSubTable('1', '" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm2").FldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm3").FldName) & "')""><strong>+</strong></a>"
                    Case 4
                        Value = "<a href=""javascript:ShowHideSubTable('1', '" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm2").FldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm3").FldName) & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).ParmColl("Parm4").FldName) & "')""><strong>+</strong></a>"
                End Select
            End If
        End If

        sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")
    End Sub

    Private Sub oldHandleChildTableSelectColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim Value As String
        Dim ColumnAttributes As String

        ' __ Handle column attributes and tooltip
        ColumnAttributes = "width= '20px' "
        ColumnAttributes &= cColumnColl(ColNum).Attributes

        If dt.Rows(RowNum)(cColumnColl(ColNum).EnableFldName) = 0 Then
            Value = "&nbsp;"
        Else
            If cColumnColl(ColNum).Title <> Nothing Then
                ColumnAttributes &= " title='" & cColumnColl(ColNum).Title & "' "
            End If
            'If dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) = cChildTables.Value Then
            '    ' Value = "<a href=""javascript:ShowHideSubTable('" & cColumnColl(ColNum).ParmColl("DataFldName") & "', '')""><strong>-</strong></a>"
            '    Value = "<a href=""javascript:ShowHideSubTable('" & cColumnColl(ColNum).DataFldName & "', '')""><strong>-</strong></a>"
            'Else
            '    'Value = "<a href=""javascript:ShowHideSubTable('" & cColumnColl(ColNum).DataFldName & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "')""><strong>+</strong></a>"
            '    Value = "<a href=""javascript:ShowHideSubTable('" & cColumnColl(ColNum).ParmColl("DataFldName") & "', '" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "')""><strong>+</strong></a>"
            'End If
        End If
        sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")
    End Sub

    'Private Sub HandleNonTemplateColumns(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
    '    Dim Value As String
    '    Dim ColumnAttributes As String
    '    Dim Box As Object
    '    Dim DisplayLinkAsLink As Boolean
    '    Dim DisplayLinkAsText As Boolean
    '    Dim DataFldName As String
    '    Dim CheckboxLabel As String

    '    ' __ Handle column attributes and tooltip
    '    ColumnAttributes = cColumnColl(ColNum).Attributes
    '    If cColumnColl(ColNum).TitleFldName <> Nothing Then
    '        ' ColumnAttributes &= " title='" & Common.StrInHandler(objDataReader(cColumnColl(i).TitleFldName)) & "' "
    '        ColumnAttributes &= " title='" & Common.StrInHandler(dt.Rows(RowNum)(cColumnColl(ColNum).TitleFldName)) & "' "
    '    End If

    '    If cColumnColl(ColNum).ColumnType = DG.ColumnType.Databound Then

    '        'Value = Common.StrInHandler(objDataReader(cColumnColl(i).DataFldName))
    '        Value = Common.StrInHandler(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName))

    '        ' ___ Handle boolean
    '        If dt.Columns(cColumnColl(ColNum).DataFldName).datatype.ToString.ToLower = "system.boolean" Then
    '            If IsDBNull(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName)) Then
    '                Value = "No"
    '            Else
    '                If dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) Then
    '                    Value = "Yes"
    '                Else
    '                    Value = "No"
    '                End If
    '            End If
    '        End If

    '        ' ___ Format value
    '        If Value.Length = 0 Then
    '            Value = "&nbsp;"
    '        Else
    '            If Not cColumnColl(ColNum).DataFormatString = Nothing Then
    '                Value = CDbl(Value).ToString(cColumnColl(ColNum).DataFormatString)
    '            End If
    '        End If

    '    ElseIf cColumnColl(ColNum).columntype = DG.ColumnType.CheckboxToggle Then
    '        If dt.Rows(RowNum)(cColumnColl(ColNum).TestFld) = "1" Then
    '            CheckboxLabel = cColumnColl(ColNum).TrueText
    '        Else
    '            CheckboxLabel = cColumnColl(ColNum).FalseText
    '        End If

    '        Value = "<input type='checkbox' name='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName) & "' value='on'>&nbsp;" & CheckboxLabel
    '        'Value &= "<input type='hidden' name='hd" & cColumnColl(i).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName) & " value='" & dt.Rows(RowNum)(cColumnColl(i).DataFldName) & "'>"

    '        ' ___ This is provided as an alternative way of iterating through the checkbox items.
    '        cCheckboxToggleColl(cColumnColl(ColNum).ItemName).append(dt.Rows(RowNum)(cKeyFieldName) & "~" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "|")

    '    ElseIf cColumnColl(ColNum).columntype = DG.ColumnType.Hidden Then
    '        Value = "<input type='hidden' id='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName) & "' name='" & cColumnColl(ColNum).ItemName & "|" & dt.Rows(RowNum)(cKeyFieldName) & "' value='" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName) & "'>"
    '        cHiddenColumnColl(cColumnColl(ColNum).ItemName).add(dt.Rows(RowNum)(cKeyFieldName) & "~" & dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName))

    '    ElseIf cColumnColl(ColNum).columntype = DG.ColumnType.Link Then

    '        ' ___ Link
    '        ' HRef = "javascript:fnToggle"
    '        ' DataFldName "ReadyForPrint|1"       Display as link
    '        ' DataFldName "ReadyForPrint|0"       Display as text
    '        ' DataFldName "ReadyForPrint"          Display as link

    '        'If InStr(objDataReader(cColumnColl(i).DataFldName), "|") = 0 Then
    '        If InStr(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName), "|") = 0 Then
    '            DisplayLinkAsText = True
    '            DisplayLinkAsLink = False
    '            DataFldName = cColumnColl(ColNum).DataFldName
    '        Else
    '            'Box = Split(objDataReader(cColumnColl(i).DataFldName), "|")
    '            Box = Split(dt.Rows(RowNum)(cColumnColl(ColNum).DataFldName), "|")
    '            If Box(1) = "0" Then
    '                DisplayLinkAsText = True
    '                DisplayLinkAsLink = False
    '                DataFldName = Box(0)
    '            Else
    '                DisplayLinkAsText = False
    '                DisplayLinkAsLink = True
    '                DataFldName = Box(0)
    '            End If
    '        End If

    '        If DisplayLinkAsLink Then
    '            Value = cColumnColl(ColNum).Href
    '            ' Value = "<a href=""""" & Value & "('" & objDataReader(cKeyFieldName) & "')"""">" & Common.StrInHandler(objDataReader(DataFldName)) & "</a>"
    '            Value = "<a href=""""" & Value & "('" & dt.Rows(RowNum)(cKeyFieldName) & "')"""">" & Common.StrInHandler(dt.Rows(RowNum)(DataFldName)) & "</a>"
    '        ElseIf DisplayLinkAsText Then
    '            'Value = Common.StrInHandler(objDataReader(DataFldName))
    '            Value = Common.StrInHandler(dt.Rows(RowNum)(DataFldName))
    '            If Value.Length = 0 Then
    '                Value = "&nbsp;"
    '            End If
    '        End If

    '    End If

    '    If cColumnColl(ColNum).columntype = DG.ColumnType.Hidden Then
    '        sb.Append(Value)
    '    Else
    '        sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")
    '        'If LeftMost Then
    '        '    sb.Append("<td " & cLeftPadding & ColumnAttributes & ">" & Value & "</td>")
    '        '    LeftMost = False
    '        'Else
    '        '    sb.Append("<td " & ColumnAttributes & ">" & Value & "</td>")
    '        'End If
    '    End If
    'End Sub

    Private Sub HandleTemplateColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
        Dim TemplateCol As DG.TemplateColumn
        Dim IsVisible As Boolean
        Dim j As Integer
        Dim PassPermissionFldTest As Boolean
        Dim PassRightTest As Boolean

        TemplateCol = cColumnColl(ColNum)

        ' __ Handle column attributes and tooltip
        sb.Append("<td ")
        If cColumnColl(ColNum).Attributes <> Nothing Then
            sb.Append(cColumnColl(ColNum).Attributes)
        End If
        If Not TemplateCol.Wrap Then
            sb.Append(" nowrap ")
        End If
        sb.Append(">")

        For j = 1 To TemplateCol.Count

            IsVisible = False
            PassPermissionFldTest = False
            PassRightTest = False

            ' ___ Permission field test
            If TemplateCol(j).PermissionFldName = Nothing Then
                PassPermissionFldTest = True
            Else
                If (Not IsDBNull(dt.Rows(RowNum)(TemplateCol(j).PermissionFldName))) AndAlso (dt.Rows(RowNum)(TemplateCol(j).PermissionFldName) = "1") Then
                    PassPermissionFldTest = True
                End If
            End If

            ' ___ Rights test
            If TemplateCol(j).Right = Nothing Then
                PassRightTest = True
            Else
                If Rights.HasThisRight(TemplateCol(j).Right) Then
                    PassRightTest = True
                End If
            End If

            If PassPermissionFldTest AndAlso PassRightTest Then
                IsVisible = True
            End If

            If IsVisible Then
                sb.Append("<a ")
                sb.Append("id=""" & TemplateCol(j).DataFldName & """ ")

                If TemplateCol(j).Parm2 = Nothing Then
                    sb.Append("onclick=""" & TemplateCol(j).OnclickMethod & "('" & dt.Rows(RowNum)(cKeyFieldName) & "')"">")
                ElseIf TemplateCol(j).Parm3 = Nothing Then
                    sb.Append("onclick=""" & TemplateCol(j).OnclickMethod & "('" & dt.Rows(RowNum)(cKeyFieldName) & "', '" & Common.StrInHandler(dt.Rows(RowNum)(TemplateCol(j).Parm2)) & "')"">")
                ElseIf TemplateCol(j).Parm4 = Nothing Then
                    sb.Append("onclick=""" & TemplateCol(j).OnclickMethod & "('" & dt.Rows(RowNum)(cKeyFieldName) & "', '" & Common.StrInHandler(dt.Rows(RowNum)(TemplateCol(j).Parm2)) & "', '" & Common.StrInHandler(dt.Rows(RowNum)(TemplateCol(j).Parm3)) & "')"">")
                Else
                    sb.Append("onclick=""" & TemplateCol(j).OnclickMethod & "('" & dt.Rows(RowNum)(cKeyFieldName) & "', '" & Common.StrInHandler(dt.Rows(RowNum)(TemplateCol(j).Parm2)) & "', '" & "', '" & Common.StrInHandler(dt.Rows(RowNum)(TemplateCol(j).Parm3)) & "', '" & Common.StrInHandler(dt.Rows(RowNum)(TemplateCol(j).Parm4)) & "')"">")
                End If

                If TemplateCol(j).UseDefaultImage Then
                    sb.Append(cColumnColl(ColNum).DefaultImageColl(TemplateCol(j).DefaultImage))
                Else
                    sb.Append("<img src=""" & TemplateCol(j).ImagePath & """ ")
                    sb.Append(TemplateCol(j).ImageAttributes)
                End If

                sb.Append("title=""" & TemplateCol(j).Title & """> ")
                sb.Append("</a>")
            End If
        Next
        sb.Append("</td>")
    End Sub

    'Private Sub HandleTemplateColumn(ByVal ColNum As Integer, ByRef dt As Data.DataTable, ByVal RowNum As Integer, ByRef sb As System.Text.StringBuilder)
    '    Dim TemplateCol As DG.TemplateColumn
    '    Dim IsVisible As Boolean
    '    Dim j As Integer
    '    Dim PassPermissionFldTest As Boolean
    '    Dim PassRightTest As Boolean
    '    Dim ColumnAttributes As String = String.Empty

    '    '            cDefaultImageColl.Add("class='MainMenuTbl><a href='AboutUs.aspx'>+</a>", "PlusSignView")
    '    '		<td class="MainMenuTbl"><a href="AboutUs.aspx">+ and a dog</a></td>
    '    '		<td class="MainMenuTbl"><a href="javascript:fnTestie()">Testie</a></td>

    '    TemplateCol = cColumnColl(ColNum)

    '    'If TemplateCol.Wrap Then
    '    '    sb.Append("<td>")
    '    'Else
    '    '    sb.Append("<td nowrap>")
    '    'End If

    '    ' __ Handle column attributes and tooltip
    '    ColumnAttributes = "<td "
    '    If cColumnColl(ColNum).Attributes <> Nothing Then
    '        ColumnAttributes &= cColumnColl(ColNum).Attributes
    '    End If
    '    'If TemplateCol.Ti <> Nothing Then
    '    '    ' ColumnAttributes &= " title='" & Common.StrInHandler(objDataReader(cColumnColl(i).TitleFldName)) & "' "
    '    '    ColumnAttributes &= " title='" & Common.StrInHandler(dt.Rows(RowNum)(cColumnColl(ColNum).TitleFldName)) & "' "
    '    'End If
    '    If Not TemplateCol.Wrap Then
    '        ColumnAttributes &= " nowrap "
    '    End If
    '    ColumnAttributes &= ">"


    '    'If LeftMost Then
    '    '    Attributes = cLeftPadding
    '    '    LeftMost = False
    '    'End If
    '    'If Not TemplateCol.Wrap Then
    '    '    Attributes &= " nowrap "
    '    'End If
    '    'sb.Append("<td" & Attributes & ">")


    '    For j = 1 To TemplateCol.Count

    '        IsVisible = False
    '        PassPermissionFldTest = False
    '        PassRightTest = False

    '        ' ___ Permission field test
    '        If TemplateCol(j).PermissionFldName = Nothing Then
    '            PassPermissionFldTest = True
    '        Else
    '            If (Not IsDBNull(dt.Rows(RowNum)(TemplateCol(j).PermissionFldName))) AndAlso (dt.Rows(RowNum)(TemplateCol(j).PermissionFldName) = "1") Then
    '                PassPermissionFldTest = True
    '            End If
    '        End If

    '        ' ___ Rights test
    '        If TemplateCol(j).Right = Nothing Then
    '            PassRightTest = True
    '        Else
    '            If Rights.HasThisRight(TemplateCol(j).Right) Then
    '                PassRightTest = True
    '            End If
    '        End If

    '        If PassPermissionFldTest AndAlso PassRightTest Then
    '            IsVisible = True
    '        End If

    '        If IsVisible Then

    '            sb.Append("<a ")
    '            '  sb.Append("id=""" & TemplateCol(j).DataFldName & """ ")
    '            sb.Append("id='" & TemplateCol(j).DataFldName & "' ")

    '            If cColumnColl(ColNum).ColumnUsage = TemplateColumn.ColumnUsageEnum.Icon Then
    '                If TemplateCol(j).Parm2 = Nothing Then
    '                    sb.Append("onclick=""" & TemplateCol(j).OnclickMethod & "('" & dt.Rows(RowNum)(cKeyFieldName) & "')"">")
    '                Else
    '                    sb.Append("onclick=""" & TemplateCol(j).OnclickMethod & "('" & dt.Rows(RowNum)(cKeyFieldName) & "', '" & Common.StrInHandler(dt.Rows(RowNum)(TemplateCol(j).Parm2)) & "')"">")
    '                End If
    '                If TemplateCol(j).UseDefaultImage Then
    '                    sb.Append(cColumnColl(ColNum).DefaultImageColl(TemplateCol(j).DefaultImage))
    '                Else
    '                    sb.Append("<img src=""" & TemplateCol(j).ImagePath & """ ")
    '                    sb.Append(TemplateCol(j).ImageAttributes)
    '                End If

    '            ElseIf cColumnColl(ColNum).ColumnUsage = TemplateColumn.ColumnUsageEnum.CellAsLink Then
    '                If TemplateCol(j).Parm2 = Nothing Then
    '                    sb.Append("href=""javascript:" & TemplateCol(j).OnClickMethod & "('" & dt.Rows(RowNum)(cKeyFieldName) & "')"">")
    '                Else
    '                    sb.Append("href=""javascript:" & TemplateCol(j).OnClickMethod & "('" & dt.Rows(RowNum)(cKeyFieldName) & "', '" & Common.StrInHandler(dt.Rows(RowNum)(TemplateCol(j).Parm2)) & "')"">")
    '                End If
    '            End If

    '            sb.Append("title=""" & TemplateCol(j).Title & """> ")
    '            sb.Append("</a>")
    '        End If
    '    Next
    '    sb.Append("</td>")
    'End Sub
#End Region

#Region " Column classes "
    Public Class FreeFormColumn
        Inherits DataBoundColumnItems
        Private cCellText As String
        Private cTitle As String
        Public ReadOnly Property Title()
            Get
                Return cTitle
            End Get
        End Property

        Public ReadOnly Property CellText()
            Get
                Return cCellText
            End Get
        End Property
        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal CellText As String, ByVal HeaderText As String, ByVal Title As String, ByVal Visible As Boolean, ByVal Attributes As String)
            MyBase.New(ColumnType, ItemName, Nothing, HeaderText, Nothing, Visible, Nothing, Nothing, Attributes)
            cCellText = CellText
            cTitle = Title
        End Sub
    End Class

    Public Class ChildTableSelectColumn
        Inherits DataBoundColumnItems
        Private cTitle As String
        Private cEnableFldName As String
        Private cParmColl As Collection
        Private cChildTables As DG.ChildTablesClass

        Public ReadOnly Property ParmColl()
            Get
                Return cParmColl
            End Get
        End Property
        Public ReadOnly Property EnableFldName()
            Get
                Return cEnableFldName
            End Get
        End Property
        Public ReadOnly Property Title()
            Get
                Return cTitle
            End Get
        End Property
        Public Sub New(ByRef ChildTables As DG.ChildTablesClass, ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal EnableFldName As String, ByVal DataFormatString As String, ByVal Title As String, ByVal Attributes As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, Nothing, Attributes)
            cEnableFldName = EnableFldName
            cTitle = Title
            cChildTables = ChildTables
            cParmColl = New Collection
            cParmColl.Add(New Items("DataFldName", DataFldName), "DataFldName")
            If Not Parm2 = Nothing Then
                cParmColl.Add(New Items("Parm2", Parm2), "Parm2")
            End If
            If Not Parm3 = Nothing Then
                cParmColl.Add(New Items("Parm3", Parm3), "Parm3")
            End If
            If Not Parm4 = Nothing Then
                cParmColl.Add(New Items("Parm4", Parm4), "Parm4")
            End If
        End Sub

        Public Class Items
            Private cItemName As String
            Private cFldName As String
            Private cValue As String
            Public ReadOnly Property ItemName()
                Get
                    Return cItemName
                End Get
            End Property
            Public ReadOnly Property FldName()
                Get
                    Return cFldName
                End Get
            End Property
            Public Property Value()
                Get
                    Return cValue
                End Get
                Set(ByVal Value)
                    cValue = Value
                End Set
            End Property

            Public Sub New(ByVal ItemName As String, ByVal FldName As String)
                cItemName = ItemName
                cFldName = FldName
            End Sub
        End Class
    End Class

    Public Class TemplateColumn
        Private cColl As New Collection
        Private cHeaderText As String
        Private cWrap As Boolean
        Private cVisible As Boolean
        Private cDefaultImageColl As New Collection
        Private cItemName As String
        Private cAttributes As String

        Default Public ReadOnly Property Coll(ByVal Idx As Integer)
            Get
                Return cColl(Idx)
            End Get
        End Property
        Public ReadOnly Property Attributes() As String
            Get
                Return cAttributes
            End Get
        End Property
        Public ReadOnly Property ItemName()
            Get
                Return cItemName
            End Get
        End Property
        Public ReadOnly Property DefaultImageColl()
            Get
                Return cDefaultImageColl
            End Get
        End Property
        Public ReadOnly Property ColumnType()
            Get
                Return DG.ColumnType.Template
            End Get
        End Property
        Public ReadOnly Property HeaderText() As String
            Get
                Return cHeaderText
            End Get
        End Property
        Public ReadOnly Property Wrap()
            Get
                Return cWrap
            End Get
        End Property
        Public ReadOnly Property Visible()
            Get
                Return cVisible
            End Get
        End Property
        Public ReadOnly Property Count()
            Get
                Return cColl.Count
            End Get
        End Property

        Public Sub New(ByVal ItemName As String, ByVal HeaderText As String, ByVal Wrap As Boolean, ByVal Attributes As String, ByVal Visible As Boolean)
            cItemName = ItemName
            cHeaderText = HeaderText
            cWrap = Wrap
            cVisible = Visible
            cAttributes = Attributes
            ' cDefaultImageColl.Add("<img src='img/edit.ico'  border='0' width='14' height='14' vspace='0' hspace='0'", "StandardView")
            cDefaultImageColl.Add("<img src='img/view.gif'  border='0' width='14' height='14' vspace='0' hspace='0'", "StandardView")
            cDefaultImageColl.Add("<img src='img/delete.gif'  border='0' width='14' height='14' vspace='0' hspace='0'", "StandardDelete")
            cDefaultImageColl.Add("<img src='img/Clip.ico'  border='0' width='16' height='16' vspace='0' hspace='0'", "StandardClip")
            cDefaultImageColl.Add("<img src='img/BriefCse.ico'  border='0' width='16' height='16' vspace='0' hspace='0'", "StandardBriefcase")
            cDefaultImageColl.Add("<img src='img/Cardfil1.ico'  border='0' width='16' height='16' vspace='0' hspace='0'", "StandardCardfile")
            cDefaultImageColl.Add("<img src='img/23_6.png'  border='0' width='16' height='16' vspace='0' hspace='0'", "Standard23_6")
            cDefaultImageColl.Add("<img src='img/133.png'  border='0' width='16' height='16' vspace='0' hspace='0'", "Standard133")
            cDefaultImageColl.Add("<img src='img/Notes.gif'  border='0' width='16' height='16' vspace='0' hspace='0'", "StandardNotes")
            ' cDefaultImageColl.Add("<img src='genicon.gif'  border='0' width='16' height='16' vspace='0' hspace='0'", "StandardMisc")
        End Sub

        Public Sub AddTemplateItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal ImagePath As String, ByVal ImageAttributes As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String)
            cColl.Add(New TemplateColumnItems(ItemName, OnClickMethod, ImagePath, ImageAttributes, Title, Right, PermissionFldName, Nothing, Nothing, Nothing), ItemName)
        End Sub

        Public Sub AddTemplateItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal ImagePath As String, ByVal ImageAttributes As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String)
            cColl.Add(New TemplateColumnItems(ItemName, OnClickMethod, ImagePath, ImageAttributes, Title, Right, PermissionFldName, Parm2, Nothing, Nothing), ItemName)
        End Sub

        Public Sub AddTemplateItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal ImagePath As String, ByVal ImageAttributes As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String)
            cColl.Add(New TemplateColumnItems(ItemName, OnClickMethod, ImagePath, ImageAttributes, Title, Right, PermissionFldName, Parm2, Parm3, Nothing), ItemName)
        End Sub

        Public Sub AddTemplateItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal ImagePath As String, ByVal ImageAttributes As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
            cColl.Add(New TemplateColumnItems(ItemName, OnClickMethod, ImagePath, ImageAttributes, Title, Right, PermissionFldName, Parm2, Parm3, Parm4), ItemName)
        End Sub

        Public Sub AddDefaultTemplateItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal DefaultImage As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String)
            cColl.Add(New TemplateColumnItems(ItemName, OnClickMethod, DefaultImage, Title, Right, PermissionFldName, Nothing, Nothing, Nothing), ItemName)
        End Sub

        Public Sub AddDefaultTemplateItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal DefaultImage As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String)
            cColl.Add(New TemplateColumnItems(ItemName, OnClickMethod, DefaultImage, Title, Right, PermissionFldName, Parm2, Nothing, Nothing), ItemName)
        End Sub

        Public Sub AddDefaultTemplateItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal DefaultImage As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String)
            cColl.Add(New TemplateColumnItems(ItemName, OnClickMethod, DefaultImage, Title, Right, PermissionFldName, Parm2, Parm3, Nothing), ItemName)
        End Sub

        Public Sub AddDefaultTemplateItem(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal DefaultImage As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
            cColl.Add(New TemplateColumnItems(ItemName, OnClickMethod, DefaultImage, Title, Right, PermissionFldName, Parm2, Parm3, Parm4), ItemName)
        End Sub

        Public Class TemplateColumnItems
            Inherits DataBoundColumnItems
            Private cOnClickMethod As String
            Private cImagePath As String
            Private cImageAttributes As String
            Private cTitle As String
            Private cRight As String
            Private cPermissionFldName As String
            Private cParm2 As String
            Private cParm3 As String
            Private cParm4 As String
            Private cDefaultImage
            Private cUseDefaultImage As Boolean
            Private cIsImageType As Boolean
            Private cIsFreeForm As Boolean
            Private cCellText As String

            Public ReadOnly Property IsImageType()
                Get
                    Return cIsImageType
                End Get
            End Property
            Public ReadOnly Property IsFreeForm()
                Get
                    Return cIsFreeForm
                End Get
            End Property
            Public ReadOnly Property CellText()
                Get
                    Return cCellText
                End Get
            End Property
            Public ReadOnly Property DefaultImage()
                Get
                    Return cDefaultImage
                End Get
            End Property
            Public ReadOnly Property ImageAttributes()
                Get
                    Return cImageAttributes
                End Get
            End Property
            Public ReadOnly Property UseDefaultImage()
                Get
                    Return cUseDefaultImage
                End Get
            End Property
            Public ReadOnly Property OnClickMethod()
                Get
                    Return cOnClickMethod
                End Get
            End Property
            Public ReadOnly Property ImagePath()
                Get
                    Return cImagePath
                End Get
            End Property
            Public ReadOnly Property Title()
                Get
                    Return cTitle
                End Get
            End Property
            Public ReadOnly Property Right()
                Get
                    Return cRight
                End Get
            End Property
            Public ReadOnly Property PermissionFldName()
                Get
                    Return cPermissionFldName
                End Get
            End Property
            Public ReadOnly Property Parm2()
                Get
                    Return cParm2
                End Get
            End Property
            Public ReadOnly Property Parm3()
                Get
                    Return cParm3
                End Get
            End Property
            Public ReadOnly Property Parm4()
                Get
                    Return cParm4
                End Get
            End Property

            Public Sub New(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal ImagePath As String, ByVal ImageAttributes As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
                MyBase.New(DG.ColumnType.Template, ItemName, Nothing)
                cOnClickMethod = OnClickMethod
                cImagePath = ImagePath
                cImageAttributes = ImageAttributes
                cTitle = Title
                cRight = Right
                cPermissionFldName = PermissionFldName
                cParm2 = Parm2
                cParm3 = Parm3
                cParm4 = Parm4
                cUseDefaultImage = False
                cIsImageType = True
            End Sub

            Public Sub New(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal DefaultImage As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String)
                MyBase.New(DG.ColumnType.Template, ItemName, Nothing)
                cOnClickMethod = OnClickMethod
                cDefaultImage = DefaultImage
                cTitle = Title
                cRight = Right
                cPermissionFldName = PermissionFldName
                cParm2 = Parm2
                cParm3 = Parm3
                cParm4 = Parm4
                cUseDefaultImage = True
                cIsImageType = True
            End Sub

            Public Sub New(ByVal ItemName As String, ByVal OnClickMethod As String, ByVal CellText As String, ByVal Title As String, ByVal Right As String, ByVal PermissionFldName As String, ByVal Parm2 As String, ByVal Parm3 As String, ByVal Parm4 As String, ByVal IDoNothing1 As String, ByVal IDoNothing2 As String)
                MyBase.New(DG.ColumnType.Template, ItemName, Nothing)
                cOnClickMethod = OnClickMethod
                cDefaultImage = DefaultImage
                cTitle = Title
                cRight = Right
                cPermissionFldName = PermissionFldName
                cParm2 = Parm2
                cParm3 = Parm3
                cParm4 = Parm4
                cUseDefaultImage = True
                cIsFreeForm = True
                cCellText = CellText
            End Sub
        End Class

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class

    Public Class HiddenItems
        Inherits DataBoundColumnItems

        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String)
            MyBase.New(ColumnType, ItemName, DataFldName)
        End Sub
    End Class

    Public Class CheckboxToggleColumnItems
        Inherits DataBoundColumnItems
        Private cTrueText As String
        Private cFalseText As String
        Private cTestFld As String

        Public ReadOnly Property TestFld()
            Get
                Return cTestFld
            End Get
        End Property

        Public ReadOnly Property TrueText()
            Get
                Return cTrueText
            End Get
        End Property
        Public ReadOnly Property FalseText()
            Get
                Return cFalseText
            End Get
        End Property
        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal TitleFldName As String, ByVal Attributes As String, ByVal TestFld As String, ByVal TrueText As String, ByVal FalseText As String)
            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, Nothing, TitleFldName, Attributes)
            cTrueText = TrueText
            cFalseText = FalseText
            cTestFld = TestFld
        End Sub
    End Class

    Public Class LinkColumnItems
        Inherits DataBoundColumnItems
        Private cHRef As String
        Public ReadOnly Property HRef()
            Get
                Return cHRef
            End Get
        End Property
        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal HRef As String, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, TitleFldName, Attributes)
            cHRef = HRef
        End Sub
    End Class

    Public Class DateColumn
        Inherits DataBoundColumnItems
        Private cTitle As String

        Public ReadOnly Property Title()
            Get
                Return cTitle
            End Get
        End Property

        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal Title As String, ByVal Attributes As String)
            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, DataFormatString, Nothing, Attributes)
            cTitle = Title
        End Sub
    End Class

    Public Class BooleanColumn
        Inherits DataBoundColumnItems
        Private cTrueValue As String
        Private cTrueText As String
        Private cFalseText As String
        Private cTitle As String
        Public ReadOnly Property TrueText()
            Get
                Return cTrueText
            End Get
        End Property
        Public ReadOnly Property FalseText()
            Get
                Return cFalseText
            End Get
        End Property
        Public ReadOnly Property Truevalue()
            Get
                Return cTrueValue
            End Get
        End Property
        Public ReadOnly Property Title()
            Get
                Return cTitle
            End Get
        End Property
        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal TrueValue As String, ByVal TrueText As String, ByVal FalseText As String, ByVal Title As String, ByVal Attributes As String)
            MyBase.New(ColumnType, ItemName, DataFldName, HeaderText, SortExpression, Visible, Nothing, Nothing, Attributes)
            cTrueValue = TrueValue
            cTrueText = TrueText
            cFalseText = FalseText
            cTitle = Title
        End Sub
    End Class

    Public Class DataBoundColumnItems
        Private cColumnType As DG.ColumnType
        Private cItemName As String
        Private cDataFldName As String
        Private cHeaderText As String
        Private cSortExpression As String
        Private cVisible As Boolean
        Private cDataFormatString As String
        Private cTitleFldName As String
        Private cAttributes As String

        Public ReadOnly Property ColumnType()
            Get
                Return cColumnType
            End Get
        End Property
        Public ReadOnly Property TitleFldName()
            Get
                Return cTitleFldName
            End Get
        End Property
        Public ReadOnly Property ItemName()
            Get
                Return cItemName
            End Get
        End Property
        Public ReadOnly Property DataFldName()
            Get
                Return cDataFldName
            End Get
        End Property
        Public ReadOnly Property HeaderText()
            Get
                Return cHeaderText
            End Get
        End Property
        Public ReadOnly Property SortExpression()
            Get
                Return cSortExpression
            End Get
        End Property
        Public ReadOnly Property Visible()
            Get
                Return cVisible
            End Get
        End Property
        Public ReadOnly Property DataFormatString()
            Get
                Return cDataFormatString
            End Get
        End Property
        Public ReadOnly Property Attributes()
            Get
                Return cAttributes
            End Get
        End Property

        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String)
            cColumnType = ColumnType
            cItemName = ItemName
            cDataFldName = DataFldName
        End Sub

        Public Sub New(ByVal ColumnType As DG.ColumnType, ByVal ItemName As String, ByVal DataFldName As String, ByVal HeaderText As String, ByVal SortExpression As String, ByVal Visible As Boolean, ByVal DataFormatString As String, ByVal TitleFldName As String, ByVal Attributes As String)
            cColumnType = ColumnType
            cItemName = ItemName
            cDataFldName = DataFldName
            cHeaderText = HeaderText
            cSortExpression = SortExpression
            cVisible = Visible
            cDataFormatString = DataFormatString
            cTitleFldName = TitleFldName
            If Attributes = Nothing Then
                cAttributes = String.Empty
            Else
                cAttributes = Attributes
            End If
        End Sub
    End Class
#End Region

#Region " SortItem "
    Public Class SortItem
        Private cSortExpression As String
        Private cPosition As String
        Private cSortDirection As String
        Private cLastFieldSorted As String
        Public ReadOnly Property SortExpression()
            Get
                Return cSortExpression
            End Get
        End Property
        Public ReadOnly Property Position()
            Get
                Return cPosition
            End Get
        End Property
        Public Property SortDirection()
            Get
                Return cSortDirection
            End Get
            Set(ByVal Value)
                cSortDirection = Value
            End Set
        End Property
        Public Property LastFieldSorted()
            Get
                Return cLastFieldSorted
            End Get
            Set(ByVal Value)
                cLastFieldSorted = Value
            End Set
        End Property
        Public Sub New(ByVal SortExpression As String, ByVal Position As String)
            cSortExpression = SortExpression
            cPosition = Position
            cLastFieldSorted = "F"
            cSortDirection = "N"
        End Sub
    End Class
#End Region

#Region " Filter "
    Public Class Filter
        Private cColl As New Collection

        Default Public ReadOnly Property Item(ByVal ItemName As String) As Filter.FilterItem
            Get
                Return cColl(ItemName)
            End Get
        End Property

        Public ReadOnly Property Coll()
            Get
                Return cColl
            End Get
        End Property

        'Public Sub New(ByVal FilterInitialOnOff As Boolean, ByVal FilterInitialHideRecords As Boolean)
        '    cFilterInitialOnOff = FilterInitialOnOff
        '    cFilterInitialHideRecords = FilterInitialHideRecords
        'End Sub

        Public Sub AddTextbox(ByVal ItemName As String, ByVal DataFldName As String, ByVal MaxLength As Integer)
            cColl.Add(New TextboxCtl(ItemName, DataFldName, MaxLength, Nothing, Nothing), ItemName)
        End Sub

        Public Sub AddTextbox(ByVal ItemName As String, ByVal DataFldName As String, ByVal MaxLength As Integer, ByVal FilterField As String, ByVal DefaultValue As String)
            cColl.Add(New TextboxCtl(ItemName, DataFldName, MaxLength, FilterField, DefaultValue), ItemName)
        End Sub

        Public Sub AddDropdown(ByVal ItemName As String, ByVal DataFldName As String)
            cColl.Add(New DropdownCtl(ItemName, True, False, DataFldName), ItemName)
        End Sub

        Public Sub AddExtendedDropdown(ByVal ItemName As String)
            cColl.Add(New DropdownCtl(ItemName, False, True, Nothing), ItemName)
        End Sub

        Public Class DropdownCtl
            Inherits FilterItem
            Private cColl As New Collection
            Private cSelectedValue As String = String.Empty
            Private cIsStandard As Boolean
            Private cIsExtended As Boolean

            Public ReadOnly Property IsStandard()
                Get
                    Return cIsStandard
                End Get
            End Property

            Public ReadOnly Property IsExtended()
                Get
                    Return cIsExtended
                End Get
            End Property

            Public ReadOnly Property Coll()
                Get
                    Return cColl
                End Get
            End Property

            Public Overrides Sub SetValue(ByVal Value As String)
                cSelectedValue = Value
            End Sub

            Public Overrides Function GetValue() As String
                Return cSelectedValue
                'Dim i As Integer
                'For i = 1 To cColl.Count
                '    If cColl(i).Selected = True Then
                '        Return cColl(i).Value
                '    End If
                'Next
                'Return String.Empty
            End Function

            Public Sub New(ByVal ItemName As String, ByVal IsStandard As Boolean, ByVal IsExtended As Boolean, ByVal DataFldName As String)
                MyBase.new(ItemName, DataFldName, False, True)
                cIsStandard = IsStandard
                cIsExtended = IsExtended
            End Sub

            Public Overrides Sub AddDropdownItem(ByVal Value As String, ByVal Text As String, Optional ByVal Selected As Boolean = False)
                If Selected Then
                    cSelectedValue = Value
                End If
                cColl.Add(New DropdownItem(Value, Text), Value)
            End Sub

            Public Overrides Sub AddExtendedDropdownItem(ByVal Value As String, ByVal Text As String, ByVal Sql As String, Optional ByVal Selected As Boolean = False)
                If Selected Then
                    cSelectedValue = Value
                End If
                cColl.Add(New DropdownItem(Value, Text, Sql), Value)
            End Sub

            Public Class DropdownItem
                Private cValue As String
                Private cText As String
                Private cSql As String

                Public ReadOnly Property Sql()
                    Get
                        Return cSql
                    End Get
                End Property
                Public ReadOnly Property Value()
                    Get
                        Return cValue
                    End Get
                End Property
                Public ReadOnly Property Text()
                    Get
                        Return cText
                    End Get
                End Property

                Public Sub New(ByVal Value As String, ByVal Text As String)
                    cValue = Value
                    cText = Text
                End Sub

                Public Sub New(ByVal Value As String, ByVal Text As String, ByVal Sql As String)
                    cValue = Value
                    cText = Text
                    cSql = Sql
                End Sub

            End Class

        End Class

        Public Class TextboxCtl
            Inherits FilterItem
            Private cMaxLength As Integer
            Private cFilterField As String
            Private cText As String

            Public Overrides Function GetValue() As String
                Return cText
            End Function
            Public Overrides Sub SetValue(ByVal Value As String)
                cText = Value
            End Sub

            Public ReadOnly Property MaxLength()
                Get
                    Return cMaxLength
                End Get
            End Property

            Public ReadOnly Property FilterField()
                Get
                    Return cFilterField
                End Get
            End Property

            Public Sub New(ByVal ItemName As String, ByVal DataFldName As String, ByVal MaxLength As Integer)
                MyBase.New(ItemName, DataFldName, True, False)
                cMaxLength = MaxLength
            End Sub

            Public Sub New(ByVal ItemName As String, ByVal DataFldName As String, ByVal MaxLength As Integer, ByVal FilterField As String, ByVal Text As String)
                MyBase.New(ItemName, DataFldName, True, False)
                cMaxLength = MaxLength
                cFilterField = FilterField
                cText = Text
            End Sub
        End Class

        Public Class FilterItem
            Private cItemName As String
            Private cDataFldName As String
            Private cIsTextBox As Boolean
            Private cIsDropdown As Boolean

            Public ReadOnly Property CtlName()
                Get
                    If cIsTextBox Then
                        Return "txt" & cDataFldName
                    ElseIf cIsDropdown Then
                        Return "dd" & cDataFldName
                    End If
                End Get
            End Property
            Public ReadOnly Property ItemName()
                Get
                    Return cItemName
                End Get
            End Property
            Public ReadOnly Property DataFldName()
                Get
                    Return cDataFldName
                End Get
            End Property
            Public ReadOnly Property IsTextBox()
                Get
                    Return cIsTextBox
                End Get
            End Property
            Public ReadOnly Property IsDropdown()
                Get
                    Return cIsDropdown
                End Get
            End Property

            Public Sub New(ByVal ItemName As String, ByVal DataFldName As String, ByVal IsTextBox As Boolean, ByVal IsDropdown As Boolean)
                cItemName = ItemName
                cDataFldName = DataFldName
                cIsTextBox = IsTextBox
                cIsDropdown = IsDropdown
            End Sub

            Public Overridable Sub AddDropdownItem(ByVal Value As String, ByVal Text As String, Optional ByVal Selected As Boolean = False)
                'cColl.Add(New DropdownItem(Value, Text, Selected), Value)
            End Sub

            Public Overridable Sub AddExtendedDropdownItem(ByVal Value As String, ByVal Text As String, ByVal Sql As String, Optional ByVal Selected As Boolean = False)
            End Sub

            Public Overridable Function GetValue() As String
            End Function

            Public Overridable Sub SetValue(ByVal Value As String)
            End Sub

        End Class

    End Class
#End Region

#Region " ChildTables "
    Public Class ChildTablesClass
        Private cItemName As String
        Private cDataFldName As String
        ' Private cValue As String
        Private cPermissionFldName As String
        Private cChildTableSelectColumn As ChildTableSelectColumn

        Public Property ChildTableSelectColumn()
            Get
                Return cChildTableSelectColumn
            End Get
            Set(ByVal Value)
                cChildTableSelectColumn = Value
            End Set
        End Property

        Public ReadOnly Property ItemName()
            Get
                Return cItemName
            End Get
        End Property
        Public ReadOnly Property DataFldName()
            Get
                Return cDataFldName
            End Get
        End Property
        'Public Property Value()
        '    Get
        '        Return cValue
        '    End Get
        '    Set(ByVal Value)
        '        cValue = Value
        '    End Set
        'End Property
        Public ReadOnly Property PermissionFldName()
            Get
                Return cPermissionFldName
            End Get
        End Property

        Public Sub New(ByVal ItemName As String, ByVal DataFldName As String, ByVal PermissionFldName As String)
            cItemName = ItemName
            cDataFldName = DataFldName
            cPermissionFldName = PermissionFldName
        End Sub
    End Class
#End Region

#Region " Menu "
    Public Class Menu
        Private cColl As Collection

        Public Enum ObjectTypeEnum
            IsLink = 1
            IsButton = 2
        End Enum
        Public ReadOnly Property Coll()
            Get
                Return cColl
            End Get
        End Property

        Public Sub New()
            cColl = New Collection
        End Sub

        Public Sub AddItem(ByVal ObjectType As ObjectTypeEnum, ByVal OnClickMethod As String, ByVal Text As String, ByVal Right As String)
            cColl.Add(New MenuItem(ObjectType, OnClickMethod, Text, Right))
        End Sub

        Public Class MenuItem
            Private cIsLink As Boolean
            Private cIsButton As Boolean
            Private cOnClickMethod As String
            Private cText As String
            Private cIsVisible As Boolean = True
            Private cRight As String

            Public ReadOnly Property Right()
                Get
                    Return cRight
                End Get
            End Property

            Public Property IsVisible()
                Get
                    Return cIsVisible
                End Get
                Set(ByVal Value)
                    cIsVisible = Value
                End Set
            End Property

            Public ReadOnly Property Text()
                Get
                    Return cText
                End Get
            End Property
            Public ReadOnly Property IsLink()
                Get
                    Return cIsLink
                End Get
            End Property
            Public ReadOnly Property IsButton()
                Get
                    Return cIsButton
                End Get
            End Property
            Public ReadOnly Property OnClickMethod()
                Get
                    Return cOnClickMethod
                End Get
            End Property

            Public Sub New(ByVal ObjectType As ObjectTypeEnum, ByVal OnClickMethod As String, ByVal Text As String, ByVal Right As String)
                If ObjectType = ObjectTypeEnum.IsButton Then
                    cIsButton = True
                Else
                    cIsLink = True
                End If
                cOnClickMethod = OnClickMethod
                cText = Text
                cRight = Right
            End Sub
        End Class
    End Class
#End Region

End Class