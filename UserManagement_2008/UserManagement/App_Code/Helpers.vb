Imports System.Data

Public Class CollX
    Inherits System.Collections.CollectionBase
    ' Private Bittem As ListItem

    Public Sub New()
        List.Add(DBNull.Value)
    End Sub

    Public Overloads ReadOnly Property Count() As Integer
        Get
            Return List.Count - 1
        End Get
    End Property

    Default Public ReadOnly Property Coll(ByVal Idx As Integer) As Object
        Get
            Return List(Idx).Value
        End Get
    End Property

    Default Public ReadOnly Property Coll(ByVal Key As String) As Object
        Get
            Dim i As Integer
            Dim KeyUpper As String
            KeyUpper = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = KeyUpper Then
                    Return List(i).Value
                End If
            Next
            Throw New CollXError("Error #3604: CallX.Coll item not found error. Key: " & Key)  'CollXError
        End Get
    End Property

    Public Function TreatKeyAsString(ByVal Key As String) As Object
        Dim i As Integer
        Dim KeyUpper As String

        Try
            KeyUpper = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = KeyUpper Then
                    Return List(i).Value
                End If
            Next
            Throw New CollXError("Error #3611: CallX.GetValueKeyAsString item not found error. Key: " & Key)  'CollXError

        Catch ex As Exception
            Throw New CollXError("Error #3611: CallX.GetValueKeyAsString item not found error. Key: " & Key)  'CollXError
        End Try
    End Function
    Public ReadOnly Property Key(ByVal Idx As Integer) As String
        Get
            Dim i As Integer
            For i = 1 To List.Count - 1
                If i = Idx Then
                    'Return List(i).Value
                    Return List(i).Key
                End If
            Next
            Return Nothing
        End Get
    End Property
    Public ReadOnly Property DoesKeyExist(ByVal Key As String) As Boolean
        Get
            Dim i As Integer
            Key = Key.ToUpper
            For i = 1 To List.Count - 1
                If List(i).Key.ToUpper = Key Then
                    Return True
                End If
            Next
            Return False
        End Get
    End Property

    Public Sub Assign(ByVal Key As String, ByVal Value As Object)
        Dim i As Integer
        Dim Found As Boolean

        Try

            For i = 1 To List.Count - 1
                If List(i).Key = Key Then
                    List(i).Value = Nothing
                    List(i).Value = Value
                    Found = True
                End If
            Next

            'For Each Item In List
            '    If Item.Key = Key Then
            '        Item.Value = Value
            '        Found = True
            '    End If
            'Next
            If Not Found Then
                List.Add(New KeyValuePair(Key, Value))
            End If

        Catch ex As Exception
            Throw New CollXError("Error #3604: CallX.Assign. Item not found error. Key: " & Key)
        End Try
    End Sub

    Public Sub Assign(ByVal Key_Value As String)
        Dim i As Integer
        Dim Found As Boolean

        Try

            For i = 1 To List.Count - 1
                If List(i).Key = Key_Value Then
                    List(i).Value = Nothing
                    List(i).Value = Key_Value
                    Found = True
                End If
            Next

            'For Each Item In List
            '    If Item.Key = Key Then
            '        Item.Value = Value
            '        Found = True
            '    End If
            'Next
            If Not Found Then
                List.Add(New KeyValuePair(Key_Value, Key_Value))
            End If

        Catch ex As Exception
            Throw New CollXError("Error #3605: CallX.Assign. " & ex.Message)
        End Try
    End Sub

    'Public Sub ConvertArr(ByRef obj As Object)
    '    Try

    '        Dim i As Integer
    '        For i = 0 To obj.GetUpperBound(0)
    '            Assign(obj(i))
    '        Next

    '    Catch ex As Exception
    '        Throw New CollXError("Error #3606: CallX.ConvertArr. " & ex.Message)
    '    End Try
    'End Sub

    Public Sub ConvertRow(ByRef dr As DataRow)
        Try

            Dim i As Integer
            For i = 0 To dr.ItemArray.GetUpperBound(0)
                Assign(dr.Table.Columns(i).ColumnName, dr(i))
            Next

        Catch ex As Exception
            Throw New CollXError("Error #3610: CallX.ConvertRow. " & ex.Message)
        End Try
    End Sub

    Public Function ConvertToStr(ByVal Delimiter As String) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        Try

            For i = 1 To List.Count - 1
                If i < List.Count - 1 Then
                    sb.Append(List(i).Value & Delimiter)
                Else
                    sb.Append(List(i).Value)
                End If
            Next
            Return sb.ToString
        Catch ex As Exception
            Throw New CollXError("Error #3612: CallX.ConvertToStr. " & ex.Message)
        End Try
    End Function

    'Public Sub ConvertStr(ByRef Input As String, ByRef Delimiter As String)
    '    Dim i As Integer
    '    Dim Box As String()

    '    Try

    '        If Input.Length > 0 Then
    '            If Input.Substring(Input.Length - 1) = Delimiter Then
    '                Input = Input.Substring(0, Input.Length - 1)
    '            End If
    '            Box = Split(Input, Delimiter)
    '            For i = 0 To Box.GetUpperBound(0)
    '                Assign(Box(i))
    '            Next
    '        End If

    '    Catch ex As Exception
    '        Throw New CollXError("Error #3607: CallX.ConvertStr. " & ex.Message)
    '    End Try
    'End Sub

    Public Function CollxToSql() As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        For i = 1 To List.Count - 1
            If i < List.Count - 1 Then
                sb.Append(List(i).Key & "=" & List(i).Value & ", ")
            Else
                sb.Append(List(i).Key & "=" & List(i).Value)
            End If
        Next
        Return sb.ToString
    End Function

    Public Function CollxToParameters() As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder

        For i = 1 To List.Count - 1
            If i = 1 Then
                sb.Append("?" & List(i).Key & "=" & List(i).Value)
            Else
                sb.Append("&" & List(i).Key & "=" & List(i).Value)
            End If
        Next

        Return sb.ToString
    End Function

#Region " New from... "
    Public Shared Function NewFromList(ByVal Input As String, ByVal Delimiter As String) As CollX
        Dim i As Integer
        Dim Box As String()
        Dim Coll As New CollX
        Box = Input.Split("|")
        If Box.GetUpperBound(0) > -1 Then
            For i = 0 To Box.GetUpperBound(0)
                Coll.Assign(Box(i))
            Next
        End If
        Return Coll
    End Function

    Public Shared Function NewFromDataRow(ByRef dr As DataRow) As CollX
        Dim i As Integer
        Dim dt As DataTable
        Dim Coll As New CollX
        dt = dr.Table
        For i = 0 To dt.Columns.Count - 1
            Coll.Assign(dt.Columns(i).ColumnName, dr(i))
        Next
        Return Coll
    End Function

    Public Shared Function NewFromTable(ByRef dt As DataTable) As CollX
        Dim i As Integer
        Dim Coll As New CollX

        If dt.Columns.Count = 1 Then
            For i = 0 To dt.Rows.Count - 1
                Coll.Assign(dt.Rows(i)(0), dt.Rows(i)(0))
            Next
        Else
            For i = 0 To dt.Rows.Count - 1
                Coll.Assign(dt.Rows(i)(0), dt.Rows(i)(1))
            Next
        End If

        Return Coll
    End Function

    Public Shared Function NewFromKeyValue(ByVal Input As String, ByVal RowDelimter As String, ByVal ColDelimter As String) As CollX
        Dim i As Integer
        Dim Box As String()
        Dim Box2 As String()
        Dim Coll As New CollX
        Box = Input.Split(RowDelimter)
        For i = 0 To Box.GetUpperBound(0)
            Box2 = Box(i).Split(ColDelimter)
            Coll.Assign(Box2(0), Box2(1))
        Next
        Return Coll
    End Function
#End Region

    Public Overloads Sub RemoveAt(ByVal Index As Integer)
        List.RemoveAt(Index)
    End Sub

    Public Overloads Sub Remove(ByVal Key As String)
        Dim i As Integer
        For i = 1 To List.Count - 1
            If List(i).Key.ToUpper = Key.ToUpper Then
                List.Remove(List(i))
                Exit For
            End If
        Next
    End Sub

    Public Shared Function GetDistinct(ByRef dt As DataTable, ByVal FieldName As String) As CollX
        Dim i As Integer
        Dim Coll As New CollX
        Dim FirstValue As Object = DBNull.Value

        If dt.Rows.Count = 0 Then
            Return Nothing
        Else
            Coll.Assign(dt.Rows(0)(FieldName))
            For i = 0 To dt.Rows.Count - 1
                If dt.Rows(0)(FieldName) <> Coll(Coll.Count) Then
                    Coll.Assign(dt.Rows(i)(FieldName))
                End If
            Next
        End If
        Return Coll
    End Function

    'Public Function View() As String()
    '    Dim Output(Me.Count) As String
    '    Dim Val As String

    '    Try

    '        For i = 1 To List.Count - 1
    '            Try
    '                Val = List(i).Value
    '            Catch ex As Exception
    '                Val = "<object>"
    '            End Try
    '            Output(i) = List(i).Key & "|" & Val
    '        Next
    '        Return Output

    '    Catch ex As Exception
    '        Throw New CollXError("Error #3608: CallX.View. " & ex.Message)
    '    End Try
    'End Function

    Public Function View() As String
        Dim sbOutput As New System.Text.StringBuilder
        Dim Val As String

        Try
            sbOutput.Append(Environment.NewLine)
            For i = 1 To List.Count - 1
                Try
                    Val = List(i).Value
                Catch ex As Exception
                    Val = "<object>"
                End Try
                sbOutput.Append("(" & i.ToString & ") " & List(i).Key & "|" & Val & Environment.NewLine)
            Next
            Return sbOutput.ToString

        Catch ex As Exception
            Throw New CollXError("Error #3608: CallX.View. " & ex.Message)
        End Try
    End Function

    Public Shared Function Clone(ByVal InputColl As CollX) As CollX
        Dim i As Integer
        Dim OutputColl As New CollX

        Try

            For i = 1 To InputColl.Count
                OutputColl.Assign(InputColl.Key(i), InputColl(i))
            Next
            Return OutputColl

        Catch ex As Exception
            Throw New CollXError("Error #3609: CallX.Clone. " & ex.Message)
        End Try
    End Function

    Public Class KeyValuePair
        Private cKey As String
        Private cValue As Object

        Public Sub New(ByVal Key As String, ByVal Value As Object)
            cKey = Key
            cValue = Value
        End Sub
        Public Property Key() As String
            Get
                Return cKey
            End Get
            Set(ByVal value As String)
                cKey = value
            End Set
        End Property

        Public Property Value() As Object
            Get
                Return cValue
            End Get
            Set(ByVal value As Object)
                cValue = value
            End Set
        End Property
    End Class

    Public Class CollXError
        Inherits Exception
        Private cMessage As String

        Public Sub New(ByVal Message As String)
            cMessage = Message
        End Sub

        Public Overrides ReadOnly Property Message() As String
            Get
                Return cMessage
            End Get
        End Property

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
End Class
