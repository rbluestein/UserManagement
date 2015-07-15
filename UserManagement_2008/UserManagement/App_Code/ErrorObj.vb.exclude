'Public Class ErrorObj
'    Public Sub New(ByRef Page As Page, ByVal Message As String)
'        Message = Replace(Message, "#", "[sharp]")
'        Message = Replace(Message, vbCrLf, "~")
'        Page.Response.Redirect("ErrorPage.aspx?ErrorMsg=" & Message)
'    End Sub
'End Class

Imports System.Web.SessionState

Public Class ErrorObj
    Private cException As Exception
    Private cHTTPSessionState As HttpSessionState
    Private cHTTPRequest As HttpRequest
    Private cErrorAsString As String
    Private cErrorSessionNumber As String
    Private cErrorDateTime As String
    Private cIsCookieError As Boolean
    Private cIsTimeOutError As Boolean
    Private cIsPostTimeOutError As Boolean

    Public Sub New(ByVal Exception As Exception, ByVal HTTPRequest As HttpRequest)
        cException = Exception
        cHTTPRequest = HTTPRequest
    End Sub

    Public ReadOnly Property IsCookieError() As Boolean
        Get
            Return cIsCookieError
        End Get
    End Property
    Public ReadOnly Property IsTimeOutError() As Boolean
        Get
            Return cIsTimeOutError
        End Get
    End Property
    Public ReadOnly Property IsPostTimeOutError() As Boolean
        Get
            Return cIsPostTimeOutError
        End Get
    End Property

    ReadOnly Property ErrorSessionNumber() As String
        Get
            If cErrorSessionNumber Is Nothing Then
                cErrorSessionNumber = String.Empty
            End If
            Return cErrorSessionNumber
        End Get
    End Property
    ReadOnly Property ErrorDateTime() As String
        Get
            Return cErrorDateTime
        End Get
    End Property

    Public Function GetCompleteErrorString() As String
        Dim TempException As Exception
        Dim ExceptionLevel As Integer
        Dim sb As New System.Text.StringBuilder
        Dim FormData As String

        Try

            sb.Append(vbCrLf)
            sb.Append("----------------- " & cErrorDateTime & " Pacific -----------------" & vbCrLf)


            FormData = vbTab & Me.GetStringFromArray(cHTTPRequest.Form.AllKeys).Replace(vbCrLf, vbCrLf & vbTab)
            If FormData.Length > 1 Then '1 because if the form is empty it will just contain the tab prefixed to the line.
                sb.Append("Form Data:" & vbCrLf)
                'remove the last tab so it doesn't screw up formatting on the line after it.
                sb.Append(FormData.Substring(0, FormData.Length - 1))
            Else
                sb.Append("Form Data: No Form Data Found" & vbCrLf)
            End If

            sb.Append("Browser=" & cHTTPRequest.Browser.Browser & " Major Version=" & cHTTPRequest.Browser.MajorVersion & " Minor Version=" & cHTTPRequest.Browser.MinorVersion & " IP=" & cHTTPRequest.UserHostAddress & vbCrLf)

            TempException = cException

            While Not (TempException Is Nothing)

                If InStr(TempException.Message, "cookie", CompareMethod.Text) > 0 Then
                    cIsCookieError = True
                ElseIf InStr(TempException.Message, "timeout", CompareMethod.Text) > 0 Then
                    cIsTimeOutError = True
                End If

                ExceptionLevel += 1
                sb.Append(vbCrLf)
                sb.Append(ExceptionLevel & ": Error Description: " & TempException.Message & vbCrLf)
                sb.Append(ExceptionLevel & ": Source: " & Replace(TempException.Source, vbCrLf, vbCrLf & ExceptionLevel & ": ") & vbCrLf)
                sb.Append(ExceptionLevel & ": Stack Trace: " & Replace(TempException.StackTrace, vbCrLf, vbCrLf & ExceptionLevel & ": ") & vbCrLf)
                sb.Append(ExceptionLevel & ": Target Site: " & Replace(TempException.TargetSite.ToString, vbCrLf, vbCrLf & ExceptionLevel & ": ") & vbCrLf)

                TempException = TempException.InnerException
            End While

        Catch excp As Exception
            sb.Append(vbCrLf + "----I erred trying to tell you about the error...-------")
        End Try

        Return sb.ToString()

    End Function

    Public Function GetDisplayErrorMessage() As String
        Dim CurException As Exception
        Dim Coll As New Collection
        Dim Box As String()

        CurException = cException
        While Not (CurException Is Nothing)
            Coll.Add(CurException.Message)
            CurException = CurException.InnerException
        End While
        'Return Coll(Coll.Count)
        Box = Split(Coll(Coll.Count), "Error #")

        If InStr(Box(Box.GetUpperBound(0)), "cookie", CompareMethod.Text) > 0 Then
            cIsCookieError = True
        ElseIf InStr(Box(Box.GetUpperBound(0)), "posttimeout", CompareMethod.Text) > 0 Then
            cIsPostTimeOutError = True
        ElseIf InStr(Box(Box.GetUpperBound(0)), "timeout", CompareMethod.Text) > 0 Then
            cIsTimeOutError = True
        End If

        Return "Error #" & Box(Box.GetUpperBound(0))
    End Function

    Private Function GetStringFromArray(ByVal Arr As String()) As String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder
        Dim Key As String

        For i = Arr.GetLowerBound(0) To Arr.GetUpperBound(0)
            Key = CType(Arr.GetValue(i), String)
            sb.Append(Key & " - " & cHTTPRequest.Form.Item(Key) & vbCrLf)
        Next
        Return sb.ToString
    End Function

    Function PrintKeysAndValues() As String
        'Dim myEnumerator As IDictionaryEnumerator = cHTTPSessionState.GetEnumerator()
        Dim Output As New System.Text.StringBuilder("")
        Dim i As Integer
        Dim Key As String
        Dim Obj As Object

        Try

            Dim myEnumerator As System.Collections.Specialized.NameObjectCollectionBase.KeysCollection = cHTTPSessionState.Keys()
            Output.Append("Session Variables" & vbCrLf)

            For i = 0 To myEnumerator.Count - 1
                Key = myEnumerator.Item(i).ToString
                Obj = cHTTPSessionState.Item(Key)

                Output.Append(vbTab)
                If IsNothing(Obj) Then
                    Output.Append("Key= " & Key & vbTab & "Type=Nothing" & vbTab & "Value=" & "Nothing")
                ElseIf IsDBNull(Obj) Then
                    Output.Append("Key= " & Key & vbTab & "Type=DBNull" & vbTab & "Value=" & "Null")
                Else
                    Select Case Obj.GetType().Name
                        Case "DBNull"
                            Output.Append("Key= " & Key & vbTab & "Type=" & Obj.GetType().Name & vbTab & "Value=" & "null")
                        Case "String"
                            Output.Append("Key= " & Key & vbTab & "Type=" & Obj.GetType().Name & vbTab & "Value=" & Obj)
                        Case "Int32"
                            Output.Append("Key= " & Key & vbTab & "Type=" & Obj.GetType().Name & vbTab & "Value=" & CStr(Obj))
                        Case "Boolean"
                            Output.Append("Key= " & Key & vbTab & "Type=" & Obj.GetType().Name & vbTab & "Value=" & CStr(Obj))
                        Case "VariantType"
                            Output.Append("Key= " & Key & vbTab & "Type=" & Obj.GetType().Name & vbTab & "Value=" & CStr(Obj))
                        Case "XmlDocument"
                            Output.Append("Key= " & Key & vbTab & "Type=" & "XML Document" & vbTab & "Value=" & CStr(Obj.OuterXml))
                        Case Else
                            Output.Append("Key= " & Key & vbTab & "Type=" & Obj.GetType().Name & vbTab & "Value=" & "Unhandled Type")
                    End Select
                End If
                Output.Append(vbCrLf)

            Next
            Output.Append("")
        Catch eObj As Exception
            ' --- Do Nothing -- Just Don't let it fail
        End Try
        PrintKeysAndValues = Output.ToString
    End Function

End Class

