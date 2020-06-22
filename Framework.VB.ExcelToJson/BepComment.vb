Imports System.IO
Imports System.Reflection
Imports System.Web.Script.Serialization

Public Class BepComment
    Public Property DocumentNumber As String
    Public Property ProjectCode As String
    Public Property ProjectName As String
    Public Property GroupHeader As String
    Public Property IssueNumberA As String
    Public Property IssueNumberB As String
    Public Property DocumentSectionA As String
    Public Property DocumentSectionB As String
    Public Property Description As Object
    Public Property Status As String
    Public Property PinnacleActionComment As String
    Public Property BmoComment As String

    Public Sub SerializeToJson(output_folder_ As String)

        Try
            If Not output_folder_.EndsWith("\") Then
                output_folder_ = output_folder_ & "\"
                Directory.CreateDirectory(output_folder_)
            End If

            Dim file_name = generate_json_filename()
            Dim full_path = output_folder_ & file_name

            Dim jss = New JavaScriptSerializer()
            jss.MaxJsonLength = Integer.MaxValue

            Dim json = jss.Serialize(Me)
            File.WriteAllText(full_path, json)

        Catch ex As Exception
            Trace.WriteLine(ex, MethodBase.GetCurrentMethod().Name)
        End Try
    End Sub

    Public Function generate_json_filename() As String

        Dim docNum, issueA, issueB, docSecA, docSecB As String

        If String.IsNullOrEmpty(Me.DocumentNumber) Then
            docNum = "DocNum"
        Else
            docNum = Me.DocumentNumber
        End If

        If String.IsNullOrEmpty(Me.IssueNumberA) Then
            issueA = "x"
        Else
            issueA = Me.IssueNumberA
        End If

        If String.IsNullOrEmpty(Me.IssueNumberB) Then
            issueB = "x"
        Else
            issueB = Me.IssueNumberB
        End If

        If String.IsNullOrEmpty(Me.DocumentSectionA) Then
            docSecA = "x"
        Else
            docSecA = Me.DocumentSectionA
        End If

        If String.IsNullOrEmpty(Me.DocumentSectionB) Then
            docSecB = "x"
        Else
            docSecB = Me.DocumentSectionB
        End If

        Dim filename = $"{docNum}_{issueA}_{issueB}_{docSecA}_{docSecB}.json"
        Return filename
    End Function
End Class
