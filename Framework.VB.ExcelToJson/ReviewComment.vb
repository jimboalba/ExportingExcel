Imports System.IO
Imports System.Reflection
Imports System.Web.Script.Serialization

Public Class ReviewComment
    Public Property DocumentNumber As String
    Public Property ProjectCode As Long
    Public Property DisciplineCode As Long
    Public Property SequentialNumber As String
    Public Property IssueTypeName As String
    Public Property Subject As String
    Public Property Description As Object
    Public Property Level As String
    Public Property IssuesCompliance As Object
    Public Property Status As String
    Public Property FileFormat As String
    Public Property ReferenceDoc As String
    Public Property CodeApproveResubmit As String
    Public Property DarComment As String
    Public Property UserId As Long
    Public Property SnapshotImages As List(Of String)
    Public Property DaepResponseImages As List(Of String)

    Public Sub SerializeToJson(current_row_ As Integer, output_folder_ As String)

        Try
            If Not output_folder_.EndsWith("\") Then
                output_folder_ = output_folder_ & "\"
                Directory.CreateDirectory(output_folder_)
            End If

            Dim file_name = generate_json_filename(current_row_)
            Dim full_path = output_folder_ & file_name

            Dim jss = New JavaScriptSerializer()
            jss.MaxJsonLength = Integer.MaxValue

            Dim json = jss.Serialize(Me)
            File.WriteAllText(full_path, json)

        Catch ex As Exception
            Trace.WriteLine(ex, MethodBase.GetCurrentMethod().Name)
        End Try
    End Sub

    Public Function generate_json_filename(row_number_) As String

        Dim docNum, projcode, disciplinecode, sequentialnum, userid As String

        If String.IsNullOrEmpty(Me.DocumentNumber) Then
            docNum = "DocNum"
        Else
            docNum = Me.DocumentNumber
        End If

        If String.IsNullOrEmpty(Me.ProjectCode) Then
            projcode = "ProjCode"
        Else
            projcode = Me.ProjectCode
        End If

        If String.IsNullOrEmpty(Me.DisciplineCode) Then
            disciplinecode = "Discipline"
        Else
            disciplinecode = Me.DisciplineCode
        End If

        If String.IsNullOrEmpty(Me.SequentialNumber) Then
            sequentialnum = "SeqNum"
        Else
            sequentialnum = Me.SequentialNumber
        End If

        Dim filename = $"{docNum}-{row_number_}_{projcode}_{disciplinecode}_{sequentialnum}_{Me.UserId}.json"
        Return filename
    End Function
End Class
