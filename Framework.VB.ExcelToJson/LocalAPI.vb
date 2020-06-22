Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Reflection
Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing

Friend Class LocalAPI
    Public Property ExcelFile As String
    Public Property OutputDirectory As String
    Public Property TableRowStart As Integer
    Public Property DocumentNumber As String
    Public Property ProjectCode As String
    Public Property ProjectName As String


    Public Sub New()
    End Sub

    Public Sub New(excelFile As String, tableRowStart As Integer)
        If excelFile Is Nothing Then Throw New ArgumentNullException(NameOf(excelFile))
        Me.ExcelFile = excelFile

        Me.TableRowStart = tableRowStart
    End Sub

    Public Sub New(excelFile As String, outputDirectory As String)

        If excelFile Is Nothing Then Throw New ArgumentNullException(NameOf(excelFile))
        Me.ExcelFile = excelFile

        If outputDirectory Is Nothing Then
            Me.OutputDirectory = Path.GetDirectoryName(excelFile) & "\" & Path.GetFileNameWithoutExtension(excelFile)
        Else
            Me.OutputDirectory = outputDirectory
        End If
    End Sub

    Public Function export_bim_review_comments() As Boolean
        Return export_bim_review_comments(Me.ExcelFile, Me.TableRowStart)
    End Function

    Public Function export_bim_review_comments(excel_file_path_ As String, table_row_start_ As Integer,
                                               Optional output_folder_ As String = Nothing,
                                               Optional tab_name_ As String = Nothing) As Boolean
        Dim function_name = MethodBase.GetCurrentMethod().Name
        Try
            If Not File.Exists(excel_file_path_) Then
                Trace.WriteLine($"File not found: {excel_file_path_}", function_name)
                Return False
            Else
                If output_folder_ Is Nothing Then
                    output_folder_ = Path.GetDirectoryName(excel_file_path_) & "\" &
                                     Path.GetFileNameWithoutExtension(excel_file_path_)
                End If
            End If

            Select Case Path.GetExtension(excel_file_path_).ToLower()
                Case ".xlsx", ".xls"

                    Dim document_num_pos = ""

                    Dim project_code_col As Char = "A"
                    Dim discipline_code_col As Char = "B"
                    Dim sequential_number_col As Char = "C"
                    Dim subject_col As Char = "D"
                    Dim description_col As Char = "E"
                    Dim comments_col As Char = "F"
                    Dim level_col As Char = "G"
                    Dim snapshots_col As Char = "H"
                    Dim issues_compliance_col As Char = "I"
                    Dim status_col As Char = "J"
                    Dim file_format_col As Char = "K"
                    Dim reference_doc_col As Char = "L"
                    Dim code_approved_resubmit_col As Char = "M"
                    Dim contractor_response_col As Char = "N"
                    Dim daep_response_col As Char = "O"

                    Dim snapshots_col_number = 8
                    Dim daep_response_col_number = 15

                    Dim fileInfo = New FileInfo(excel_file_path_)
                    Dim package = New ExcelPackage(fileInfo)
                    Dim worksheet As ExcelWorksheet
                    If tab_name_ Is Nothing Then
                        worksheet = package.Workbook.Worksheets.FirstOrDefault
                    Else
                        worksheet = package.Workbook.Worksheets(tab_name_)
                    End If
                    Dim rows As Integer = worksheet.Dimension.Rows
                    Dim columns As Integer = worksheet.Dimension.Columns

                    If String.IsNullOrEmpty(Me.DocumentNumber) Then
                        Me.DocumentNumber = worksheet.Cells(document_num_pos)?.Value
                    End If

                    Dim api As New LocalAPI

                    For current_row As Integer = table_row_start_ To rows

                        Dim project_code As String = worksheet.Cells(project_code_col & current_row)?.Value
                        Dim displine_code As String = worksheet.Cells(discipline_code_col & current_row)?.Value
                        Dim sequential_number As String = worksheet.Cells(sequential_number_col & current_row)?.Value
                        Dim subject As String = worksheet.Cells(subject_col & current_row)?.Value
                        Dim description As String = worksheet.Cells(description_col & current_row)?.Value
                        Dim comments As String = worksheet.Cells(comments_col & current_row)?.Value
                        Dim level As String = worksheet.Cells(level_col & current_row)?.Value
                        Dim snapshots As String = worksheet.Cells(snapshots_col & current_row)?.Value
                        Dim issues_compliance As String = worksheet.Cells(issues_compliance_col & current_row)?.Value
                        Dim status As String = worksheet.Cells(status_col & current_row)?.Value
                        Dim file_format As String = worksheet.Cells(file_format_col & current_row)?.Value
                        Dim reference_doc As String = worksheet.Cells(reference_doc_col & current_row)?.Value
                        Dim code_approved_resubmit As String =
                                worksheet.Cells(code_approved_resubmit_col & current_row)?.Value
                        Dim contractor_response As String =
                                worksheet.Cells(contractor_response_col & current_row)?.Value
                        Dim daep_response As String = worksheet.Cells(daep_response_col & current_row)?.Value

                        Dim snapshot_images As New List(Of String)
                        Dim daep_response_images As New List(Of String)

#Region "getting pictures"

                        Dim excel_pic As ExcelPicture
                        Dim cell_position As ExcelDrawing.ExcelPosition

                        For Each drawing As ExcelDrawing In worksheet.Drawings

                            excel_pic = TryCast(drawing, ExcelPicture)
                            If excel_pic IsNot Nothing Then

                                cell_position = drawing.From
                                If cell_position.Row = current_row Then

                                    Trace.WriteLine(
                                        $"Cell ({current_row}:{cell_position.Column}) Picture is { _
                                                       excel_pic.ImageFormat.ToString()}", function_name)

                                    If cell_position.Column = (snapshots_col_number - 1) Then

                                        Dim b64_snapshot As String = api.convert_image_to_base64(excel_pic.Image,
                                                                                                 ImageFormat.Png)

                                        If b64_snapshot IsNot Nothing Then
                                            snapshot_images.Add(b64_snapshot)
                                        Else
                                            Trace.WriteLine(
                                                $"Cell ({current_row}:{cell_position.Column _
                                                               }) Conversion to base64 failed", function_name)
                                        End If

                                    ElseIf cell_position.Column = (daep_response_col_number - 1) Then

                                        Dim b64_daep_response As String = api.convert_image_to_base64(excel_pic.Image,
                                                                                                      ImageFormat.Png)

                                        If b64_daep_response IsNot Nothing Then
                                            daep_response_images.Add(b64_daep_response)
                                        Else
                                            Trace.WriteLine(
                                                $"Cell ({current_row}:{cell_position.Column _
                                                               }) Conversion to base64 failed", function_name)
                                        End If
                                    End If
                                End If

                            Else
                                'Trace.WriteLine("Excel picture is null", function_name)
                            End If
                        Next
                        Trace.WriteLine(
                            $"Cell ({current_row}:{cell_position.Column}) Snapshots: {snapshot_images.Count _
                                           }   Daep response:  {daep_response_images.Count}", function_name)
#End Region

                        Dim reviewComment As New ReviewComment
                        reviewComment.DocumentNumber = Me.DocumentNumber
                        If IsNumeric(project_code) Then reviewComment.ProjectCode = CLng(project_code)
                        If IsNumeric(displine_code) Then reviewComment.DisciplineCode = CLng(displine_code)
                        reviewComment.SequentialNumber = sequential_number
                        reviewComment.IssueTypeName = subject
                        reviewComment.Subject = description
                        reviewComment.Description = comments
                        reviewComment.Level = level
                        reviewComment.SnapshotImages = snapshot_images
                        reviewComment.IssuesCompliance = issues_compliance
                        reviewComment.Status = status
                        reviewComment.FileFormat = file_format
                        reviewComment.ReferenceDoc = reference_doc
                        reviewComment.CodeApproveResubmit = code_approved_resubmit
                        reviewComment.DarComment = contractor_response
                        reviewComment.DaepResponseImages = daep_response_images
                        reviewComment.UserId = get_assigned_user_on_discipline(displine_code)

                        reviewComment.SerializeToJson(current_row, output_folder_)
                    Next

                Case Else
                    Trace.WriteLine("Not an excel file: " & excel_file_path_, function_name)
            End Select

            Return True
        Catch ex As Exception
            'api.pop_error(ex)
            'api.write_log("Could not import excel file", ex)

            Trace.WriteLine(ex, function_name)
            Throw ex
            Return False
        End Try
    End Function

    Public Function export_bim_bep_comments() As Boolean
        Me.export_bim_bep_comments(Me.ExcelFile, Me.TableRowStart)
    End Function

    Public Function export_bim_bep_comments(excel_file_path_ As String, table_row_start_ As Integer,
                                            Optional output_folder_ As String = Nothing,
                                            Optional tab_name_ As String = Nothing) As Boolean
        Dim function_name = MethodBase.GetCurrentMethod().Name
        Try
            If Not File.Exists(excel_file_path_) Then
                Trace.WriteLine($"File not found: {excel_file_path_}", function_name)
                Return False
            Else
                If output_folder_ Is Nothing Then
                    output_folder_ = Path.GetDirectoryName(excel_file_path_) & "\" &
                                     Path.GetFileNameWithoutExtension(excel_file_path_)
                End If
            End If

            Dim isFlatFile = True

            Select Case Path.GetExtension(excel_file_path_).ToLower()
                Case ".xlsx", ".xls"

                    Dim document_num_pos = ""
                    Dim project_num_pos = ""
                    Dim project_name_pos = ""

                    Dim group_header_col As Char
                    Dim issueNumCol_a As Char
                    Dim issueNumCol_b As Char
                    Dim documentSectionCol_a As Char
                    Dim documentSectionCol_b As Char
                    Dim description_col As Char
                    Dim status_col As Char
                    Dim pinnacle_action_comment_col As Char
                    Dim bmo_comment_col As Char

                    If (isFlatFile) Then
                        group_header_col = "A"
                        issueNumCol_a = "B"
                        issueNumCol_b = "C"
                        documentSectionCol_a = "D"
                        documentSectionCol_b = "E"
                        description_col = "G"
                        status_col = "H"
                        pinnacle_action_comment_col = "I"
                        bmo_comment_col = "J"
                    Else

                    End If

                    Dim fileInfo = New FileInfo(excel_file_path_)
                    Dim package = New ExcelPackage(fileInfo)
                    Dim worksheet As ExcelWorksheet
                    If tab_name_ Is Nothing Then
                        worksheet = package.Workbook.Worksheets.FirstOrDefault
                    Else
                        worksheet = package.Workbook.Worksheets(tab_name_)
                    End If
                    Dim rows As Integer = worksheet.Dimension.Rows
                    Dim columns As Integer = worksheet.Dimension.Columns

                    If String.IsNullOrEmpty(Me.DocumentNumber) Then
                        Me.DocumentNumber = worksheet.Cells(document_num_pos)?.Value
                    End If

                    If String.IsNullOrEmpty(Me.ProjectCode) Then
                        Me.ProjectCode = worksheet.Cells(project_num_pos)?.Value
                    End If

                    If String.IsNullOrEmpty(Me.ProjectName) Then
                        Me.ProjectName = worksheet.Cells(project_name_pos)?.Value
                    End If

                    Dim api As New LocalAPI
                    For current_row As Integer = table_row_start_ To rows

                        Dim group_header As String = worksheet.Cells(group_header_col & current_row)?.Value
                        Dim issue_num_a As String = worksheet.Cells(issueNumCol_a & current_row)?.Value
                        Dim issue_num_b As String = worksheet.Cells(issueNumCol_b & current_row)?.Value
                        Dim document_section_a As String = worksheet.Cells(documentSectionCol_a & current_row)?.Value
                        Dim document_section_b As String = worksheet.Cells(documentSectionCol_b & current_row)?.Value
                        Dim description As String = worksheet.Cells(description_col & current_row)?.Value
                        Dim status As String = worksheet.Cells(status_col & current_row)?.Value
                        Dim pinnacle_action_comment As String =
                                worksheet.Cells(pinnacle_action_comment_col & current_row)?.Value
                        Dim bmo_comment As String = worksheet.Cells(bmo_comment_col & current_row)?.Value

                        Dim bepComment As New BepComment
                        bepComment.DocumentNumber = Me.DocumentNumber
                        bepComment.ProjectCode = Me.ProjectCode
                        bepComment.ProjectName = Me.ProjectName
                        bepComment.GroupHeader = group_header
                        bepComment.IssueNumberA = issue_num_a
                        bepComment.IssueNumberB = issue_num_b
                        bepComment.DocumentSectionA = document_section_a
                        bepComment.DocumentSectionB = document_section_b
                        bepComment.Description = description
                        bepComment.Status = status
                        bepComment.PinnacleActionComment = pinnacle_action_comment
                        bepComment.BmoComment = bmo_comment
                        bepComment.SerializeToJson(output_folder_)
                    Next

                Case Else
                    Trace.WriteLine("Not an excel file: " & excel_file_path_, function_name)
            End Select

            Return True
        Catch ex As Exception
            Trace.WriteLine(ex, function_name)
            Throw ex
            Return False
        End Try
    End Function

    Public Function get_assigned_user_on_discipline(discipline As Integer) As Long
        Select Case discipline
            Case 0
                Return 214446053
            Case 1
                Return 208522440
            Case 2
                Return 213754322
            Case 3
                Return 214446053
            Case Else
                Return 214446053
        End Select
    End Function


    Public Function convert_image_to_base64(image As Image, format As ImageFormat) As String

        Dim str As String
        Try

            If image IsNot Nothing Then

                Dim memoryStream = New MemoryStream()
                image.Save(CType(memoryStream, Stream), ImageFormat.Png)
                memoryStream.Close()
                Dim array As Byte() = memoryStream.ToArray()
                memoryStream.Dispose()
                str = Convert.ToBase64String(array)
            End If

        Catch ex As Exception

        End Try
        Return str
    End Function
End Class