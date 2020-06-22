Imports System.IO
Imports System.Reflection

Module Module1
    Sub Main()

        Dim api As New LocalAPI
        Dim excelFilesFolderFullPath As String = GetExcelFilesFolder()

        ' BIM Review Comments
        If (True) Then
            Dim excel_file_path1 = $"{excelFilesFolderFullPath}\DXBBIM002ZZ-ZZ-DEP-TB3-DRC-ELX-XX-ZZ-00001.xlsx"
            Dim excel_file_path2 = $"{excelFilesFolderFullPath}\DXBBIM002ZZ-ZZ-DEP-TB3-DRC-MFX-XX-ZZ-00001.xlsx"
            Dim excel_file_path3 = $"{excelFilesFolderFullPath}\DXBBIM002ZZ-ZZ-DEP-TB3-DRC-MHX-XX-ZZ-00001.xlsx"
            Dim excel_file_path4 = $"{excelFilesFolderFullPath}\DXBBIM002ZZ-ZZ-DEP-TB3-DRC-MPX-XX-ZZ-00001.xlsx"

            Dim bimrev1 = New LocalAPI(excel_file_path1, 11)
            Dim bimrev2 = New LocalAPI(excel_file_path2, 14)
            Dim bimrev3 = New LocalAPI(excel_file_path3, 15)
            Dim bimrev4 = New LocalAPI(excel_file_path4, 14)

            bimrev1.DocumentNumber = "DXBBIM002ZZ-ZZ-DEP-TB3-DRC-ELX-XX-ZZ-00001"
            bimrev1.export_bim_review_comments()

            bimrev2.DocumentNumber = "DXBBIM002ZZ-ZZ-DEP-TB3-DRC-MFX-XX-ZZ-00001"
            bimrev2.export_bim_review_comments()

            bimrev3.DocumentNumber = "DXBBIM002ZZ-ZZ-DEP-TB3-DRC-MHX-XX-ZZ-00001"
            bimrev3.export_bim_review_comments()

            bimrev4.DocumentNumber = "DXBBIM002ZZ-ZZ-DEP-TB3-DRC-MPX-XX-ZZ-00001"
            bimrev4.export_bim_review_comments()
        End If

        ' BEP Comments
        If False Then
            Dim excel_file_path1 =
                    "C:\Users\jimbo\Dropbox\jimbo\master\ExcelExport\DXBBIM00200-ZZ-DEP-TB1-DRC-GEN-GN-ZZ-00002 (1)_flat.xlsx"
            Dim excel_file_path2 =
                    "C:\Users\jimbo\Dropbox\jimbo\master\ExcelExport\DPZZD00200-BMO-TB3-BEP-BIM-ZZ-XX-01001_flat.xlsx"
            Dim excel_file_path3 =
                    "C:\Users\jimbo\Dropbox\jimbo\master\ExcelExport\DPZZD00200-BMO-CP3-BEP-BIM-ZZ-XX-01001_flat.xlsx"

            Dim bep1 = New LocalAPI(excel_file_path1, 10)
            Dim bep2 = New LocalAPI(excel_file_path2, 9)
            Dim bep3 = New LocalAPI(excel_file_path3, 9)

            bep1.DocumentNumber = "DXBBIM00200-ZZ-DEP-TB1-DRC-GEN-GN-ZZ-00002"
            bep1.export_bim_bep_comments()

            bep2.DocumentNumber = "DPZZD00200-BMO-TB3-BEP-BIM-ZZ-XX-01001"
            bep2.export_bim_bep_comments()

            bep3.DocumentNumber = "DPZZD00200-BMO-CP3-BEP-BIM-ZZ-XX-01001"
            bep3.export_bim_bep_comments()
        End If
    End Sub

    Private Function GetExcelFilesFolder() As String

        Dim currentAppDirectory As String = AppDomain.CurrentDomain.BaseDirectory
        Dim excelFilesFolder = Path.Combine(currentAppDirectory, "..", "..", "..", "ExcelFiles")
        Return New DirectoryInfo(excelFilesFolder).FullName
    End Function
End Module