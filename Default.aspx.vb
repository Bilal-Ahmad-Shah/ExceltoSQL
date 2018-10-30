
Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports PostiePaymentReconciliation
Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Try
            Dim folderPath As String = Server.MapPath("~/ExcelFiles/")
        Dim excelConnectionString As String = String.Empty

        'Check whether Directory (Folder) exists.
        If Not Directory.Exists(folderPath) Then
            'If Directory (Folder) does not exists. Create it.
            Directory.CreateDirectory(folderPath)
        End If

            If FileUploadExcel.HasFile Then

                Dim filePath As String = folderPath + FileUploadExcel.PostedFile.FileName
                Dim fileExt As String = Path.GetExtension(FileUploadExcel.PostedFile.FileName)

                If fileExt = ".xls" OrElse fileExt = "XLS" OrElse fileExt = ".xlsx" OrElse fileExt = "XLSX" Then
                    'Save the File to the Directory (Folder).
                    FileUploadExcel.SaveAs(folderPath & Path.GetFileName(FileUploadExcel.FileName))


                    Dim upload As PostiePaymentReconciliation.XLStoSQL
                    upload = New XLStoSQL
                    upload.xlstosql(filePath, fileExt)

                    gvExcelData.DataSource = upload.gvBind().Tables(0)
                    gvExcelData.DataBind()

                    'Display the success message.
                    lblmsg.ForeColor = System.Drawing.Color.Green
                    lblmsg.Text = Path.GetFileName(FileUploadExcel.FileName) + " has been uploaded."
                Else
                    'Display Incorrect file format
                    lblmsg.ForeColor = System.Drawing.Color.Red
                    lblmsg.Text = Path.GetFileName(FileUploadExcel.FileName) + " is incorrect file format."
                End If
            Else
                lblmsg.ForeColor = System.Drawing.Color.Red
                lblmsg.Text = "Please select the file to upload"
                Return
            End If

        Catch e1 As Exception
            Response.Write(e1.ToString())

        End Try

    End Sub


End Class
