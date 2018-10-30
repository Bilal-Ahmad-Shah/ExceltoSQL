Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient

Namespace PostiePaymentReconciliation

    Public Class XLStoSQL
        Public Sub xlstosql(path As String, ext As String)

            Try
                Dim excelConnectionString As String = String.Empty
                Dim AcquirerID As String
                Dim OrderID As String
                Dim OrderReference As String
                Dim Amount As String
                Dim OrderDate As String
                Dim ResponseCode As String
                Dim Status As String
                Dim FoundMatchingOrder As Boolean


                Dim filePath As String = path
                Dim fileExt As String = ext


                Dim strConnection As [String] = ConfigurationManager.ConnectionStrings("XLSImport_SampleConnectionString").ConnectionString
                If fileExt = ".xls" OrElse fileExt = "XLS" OrElse fileExt = ".xlsx" OrElse fileExt = "XLSX" Then
                    excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=Excel 12.0;Persist Security Info=False"
                Else
                    Throw New System.Exception("File is of Incorrect format.")
                End If

                Dim excelConnection As New OleDbConnection(excelConnectionString)
                Dim cmd As New OleDbCommand("Select * from [Sheet1$]", excelConnection)

                excelConnection.Open()

                Dim da As New OleDbDataAdapter(cmd)
                Dim ds As New Data.DataSet()
                da.Fill(ds)

                Dim dReader As OleDbDataReader
                dReader = cmd.ExecuteReader()


                While (dReader.Read())
                    AcquirerID = Convert.ToString(dReader(0))
                    OrderID = Convert.ToString(dReader(1))
                    OrderReference = Convert.ToString(dReader(2))
                    Amount = Convert.ToString(dReader(3))
                    OrderDate = Convert.ToString(dReader(4))
                    ResponseCode = Convert.ToString(dReader(5))
                    Status = Convert.ToString(dReader(6))
                    FoundMatchingOrder = "0"


                    savedata(AcquirerID, OrderID, OrderReference, Amount, OrderDate, ResponseCode, Status, FoundMatchingOrder)

                End While

            Catch ex As Exception
                Throw New Exception("Sql connection error")

            End Try


        End Sub

        Protected Sub savedata(AcquirerID As String, OrderID As String, OrderReference As String, Amount As String, OrderDate As String, ResponseCode As String, Status As String, FoundMatchingOrder As String)
            Try

                Dim strConnection As [String] = ConfigurationManager.ConnectionStrings("XLSImport_SampleConnectionString").ConnectionString
                Dim conn As New SqlConnection(strConnection)
                If conn.State = Data.ConnectionState.Closed Then
                    conn.Open()
                End If

                Dim cmd As New SqlCommand
                cmd.CommandText = "spAnzTransaction_InputData"
                cmd.CommandType = Data.CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@AcquirerID", AcquirerID)
                cmd.Parameters.AddWithValue("@OrderID", OrderID)
                cmd.Parameters.AddWithValue("@OrderReference", OrderReference)
                cmd.Parameters.AddWithValue("@Amount", Amount)
                cmd.Parameters.AddWithValue("@OrderDate", OrderDate)
                cmd.Parameters.AddWithValue("@ResponseCode", ResponseCode)
                cmd.Parameters.AddWithValue("@Status", Status)
                cmd.Parameters.AddWithValue("@FoundMatchingOrder", FoundMatchingOrder)
                cmd.Connection = conn
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                Throw New Exception("Sql connection error")
            End Try

        End Sub

        Public Function gvBind() As Data.DataSet
            Try
                Dim strConnection As [String] = ConfigurationManager.ConnectionStrings("XLSImport_SampleConnectionString").ConnectionString
                Dim conn As New SqlConnection(strConnection)
                If conn.State = Data.ConnectionState.Closed Then
                    conn.Open()
                End If
                Dim cmd As New SqlCommand
                cmd.CommandText = "spAnzTransaction_SelectOutput"
                cmd.CommandType = Data.CommandType.StoredProcedure
                cmd.Connection = conn
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As New Data.DataSet()
                da.Fill(ds)

                da.Dispose()
                conn.Close()
                conn.Dispose()

                Return ds
            Catch ex As Exception
                Throw New Exception("Sql connection error")
            End Try

        End Function


    End Class
End Namespace
