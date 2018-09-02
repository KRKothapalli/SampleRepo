Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Object

Public Class _Default
    Inherits Page

    Public BadRecordsFilePath As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

    End Sub

    Private Function GetConnectionString() As String
        Return ConfigurationManager.ConnectionStrings("DBConnection").ConnectionString
    End Function

    Private Sub LoadDataToDatabase(ByVal tableName As String, ByVal fileFullPath As String, ByVal delimeter As String)
        BadRecordsFilePath = "C:\Users\Krishna Kothapalli\Desktop\CSVProcess\Badrecords\" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".csv"
        Dim sqlQuery As String = String.Empty
        Dim sb As StringBuilder = New StringBuilder()
        sb.AppendFormat(String.Format("BULK INSERT {0} ", tableName))
        sb.AppendFormat(String.Format(" FROM '{0}'", fileFullPath))
        sb.AppendFormat(String.Format(" WITH ( FIRSTROW = 2, FIELDTERMINATOR = '{0}' , ROWTERMINATOR = '" & vbLf & "' ,", delimeter))
        'sb.AppendFormat(String.Format(" ERRORFILE = 'C:\Users\Krishna Kothapalli\Desktop\CSVProcess\Badrecords\{0}.csv', TABLOCK)", DateTime.Now.ToString("yyyyMMddHHmmss")))
        sb.AppendFormat(String.Format(" ERRORFILE = '{0}', TABLOCK)", BadRecordsFilePath))
        sqlQuery = sb.ToString()
        Try
            Using sqlConn As SqlConnection = New SqlConnection(GetConnectionString())
                sqlConn.Open()

                Using sqlCmd As SqlCommand = New SqlCommand(sqlQuery, sqlConn)
                    sqlCmd.ExecuteNonQuery()
                    sqlConn.Close()
                End Using
            End Using
        Catch ex As Exception
        End Try

    End Sub

    Private Sub UploadAndProcessFile()
        If FileUpload1.HasFile Then
            Dim fileInfo As FileInfo = New FileInfo(FileUpload1.PostedFile.FileName)

            If fileInfo.Name.Contains(".csv") Then
                Dim fileName As String = fileInfo.Name.Replace(".csv", "").ToString()
                Dim csvFilePath As String = Server.MapPath("UploadedCSVFiles") & "\" + fileInfo.Name
                FileUpload1.SaveAs(csvFilePath)
                Dim filePath As String = Server.MapPath("UploadedCSVFiles") & "\"
                Dim strSql As String = String.Format("SELECT * FROM [{0}]", fileInfo.Name)
                Dim strCSVConnString As String = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='text;HDR=YES;'", filePath)
                Dim dtCSV As DataTable = New DataTable()
                Dim dtSchema As DataTable = New DataTable()

                Using adapter As OleDbDataAdapter = New OleDbDataAdapter(strSql, strCSVConnString)
                    adapter.FillSchema(dtCSV, SchemaType.Mapped)
                    adapter.Fill(dtCSV)
                End Using

                If dtCSV.Rows.Count > 0 Then
                    Dim fileFullPath As String = filePath & fileInfo.Name
                    'Label1.Text = String.Format("({0}) records has been loaded to the table {1}.", dtCSV.Rows.Count, fileName)
                    'Label2.Text = String.Format("The table ({0}) has been successfully created to the database.", fileName)
                    LoadDataToDatabase(fileName, fileFullPath, ",")
                    Label1.Text = String.Format("Number of records Processed from '{1}' file: {0}.", dtCSV.Rows.Count, fileName)
                    Label2.Text = String.Format("Number of Bad records found: {0}.", BadRecordsCount(BadRecordsFilePath))
                Else
                    lblError.Text = "File is empty."
                End If
            Else
                lblError.Text = "Unable to recognize file."
            End If
        End If
    End Sub

    Protected Sub btnImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnImport.Click
        UploadAndProcessFile()
    End Sub
    Function BadRecordsCount(FilePath As String) As Long
        If File.Exists(FilePath) Then
            Dim count As Integer
            count = 0
            Dim obj As StreamReader
            obj = New StreamReader(FilePath)
            ''loop through the file until the end
            Do Until obj.ReadLine Is Nothing
                count = count + 1
            Loop
            ''close file and show count
            obj.Close()

            Return count
        Else
            Return 0
        End If


    End Function

    Function Count_Lines(FilePath As String) As Long
        Dim count As Integer
        count = 0
        Dim obj As StreamReader
        obj = New StreamReader(FilePath)
        ''loop through the file until the end
        Do Until obj.ReadLine Is Nothing
            count = count + 1
        Loop
        ''close file and show count
        obj.Close()
        Return count - 1

    End Function

    'Protected Sub btnImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnImport.Click

    '    ' declare CsvDataReader object which will act as a source for data for SqlBulkCopy
    '    Using csvData = New CsvDataReader(New StreamReader(FileUpload.PostedFile.InputStream, True))
    '        ' will read in first record as a header row and
    '        ' name columns based on the values in the header row
    '        csvData.Settings.HasHeaders = True

    '        ' must define data types to use while parsing data
    '        csvData.Columns.Add("varchar") ' First
    '        csvData.Columns.Add("varchar") ' Last
    '        csvData.Columns.Add("datetime") ' Date
    '        csvData.Columns.Add("money") ' Amount

    '        ' declare SqlBulkCopy object which will do the work of bringing in data from
    '        ' CsvDataReader object, connecting to SQL Server, and handling all mapping
    '        ' of source data to destination table.
    '        Using bulkCopy = New SqlBulkCopy("Data Source=.;Initial Catalog=Test;User ID=sa;Password=")
    '            ' set the name of the destination table that data will be inserted into.
    '            ' table must already exist.
    '            bulkCopy.DestinationTableName = "Customer"

    '            ' mappings required because we're skipping the customer_id column
    '            ' and letting SQL Server handle auto incrementing of primary key.
    '            ' mappings not required if order of columns is exactly the same
    '            ' as destination table definition. here we use source column names that
    '            ' are defined in header row in file.
    '            bulkCopy.ColumnMappings.Add("First", "first_name") ' map First to first_name
    '            bulkCopy.ColumnMappings.Add("Last", "last_name") ' map Last to last_name
    '            bulkCopy.ColumnMappings.Add("Date", "first_sale") ' map Date to first_sale
    '            bulkCopy.ColumnMappings.Add("Amount", "sale_amount") ' map Amount to sale_amount

    '            ' call WriteToServer which starts import
    '            bulkCopy.WriteToServer(csvData)

    '        End Using ' dispose of SqlBulkCopy object

    '    End Using ' dispose of CsvDataReader object

    'End Sub ' end uploadButton_Click
End Class