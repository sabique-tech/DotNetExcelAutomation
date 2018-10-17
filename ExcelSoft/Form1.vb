Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim xlApp As Excel.Application = CType(CreateObject("Excel.Application"), Excel.Application)
        Dim xlBook As Excel.Workbook = CType(xlApp.Workbooks.Add, Excel.Workbook)
        Dim xlSheet1 As Excel.Worksheet = CType(xlBook.Worksheets(1), Excel.Worksheet)

        Try

            xlBook = CType(xlApp.Workbooks.Add, Excel.Workbook)
            xlSheet1 = CType(xlBook.Sheets.Add(Count:=10), Excel.Worksheet)

            '  xlSheet1.Name = "Sheet1"
            ' xlSheet1.Cells(2, 2) = "This is column B row 2" ' Place some text in the second row of the sheet.
            xlSheet1.Application.Visible = True ' Show the sheet.

        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim app As New Excel.Application
        Dim wb As Excel.Workbook = app.Workbooks.Add()
        Dim ws As Excel.Worksheet
        app = CType(CreateObject("Excel.Application"), Excel.Application)
        ws = CType(wb.Sheets.Add(Count:=10), Excel.Worksheet)
        Dim ws1 As Excel.Worksheet = CType(wb.Sheets(1), Excel.Worksheet)
        Dim ws2 As Excel.Worksheet = CType(wb.Sheets.Add(), Excel.Worksheet)
        ws.Application.Visible = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim cnn As SqlConnection
        Dim connectionString As String
        Dim sql As String
        Dim sql2 As String
        Dim i, j As Integer

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)


        connectionString = "data source=MAK2;" &
        "initial catalog=AccountsDB;user id=sa;password=pw;"
        cnn = New SqlConnection(connectionString)
        cnn.Open()


        If xlApp.Application.Sheets.Count() < 1 Then
            xlWorkSheet = CType(xlWorkBook.Sheets.Add(), Excel.Worksheet)
        Else
            xlWorkSheet = xlWorkBook.Sheets("sheet1")
        End If
        sql = "SELECT * FROM Customer"
        Dim dscmd As New SqlDataAdapter(sql, cnn)
        Dim ds As New DataSet
        dscmd.Fill(ds)

        For i = 0 To ds.Tables(0).Rows.Count - 1
            For j = 0 To ds.Tables(0).Columns.Count - 1
                xlWorkSheet.Cells(i + 1, j + 1) =
                ds.Tables(0).Rows(i).Item(j)
            Next
        Next

        If xlApp.Application.Sheets.Count() < 2 Then
            xlWorkSheet = CType(xlWorkBook.Sheets.Add(), Excel.Worksheet)
        Else
            xlWorkSheet = xlWorkBook.Sheets("sheet2")
        End If
        sql2 = "SELECT * FROM HSN"
        Dim dscmd2 As New SqlDataAdapter(sql2, cnn)
        Dim ds2 As New DataSet
        dscmd2.Fill(ds2)

        For i = 0 To ds2.Tables(0).Rows.Count - 1
            For j = 0 To ds2.Tables(0).Columns.Count - 1
                xlWorkSheet.Cells(i + 1, j + 1) =
                ds2.Tables(0).Rows(i).Item(j)
            Next
        Next




        xlWorkSheet.SaveAs("D:\vbexc2.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        cnn.Close()

        MsgBox("You can find the file D:\vbexcel.xlsx")
    End Sub
 Private Sub btnExportToExcelFromDatabase_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportExcel.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        Dim Save As New SaveFileDialog
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        Try
            FillB2BInExcel()
            FillB2CSInExcel()
            FillHSNInExcel()

            Save.Filter = "EXCEL Files (*.xls*)|*.xls |Excel Files (*.xlsx*)|*.xlsx"
            If Save.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                xlWorkSheet.SaveAs(Save.FileName)
            End If

        Catch ex As Exception
            MsgBox("Export Excel Error " & ex.Message)
        Finally
            ReleaseObject(xlWorkSheet)
            xlWorkBook.Close(False, Type.Missing, Type.Missing)
            ReleaseObject(xlWorkBook)
            xlApp.Quit()
            ReleaseObject(xlApp)
            'Some time Office application does not quit after automation: 
            'so i am calling GC.Collect method.
            GC.Collect()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
            tspBar.Value = 0
        End Try
    End Sub

#Region "FillInExcel"
    Private Sub FillB2BInExcel()

        Try
            If xlApp.Application.Sheets.Count() < 1 Then
                xlWorkSheet = CType(xlWorkBook.Sheets.Add(), Excel.Worksheet)
            Else
                xlWorkSheet = xlWorkBook.Sheets("Sheet1")
            End If

            Dim mSQL As String = String.Empty
            mSQL = "SELECT B.[GSTIN],A.[OrderNo] As InvoiceNo,A.[OrderDate] As InvoiceDate,A.[SalesOrderId],A.[CustomerId], " & vbCrLf
            mSQL = mSQL & "A.[Amount] As InvoiceValue,B.[InvoiceType], " & vbCrLf
            mSQL = mSQL & "A.[Value] As Taxable FROM [dbo].[SalesOrder] AS A " & vbCrLf
            mSQL = mSQL & "LEFT JOIN Customer As B ON A.CustomerId = B.CustomerId"

            Dim ds As System.Data.DataSet = Utilities.SQLHelper.ExecuteDataSet(gDBConnector, CommandType.Text, mSQL)

            For i = 0 To ds.Tables(0).Rows.Count - 1
                For j = 0 To ds.Tables(0).Columns.Count - 1
                    xlWorkSheet.Cells(i + 1, j + 1) =
                    ds.Tables(0).Rows(i).Item(j)
                Next
            Next

        Catch ex As DataPortalException
            mflgError = True
            Call ErrHandler(ex)
        Catch ex As Exception
            mflgError = True
            Call ErrHandler(ex)
        End Try

    End Sub

    Private Sub FillB2CSInExcel()

        Try


        Catch ex As DataPortalException
            mflgError = True
            Call ErrHandler(ex)
        Catch ex As Exception
            mflgError = True
            Call ErrHandler(ex)
        End Try

    End Sub

    Private Sub FillHSNInExcel()

        Try
            If xlApp.Application.Sheets.Count() < 2 Then
                xlWorkSheet = CType(xlWorkBook.Sheets.Add(), Excel.Worksheet)
            Else
                xlWorkSheet = xlWorkBook.Sheets("Sheet2")
            End If

            Dim mSQL As String = String.Empty

            mSQL = " Select P.ProductId,P.ProductName,H.HSNCode, SUM(POD.Quantity) As TotalQuantity" & vbCrLf
            mSQL = mSQL & ",SUM(POD.TotalAmount) AS TaxableValue,SUM(POD.Amount) AS TotalValue," & vbCrLf
            mSQL = mSQL & "SUM(POD.CGST) AS CentralTax,SUM(POD.SGST) AS StateTax FROM Product AS P" & vbCrLf
            mSQL = mSQL & "INNER JOIN PurchaseOrderDetail POD On P.ProductId=POD.ProductID" & vbCrLf
            mSQL = mSQL & "INNER JOIN HSN H On p.HSNId=H.HSNId" & vbCrLf
            mSQL = mSQL & "GROUP BY P.ProductName,P.ProductId,H.HSNCode"

            Dim ds As System.Data.DataSet = Utilities.SQLHelper.ExecuteDataSet(gDBConnector, CommandType.Text, mSQL)

            For i = 0 To ds.Tables(0).Rows.Count - 1
                For j = 0 To ds.Tables(0).Columns.Count - 1
                    xlWorkSheet.Cells(i + 1, j + 1) =
                    ds.Tables(0).Rows(i).Item(j)
                Next
            Next

        Catch ex As DataPortalException
            mflgError = True
            Call ErrHandler(ex)
        Catch ex As Exception
            mflgError = True
            Call ErrHandler(ex)
        End Try

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class
