
Imports System.ComponentModel
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Win32
Imports OfficeOpenXml

Class MainWindow

    Dim backgroundWorker1 As New BackgroundWorker With {.WorkerReportsProgress = True}
    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        AddHandler backgroundWorker1.DoWork, AddressOf backgroundWorker1_DoWork
        AddHandler backgroundWorker1.RunWorkerCompleted, AddressOf backgroundWorker1_RunWorkerCompleted
        AddHandler backgroundWorker1.ProgressChanged, AddressOf backgroundWorker1_ProgressChanged
    End Sub
    WithEvents WordApp As New Microsoft.Office.Interop.Word.ApplicationClass With {.Visible = False}
    Private Sub backgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs)
        Dim f As String = e.Argument(0)
        Dim dt As System.Data.DataTable = e.Argument(1)
        Try
            Dim doc = WordApp.Documents.Open(f, False, True)
            WordApp.ActiveWindow.View.ReadingLayout = False
            For i As Integer = 0 To dt.Rows.Count - 1
                For j As Integer = 0 To dt.Columns.Count - 1
                    doc.Content.Find.Execute(dt.Columns(j).ColumnName, False, True, False, False, False, True, False, False, dt.Rows(i)(j).ToString)
                Next
                WordApp.ActiveDocument.SaveAs2(f.Replace(f.Split("\").Last, "") & dt.Rows(i)(0) & ".pdf", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF)
                doc.Undo(dt.Columns.Count + 2)
                backgroundWorker1.ReportProgress(i + 1)
            Next
            WordApp.ActiveDocument.Close(False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub backgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs)
        MyProgressBar.Value = e.ProgressPercentage
        Title = MyProgressBar.Value & " of " & MyProgressBar.Maximum
    End Sub
    Private Sub backgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        MessageBox.Show("Done")
    End Sub

    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button.Click
        Dim ofd As New OpenFileDialog With {.Multiselect = False, .Filter = "الكتاب|*.doc;*.docx"}
        If ofd.ShowDialog Then
            MyProgressBar.Minimum = 0
            MyProgressBar.Maximum = ofd.FileNames.Length
            Dim dt As System.Data.DataTable = OpenExcel()
            MyProgressBar.Maximum = dt.Rows.Count
            backgroundWorker1.RunWorkerAsync({ofd.FileName, dt})
            backgroundWorker1.ReportProgress(0)
        End If
    End Sub


    Public Function OpenExcel() As System.Data.DataTable

        Dim dt As New System.Data.DataTable("Tbl")
        Try
            Dim ofd As New OpenFileDialog With {.Filter = "أسماء الشيوخ|*.xls;*.xlsx|All files (*.*)|*.*"}
            If Not ofd.ShowDialog Then Return dt
            Dim Path As String = ofd.FileName
            Dim package As New ExcelPackage(Path)

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial
            Dim currentSheet = package.Workbook.Worksheets
            Dim Worksheet = currentSheet.First()
            Dim noOfCol = Worksheet.Dimension.End.Column
            Dim noOfRow = Worksheet.Dimension.End.Row
            For j As Integer = 1 To noOfCol
                dt.Columns.Add(Worksheet.Cells(1, j).Value)
            Next
            For i As Integer = 2 To noOfRow
                If Not Worksheet.Cells(i, 1).Value Is Nothing Then
                    dt.Rows.Add()
                    For j As Integer = 1 To noOfCol
                        dt.Rows(i - 2)(j - 1) = Worksheet.Cells(i, j).Value
                    Next
                End If
            Next
            MyProgressBar.Maximum = dt.Rows.Count
            Return dt

        Catch ex As Exception
            Dim ss As String = ex.Message.ToString()
            MessageBox.Show(ss)
            Return dt
        Finally
        End Try
    End Function

    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        WordApp.Quit(False)
    End Sub
End Class
