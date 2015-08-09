Imports BAL
Imports DAL
Public Class ExportSchedule

    Private Sub ExportSchedule_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim startDate, endDate As DateTime
        startDate = DateTimePicker1.Value
        endDate = DateTimePicker2.Value
        displayStudentSchedules(startDate, endDate)
    End Sub
    'Clear and Populate Display
    Public Sub displayStudentSchedules(ByVal startDate As DateTime, ByVal endDate As DateTime)
        Dim getScheduleListing As IDisplaySchedule = New ReschedulingDisplay
        Dim ds As New DataTable

        'Clear display
        startDate = DateTimePicker1.Value
        dt = getScheduleListing.ReturnOfficeSchedule(startDate, endDate)
        dt.Columns.Remove("Prior Date")
        dt.Columns.Remove("Excuse")
        dt.Columns.Remove("TransferId")
        dt.Columns.Remove("State")
        dt.Columns.Remove("Callin date")
        dt.Columns.Remove("MeansofRequest")
        dt.Columns.Remove("Campus")
        dt.Columns.Remove("Attendance")
        dt.Columns.Remove("OriginalTransferDate")
        dt.Columns.Remove("Column1")
     
        'Populate GridView Control
   
        DataGridView1.DataSource = dt
   
    End Sub

    'Move forward or backwards a month
    Private Sub DateTimePicker1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim startDate, endDate As DateTime
        DateTimePicker2.Value = DateTimePicker1.Value
        startDate = DateTimePicker1.Value
        endDate = DateTimePicker1.Value
        displayStudentSchedules(startDate, endDate)
    End Sub
    'Export schedule to Excel
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        ExportSchedule()
    End Sub



    Public Function ExportSchedule()
        Try

            Dim app As New Microsoft.Office.Interop.Excel.Application


            REM creating new WorkBook within Excel application



            Dim workbook As Microsoft.Office.Interop.Excel.Workbook = app.Workbooks.Add(Type.Missing)


            REM creating new Excelsheet in workbook


            Dim Worksheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing

            REM see the excel sheet behind the program

            app.Visible = True



            REM get the reference of first sheet. By default its name is Sheet1.

            REM store its reference to worksheet

            Worksheet = workbook.Sheets("Sheet1")

            Worksheet = workbook.ActiveSheet



            REM changing the name of active sheet

            Worksheet.Name = "Student Schedules"






            REM storing header part in Excel

            For i = 1 To DataGridView1.Columns.Count



                Worksheet.Cells(1, i) = DataGridView1.Columns(i - 1).HeaderText
            Next







            REM storing Each row and column value to excel sheet

            For i = 0 To DataGridView1.Rows.Count - 2


                For j = 0 To DataGridView1.Columns.Count - 1


                    Worksheet.Cells(i + 2, j + 1) = DataGridView1.Rows(i).Cells(j).Value.ToString



                Next
            Next
            'Save Excel File
            workbook.SaveAs("c:\StudentSchedules.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing)

        Catch ex As Exception

        End Try

        Return Nothing
    End Function


     

    Private Sub DateTimePicker2_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        displayStudentSchedules(DateTimePicker1.Value, DateTimePicker2.Value)
    End Sub
End Class