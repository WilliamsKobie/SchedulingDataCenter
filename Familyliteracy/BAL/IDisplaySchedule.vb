Imports System
Imports System.Globalization
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Configuration
Imports BAL
Imports DAL

'Return Office Schedule and the schedule for a particular student within a given date range.
Public Interface IDisplaySchedule
    Function ReturnOfficeSchedule(ByVal StartDate As DateTime, ByVal EndDate As DateTime) As DataTable
    Function mainDisplay(ByVal studentid As String, ByVal StartDate As Date, ByVal EndDate As Date) As DataView
End Interface


'Return hours taht has been proposed on the schedule
Public Interface IProposedHours
    Function Proposed(ByVal dt As DataTable, ByVal studentId As String, ByVal startDate As Date, endDate As Date)
End Interface

Public Class ReschedulingDisplay
    Inherits ScheduleTemplate
    Implements IDisplaySchedule
    
    Implements IProposedHours

    Public Overloads Function MainDisplay(ByVal studentid As String, ByVal StartDate As Date, ByVal EndDate As Date) As DataView Implements IDisplaySchedule.mainDisplay

        Dim ds1 As New DataSet
        Dim schedule As New Schedule

        ds1 = StudentSchedule()

        Dim dt1 As DataTable = ds1.Tables("Clinicianinfo")

        'Return all proposed hours, rescheduled hours, and canceled hour to the datagridview
        Proposed(dt1, studentid, StartDate, EndDate)

        Dim dv As New DataView(dt1)
        dv.Sort = "Date, Hour ASC"
        Return dv

    End Function

    Public Function Proposed(ByVal dt As DataTable, ByVal studentId As String, ByVal startDate As Date, endDate As Date) Implements IProposedHours.Proposed
        Dim ds3 As New DataSet
        Dim schedule As New Schedule
        Dim date1 As Date
        Dim StartTime As DateTime
        Dim EndTime As DateTime
        Dim Location As String = String.Empty
        Dim subject As String = String.Empty
        Dim status As String = String.Empty
        Dim proposedds As New DataSet
        Dim proposeddt As DataTable
        Dim transactionId As Integer

        proposeddt = schedule.ReturnProposedSchedule(studentId, startDate, endDate, "Proposed")

        Dim row As DataRow
        Dim row1 As DataRow
        Dim row3 As DataRow


        For Each row In proposeddt.Rows
            date1 = row("Date")
            StartTime = row("TimeIn")
            EndTime = row("TimeOut")
            status = row("Status")
            transactionId = row("Count")
            ds3 = schedule.GetClassroomData(transactionId.ToString.Trim)

            Dim dt3 As DataTable = ds3.Tables("Classroom")
            For Each row3 In dt3.Rows
                Location = row3("Campus")
                subject = row3("Subject")
            Next
            row1 = dt.NewRow()
            row1("Date") = date1.ToString("M/dd/yyyy,  dddd")
            row1("Timein") = StartTime.ToString("h:mm tt")
            row1("Timeout") = EndTime.ToString("h:mm tt")
            row1("Clinician") = row("ClinicianSignature")
            row1("Processor") = row("ProcessingClinician")
            row1("Subject") = subject
            row1("Campus") = Location
            row1("Status") = status
            dt.Rows.Add(row1)
        Next
        Return Nothing
    End Function

    Public Function ReturnOfficeSchedule(ByVal StartDate As DateTime, ByVal EndDate As DateTime) As DataTable Implements IDisplaySchedule.ReturnOfficeSchedule
        Dim ds, ds1, ds2 As New DataSet
        Dim convertStudentName As INameConversion = New StudentNameConversion
        Dim schedule As New Schedule
        Dim studentname, Status As String
        Dim date1 As Date
        Dim StartTime As DateTime
        Dim EndTime As DateTime


        Dim Studentid As String
        Dim Clinician As String
        Dim Count As Integer
        ds1 = StudentCalendar()
        ds = schedule.GetSchedule(StartDate, EndDate)

        Dim dt As DataTable = ds.Tables("MainSchedule")
        Dim dt1 As DataTable = ds1.Tables("StudentCalendar")
        Dim dt2 As DataTable

        Dim row1 As DataRow
        Dim scheduleingOperator As String = String.Empty
        Dim dv As New DataView(dt)
        dv.Sort = "Date,TimeIn ASC"
        For Each rowView As DataRowView In dv
            Dim rw As DataRow = rowView.Row
            date1 = rw("Date")
            StartTime = rw("TimeIn")
            EndTime = rw("TimeOut")
            Clinician = rw("ClinicianSignature")
            Studentid = rw("Studentid")
            scheduleingOperator = rw("Processingclinician")
            Count = rw("Count")
            ds2 = schedule.GetClassroomData(Count)
            dt2 = ds2.Tables("ClassRoom")

            Status = rw("Status")
            studentname = convertStudentName.convertName(Studentid.Trim)
            row1 = dt1.NewRow()
            row1("Student Name") = studentname
            row1("New Date") = date1.ToString("M/dd/yyyy")

            row1("Start") = StartTime.ToShortTimeString
            row1("Finish") = EndTime.ToShortTimeString
            row1("Clinician") = Clinician.Trim
            row1("EnteredBy") = scheduleingOperator
            'Search the Classroom table for corresponding count id that matches the student 
            If dt2.Rows.Count > 0 Then
                For Each dr2 In dt2.Rows
                    row1("Campus") = dr2("Campus")
                    row1("Subject") = dr2("Subject")
                Next
            End If
            dt1.Rows.Add(row1)

        Next

        Return dt1

    End Function
End Class




