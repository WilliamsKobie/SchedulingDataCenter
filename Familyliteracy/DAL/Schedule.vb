Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Configuration
Imports System.Data.OleDb
'This class returns the schedule of a student,clincian or multiple students and clinicians based on the given parameters being passed to it
Public Class Schedule
    Dim connectionString As Object
    Public Sub New()
        Try

            connectionString = ConfigurationManager.ConnectionStrings("Familyliteracy").ConnectionString
        Catch ex As SqlException
            MsgBox("Cannot connect to Datasource")
        End Try
    End Sub
  

    Public Overloads Function GetSchedule(ByVal startDate As Date, ByVal finalDate As Date) As DataSet

        Dim query As String = "SELECT * FROM MainSchedule where [Date] Between '" & startDate & "' AND '" & finalDate & "' Order By Date, TimeIn ASC"

        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()


        Dim dt As DataTable = ds.Tables("MainSchedule")

        Return ds
    End Function
    Public Overloads Function MainDisplaySchedule() As String

        Dim transId As String = String.Empty
        Dim query As String = "SELECT TOP 1 * FROM MainSchedule  Order By count Desc"

        Dim conn As New SqlConnection(connectionString)

        Dim cmd As New SqlCommand(query, conn)


        Dim da As New SqlDataAdapter(cmd)

        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")

        conn.Close()
        Dim row As DataRow

        Dim dt As DataTable = ds.Tables("MainSchedule")

        For Each row In dt.Rows
            transId = Convert.ToInt32(row("count"))

        Next

        Return transId


    End Function


    Public Overloads Function returnStudentScheduleinfo(ByVal studentid As String) As DataSet

        Dim index As Integer
        index = Convert.ToInt16(Studentid)

        Dim query As String = "SELECT * FROM MainSchedule where StudentID='" & index & "' ORDER BY [DATE] ASC, TimeIn ASC"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()


        Return ds

    End Function
    Public Overloads Function returnStudentScheduleinfo(ByVal appointment As String, ByVal time1 As String, ByVal time2 As String) As DataSet


        Dim appoint As Date
        Dim timestamp1 As DateTime
        Dim timestamp2 As DateTime
        appoint = Convert.ToDateTime(appointment.Trim)
        timestamp1 = Convert.ToDateTime("1900-01-01 " & time1)
        timestamp2 = Convert.ToDateTime("1900-01-01 " & time2)
        Dim query As String = "SELECT * FROM MainSchedule where [Date]='" & appoint & "' Timein='" & timestamp1 & "' AND Timeout='" & timestamp2 & "'"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()


        Return ds

    End Function
    Public Overloads Function returnStudentSchedule(ByVal studentid As String, ByVal appointment As String) As DataSet

        Dim index As Integer
        index = Convert.ToInt16(studentid)
        Dim appoint As Date

        appoint = Convert.ToDateTime(appointment.Trim)

        Dim query As String = "SELECT * FROM MainSchedule where studentid='" & index & "' AND [Date]='" & appoint & "'"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()


        Return ds

    End Function



    Public Overloads Function returnStudentScheduleinfo(ByVal studentid As String, ByVal StartDate As Date, ByVal EndDate As Date) As DataSet

        Dim index As Integer
        index = Convert.ToInt16(studentid)
        Dim query As String = "SELECT * FROM MainSchedule where [Date] Between'" & StartDate & "' And '" & EndDate & "' And Studentid='" & index & "' Order By [Date] ASC, TimeIn ASC"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()

        Return ds

    End Function
    Public Function returnProposedSchedule(ByVal studentid As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal state As String) As DataTable
        Dim index As Integer
        index = Convert.ToInt16(studentid)

        Dim query As String = "SELECT * FROM MainSchedule where [Date] Between'" & StartDate & "' And '" & EndDate & "' And Studentid='" & index & "' AND Status='" & state & "' Order By [Date] ASC, TimeIn ASC"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()
        Dim dt As DataTable = ds.Tables("MainSchedule")
        Return dt

    End Function

    Public Overloads Function returnStudentScheduleinfo(ByVal studentid As String, ByVal startDate As Date, ByVal StartTime As DateTime, ByVal endTime As DateTime) As DataSet

        Dim index As Integer
        index = Convert.ToInt16(Studentid)
        Dim query As String = "SELECT * FROM MainSchedule where [Date]='" & startDate & "' And Studentid='" & index & "' AND TimeIn='" & StartTime & "' AND TimeOut='" & endTime & "'"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()

        Return ds

    End Function


    Public Overloads Function ReturnClinicianScheduleinfo(ByVal initialScheduleDate As Date, ByVal finalScheduleDate As Date, ByVal clincianId As String) As DataTable


        Dim query As String = "SELECT * FROM MainSchedule where [Date] Between'" & initialScheduleDate & "' AND '" & finalScheduleDate & "' AND ClinicianID='" & clincianId & "' Order By Clinicianid ASC, [Date] ASC, TimeIn ASC"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()
        Dim dt As DataTable = ds.Tables("MainSchedule")

        Return dt

    End Function


    Public Overloads Function ViewSchedule(ByVal Clinician As String, ByVal Date1 As Date, ByVal Date2 As Date) As DataSet

        Dim query As String = "SELECT * FROM MainSchedule WHERE [DATE] BETWEEN '" & Date1 & "' AND '" & Date2 & "' AND ClinicianID='" & Clinician & "' ORDER BY [Date] ASC, TimeIn ASC"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()

        Dim dt As DataTable = ds.Tables("MainSchedule")

        Return ds
    End Function
    Public Overloads Function ViewSchedule(ByVal Clinician As String) As DataSet

        Dim query As String = "SELECT * FROM MainSchedule WHERE ClinicianID='" & Clinician & "' ORDER BY [Date] ASC, TimeIn ASC"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()

        Dim dt As DataTable = ds.Tables("MainSchedule")
        Return ds
    End Function


    Public Function GetClassroomData(ByVal transid As String) As DataSet
        Dim query As String = "SELECT * FROM Classroom WHERE Transactionid='" & transid.Trim & "'"

        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet

        conn.Open()
        da.Fill(ds, "Classroom")
        conn.Close()

        Return ds
    End Function

End Class




'This class is called when user data needs to modified, added or deleted from the database.
Public Class commitChanges
    Dim connectionString As Object
    Public Sub New()
        Try
            connectionString = ConfigurationManager.ConnectionStrings("Familyliteracy").ConnectionString
        Catch ex As SqlException
            MsgBox("Cannot connect to Datasource")
        End Try
    End Sub
    Public Function EditDailySchedule(ByVal index As String, ByVal Studentid As String, ByVal Clinicianid As String, ByVal ClinicianName As String, ByVal adjustedDate As String, ByVal timein As String, ByVal timeout As String, ByVal OldTimeIn As String, ByVal OldTimeOut As String, ByVal Location As String, ByVal Subject As String)

        Dim t1 As String
        Dim t2 As String
        Dim t3 As String
        Dim t4 As String

        Dim time1 As DateTime
        Dim time2 As DateTime
        Dim time3 As DateTime
        Dim time4 As DateTime
        Dim datestamp As Date
        datestamp = Convert.ToDateTime(adjustedDate)
        t1 = "1900-01-01 " & OldTimeIn.Trim
        t2 = "1900-01-01 " & OldTimeOut.Trim
        time1 = Convert.ToDateTime(t1) 'convert start time to Datetime Datatype
        time2 = Convert.ToDateTime(t2) 'convert stop time to Datetime Datatype
        t3 = "1900-01-01 " & timein.Trim
        t4 = "1900-01-01 " & timeout.Trim
        time3 = Convert.ToDateTime(t3) 'convert start time to Datetime Datatype
        time4 = Convert.ToDateTime(t4) 'convert stop time to Datetime Datatype
        Dim studentIndex As Integer
        studentIndex = Convert.ToInt16(Studentid)

        Dim query As String = "SELECT * FROM MainSchedule Where StudentID='" & studentIndex & "' AND [Date]='" & datestamp & "' AND count='" & index & "'"

        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim cmd3 As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet

        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()
        Dim dt As DataTable = ds.Tables("MainSchedule")
        Dim dr As DataRow
        Dim dr3 As DataRow
        Dim returnstate As String = String.Empty
        Dim itimein As DateTime
        Dim status As String = String.Empty
        Dim Meansofrequest As String = String.Empty
        Dim processor As String = String.Empty
        Dim CallinDate As String = String.Empty
        Dim ts As TimeSpan
        Dim totalhours As Integer = 0

        For Each dr In dt.Rows
            returnstate = dr("Status")

            status = dr("Status")
            Meansofrequest = dr("RequestedFashion")
            processor = dr("ProcessingClinician")
            CallinDate = dr("RequestedDate")
            itimein = time2.AddHours(-1)
            ts = (time2 - time1)
            totalhours = Math.Abs(ts.TotalHours)

            If totalhours <= 1 Then
                dr.BeginEdit()
                dr("TimeIn") = time3
                dr("TimeOut") = time4
                dr("Clinicianid") = Clinicianid
                dr("ClinicianSignature") = ClinicianName
                dr.EndEdit()
                Dim objCommandBuilder As New SqlCommandBuilder(da)
                da.Update(ds, "MainSchedule")
                SubmitClassroomData(index.ToString, Location, Subject)
                Exit For
            Else
                dr.BeginEdit()
                dr("TimeIn") = itimein
                dr("TimeOut") = time2
                dr.EndEdit()
                Dim objCommandBuilder As New SqlCommandBuilder(da)
                da.Update(ds, "MainSchedule")



                Dim da3 As New SqlDataAdapter(cmd3)
                Dim ds3 As New DataSet

                conn.Open()
                da3.Fill(ds3, "MainSchedule")
                conn.Close()
                Dim dt3 As DataTable = ds3.Tables("MainSchedule")
                dr3 = dt3.NewRow
                dr3("Status") = status
                dr3("RequestedFashion") = Meansofrequest
                dr3("ProcessingClinician") = processor
                dr3("RequestedDate") = CallinDate
                dr3("Studentid") = Studentid
                dr3("TimeIn") = time3
                dr3("TimeOut") = time4
                dr3("Date") = datestamp
                dr3("Clinicianid") = Clinicianid.Trim
                dr3("ClinicianSignature") = ClinicianName.Trim
                dt3.Rows.Add(dr3)
                Dim objCommandBuilder3 As New SqlCommandBuilder(da3)
                da3.Update(ds3, "MainSchedule")

                Exit For
            End If


        Next

        Return Nothing
    End Function


    Public Function RemoveStudent(ByVal studentid As String, ByVal StartTime As String, ByVal EndTime As String, ByVal datestamp As Date, ByVal clinician As String)
        Dim index As Integer
        index = Convert.ToInt16(studentid)
        Dim t1 As String
        Dim t2 As String
        Dim time1 As DateTime
        Dim time2 As DateTime
        Dim indexid As String = Nothing
        t1 = "1900-01-01 " & StartTime.Trim

        t2 = "1900-01-01 " & EndTime.Trim
        time1 = Convert.ToDateTime(t1).ToLongTimeString 'convert start time to Datetime Datatype
        time2 = Convert.ToDateTime(t2).ToLongTimeString 'convert stop time to Datetime Datatype

        Dim query As String = "SELECT * FROM MainSchedule Where StudentID='" & index & "' AND [Date]='" & datestamp & "' AND TimeIn='" & time1 & "' AND TimeOut='" & time2 & "'"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()
        Dim dt As DataTable = ds.Tables("MainSchedule")
        Dim dr As DataRow
        For Each dr In dt.Rows

            dr.Delete()
        Next
        Dim objCommandBuilder As New SqlCommandBuilder(da)
        da.Update(ds, "MainSchedule")

        Return Nothing
    End Function

    Public Function RemoveClinicianOffDays(ByVal clinicianid As String, ByVal StartTime As String, ByVal EndTime As String, ByVal datestamp As Date)

        Dim t1 As String
        Dim t2 As String
        Dim time1 As DateTime
        Dim time2 As DateTime

        t1 = "1900-01-01 " & StartTime.Trim

        t2 = "1900-01-01 " & EndTime.Trim
        time1 = Convert.ToDateTime(t1) 'convert start time to Datetime Datatype
        time2 = Convert.ToDateTime(t2) 'convert stop time to Datetime Datatype


        Dim query As String = "SELECT * FROM Clinician_DailyOutSchedule Where ClinicianID='" & clinicianid.Trim & "'" & " AND [Date]='" & datestamp & "' AND TimeIn='" & time1 & "' AND TimeOut='" & time2 & "'"

        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "Clinician_DailyOutSchedule")
        conn.Close()
        Dim dt As DataTable = ds.Tables("Clinician_DailyOutSchedule")
        Dim dr As DataRow
        Dim r As Integer = dt.Rows.Count
        For Each dr In dt.Rows
            dr.Delete()
        Next
        Dim objCommandBuilder2 As New SqlCommandBuilder(da)
        da.Update(ds, "Clinician_DailyOutSchedule")
        Return Nothing
    End Function


    Function RemoveClassroominfo(ByVal transferid As String)
        Dim query As String = "SELECT * FROM Classroom where Transactionid='" & transferid.Trim & "'"

        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet

        conn.Open()
        da.Fill(ds, "Classroom")
        conn.Close()
        Dim dt As DataTable = ds.Tables("Classroom")
        Dim dr As DataRow
        For Each dr In dt.Rows
            dr.Delete()
        Next
        Dim objCommandBuilder As New SqlCommandBuilder(da)
        da.Update(ds, "Classroom")
        Return Nothing
    End Function

    'Add a new Scheduled Date and time for a particular student and clinician
    Public Function submitstudent(ByVal index As Integer, ByVal Studentid As String, ByVal Clinician As String, ByVal Signature As String, ByVal timestamp1 As String, ByVal timestamp2 As String, ByVal datestamp As Date, ByVal status As String, ByVal DateofRequest As String, ByVal MeansofRequest As String, ByVal Location As String, ByVal Subject As String, ByVal Processor As String) As Boolean

        Dim studentindex As Integer
        studentindex = Convert.ToInt16(Studentid)
        Dim t1 As String
        Dim t2 As String
        Dim time1 As DateTime
        Dim time2 As DateTime
        Dim rqdate As Date

        t1 = "1900-01-01 " & timestamp1.Trim
        t2 = "1900-01-01 " & timestamp2.Trim
        time1 = Convert.ToDateTime(t1) 'convert start time to Datetime Datatype
        time2 = Convert.ToDateTime(t2) 'convert stop time to Datetime Datatype
        rqdate = Convert.ToDateTime(DateofRequest)

        'Update MainSchedule Table
        ' 
        Dim query As String = "SELECT * FROM MainSchedule Where StudentID='" & studentindex & "' AND [Date]='" & datestamp & "' AND TimeIn='" & time1 & "' AND TimeOut='" & time2 & "'"
        'Dim query As String = "SELECT * FROM MainSchedule"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet

        conn.Open()
        da.Fill(ds, "MainSchedule")
        conn.Close()
        Dim dt As DataTable = ds.Tables("MainSchedule")
        Dim dr As DataRow


        'add parameters to database

        Dim count As Integer = dt.Rows.Count
        If count < 1 Then
            dr = dt.NewRow()
            dr.Item("count") = index
            dr.Item("StudentID") = studentindex
            dr.Item("Date") = datestamp
            dr("TimeIn") = time1
            dr("TimeOut") = time2
            dr.Item("ClinicianID") = Clinician
            dr.Item("ClinicianSignature") = Signature
            dr("Status") = status 'Transfer, or Proposed
            dr("RequestedDate") = rqdate
            dr("RequestedFashion") = MeansofRequest.Trim
            dr("ProcessingClinician") = Processor.Trim

            dr("Attendance") = "Proposed" 'NoShow,Completed

            dt.Rows.Add(dr)

            Dim objCommandBuilder As New SqlCommandBuilder(da)
            da.Update(ds, "MainSchedule")

            Dim query3 As String = "SELECT * FROM Classroom"


            Dim cmd3 As New SqlCommand(query3, conn)
            Dim da3 As New SqlDataAdapter(cmd3)
            Dim ds3 As New DataSet

            conn.Open()
            da3.Fill(ds3, "Classroom")
            conn.Close()
            Dim dt3 As DataTable = ds3.Tables("Classroom")

            Dim dr3 As DataRow


            'add parameters to database or edit table



            dr3 = dt3.NewRow()
            dr3("TransactionId") = index
            dr3("StudentId") = studentindex
            dr3("Campus") = Location.Trim
            dr3("Subject") = Subject.Trim

            dt3.Rows.Add(dr3)

            Dim objCommandBuilder3 As New SqlCommandBuilder(da3)
            da3.Update(ds3, "Classroom")

            Return True
        Else
            MsgBox("Already Scheduled at this exact Date and Time")

        End If

        Return Nothing
    End Function

    Public Function SubmitClassroomData(ByVal Transactionid As String, ByVal Location As String, ByVal Subject As String)


        Dim query As String = "SELECT * FROM Classroom Where Transactionid='" & Transactionid & "'"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet

        conn.Open()
        da.Fill(ds, "Classroom")
        conn.Close()
        Dim dt As DataTable = ds.Tables("Classroom")

        Dim dr As DataRow
        Dim x As Integer = dt.Rows.Count
        If x < 1 Then

            dr = dt.NewRow()
            dr("TransactionId") = Transactionid
            dr("Campus") = Location
            dr("Subject") = Subject

            dt.Rows.Add(dr)

            Dim objCommandBuilder As New SqlCommandBuilder(da)
            da.Update(ds, "Classroom")
        End If
        Return Nothing
    End Function






End Class





'This class is used to store values any of the datagrid column sizes. 
'Returns the column width value that is store in the database
Public Class StoreGridViewColumnWidth
    Dim connectionString As Object
    Public Sub New()
        Try

            connectionString = ConfigurationManager.ConnectionStrings("Familyliteracy").ConnectionString
        Catch ex As Exception
            MsgBox("Cannot connect to Datasource")
        End Try
    End Sub
    Public Function ReturnColumnWidth(ByVal screen As Integer, ByVal gridViewControl As Integer) As Integer

        Dim query As String = "SELECT * FROM GridViewColumnSizeAdjustment Where FormReferenceNumber='" & screen & "' AND GridViewControlNumber='" & gridViewControl & "'"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "GridViewColumnSizeAdjustment")
        conn.Close()
        Dim dt As DataTable = ds.Tables("GridViewColumnSizeAdjustment")

        Dim rw As DataRow

        Dim columnSize As Integer = 20

        For Each rw In dt.Rows
            columnSize = rw("GridViewColumnSize")
        Next
        Return columnSize

    End Function

    Public Function SaveColumnWidth(ByVal screen As Integer, ByVal gridViewControl As Integer, ByVal columnWidth As Integer) As Integer

        Dim query As String = "SELECT * FROM GridViewColumnSizeAdjustment Where FormReferenceNumber='" & screen & "' AND GridViewControlNumber='" & gridViewControl & "'"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "GridViewColumnSizeAdjustment")
        conn.Close()
        Dim dt As DataTable = ds.Tables("GridViewColumnSizeAdjustment")

        Dim rw As DataRow



        For Each rw In dt.Rows
            rw.BeginEdit()
            rw("GridViewColumnSize") = columnWidth
            rw.EndEdit()
        Next

        Dim objCommandBuilder As New SqlCommandBuilder(da)
        da.Update(ds, "GridViewColumnSizeAdjustment")
        Return columnWidth

    End Function
End Class

'Stores the path of a Calendar
'Returns the path location of a Calendar
'Makes use of Factory Design Pattern
Public Interface IfileStoragePath
    Function path(newPath As String, newModule As String) As String

End Interface

Public Class saveStorageModules
    Implements IfileStoragePath
    Public Function path(newPath As String, scheduleType As String) As String Implements IfileStoragePath.path

        Dim selectedModule As IStorageModule = Nothing

        Select Case (scheduleType)
            Case "student"
                selectedModule = New storeStudentScheduleFilePath
                selectedModule.storeFilePath(newPath.Trim())
                Exit Select
            Case "office"
                selectedModule = New storeOfficeScheduleFilePath
                selectedModule.storeFilePath(newPath.Trim())
                Exit Select
            Case Else

                Exit Select
        End Select

        Return newPath
    End Function
End Class
Public Class returnStorageModules
    Implements IfileStoragePath
    Public Function path(newPath As String, scheduleType As String) As String Implements IfileStoragePath.path
        Dim selectedModule As IreturnStorageModule = Nothing
        Dim defaultLocation As String = String.Empty
        Select Case (scheduleType)
            Case "student"
                selectedModule = New StudentCalendarFilelocation
                defaultLocation = selectedModule.locateFilePath(newPath.Trim())
                Exit Select
            Case "office"
                selectedModule = New OfficeCalendarFilelocation
                defaultLocation = selectedModule.locateFilePath(newPath.Trim())
                Exit Select
            Case Else

                Exit Select
        End Select

        Return defaultLocation
    End Function
End Class
Public Interface IStorageModule
    Function storeFilePath(path As String)

End Interface

Public Interface IreturnStorageModule
    Function locateFilePath(path As String) As String
End Interface

Class storeStudentScheduleFilePath
    Implements IstorageModule
    Dim connectionString As Object
    Public Sub New()
        Try

            connectionString = ConfigurationManager.ConnectionStrings("Familyliteracy").ConnectionString
        Catch ex As Exception
            MsgBox("Cannot connect to Datasource")
        End Try
    End Sub
    Function storeFilePath(path As String) Implements IStorageModule.storeFilePath
        Dim query As String = "Update DefaultScheduleFileLocations SET [Student_Schedule_Location]=@storagePath,[StorageDate]=@storageDate"

        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)


        cmd.Parameters.AddWithValue("@storagePath", path)
        cmd.Parameters.AddWithValue("@storageDate", Now)

        Dim updated As Integer = 0
        conn.Open()
        updated = cmd.ExecuteNonQuery()
        conn.Close()

        Return Nothing
    End Function
End Class

Public Class storeOfficeScheduleFilePath
    Implements IstorageModule
    Dim connectionString As Object
    Public Sub New()
        Try

            connectionString = ConfigurationManager.ConnectionStrings("Familyliteracy").ConnectionString
        Catch ex As Exception
            MsgBox("Cannot connect to Datasource")
        End Try
    End Sub
    Function storeFilePath(path As String) Implements IstorageModule.storeFilePath
        Dim query As String = "Update DefaultScheduleFileLocations SET [Office_Schedule_Location]=@storagePath,[StorageDate]=@storageDate"

        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)


        cmd.Parameters.AddWithValue("@storagePath", path)
        cmd.Parameters.AddWithValue("@storageDate", Now)

        Dim updated As Integer = 0
        conn.Open()
        updated = cmd.ExecuteNonQuery()
        conn.Close()

        Return Nothing
    End Function

End Class

Public Class StudentCalendarFilelocation
    Implements IreturnStorageModule
    Dim connectionString As Object
    Public Sub New()
        Try

            connectionString = ConfigurationManager.ConnectionStrings("Familyliteracy").ConnectionString
        Catch ex As Exception
            MsgBox("Cannot connect to Datasource")
        End Try
    End Sub

    Function locateFilePath(path As String) As String Implements IreturnStorageModule.locateFilePath
        Dim query As String = "SELECT * FROM DefaultScheduleFileLocations "
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "DefaultScheduleFileLocations")
        conn.Close()
        Dim dt As DataTable = ds.Tables("DefaultScheduleFileLocations")

        Dim rw As DataRow

        Dim defaultPath As String = String.Empty

        For Each rw In dt.Rows
            defaultPath = rw("Student_Schedule_Location")
        Next

        Dim objCommandBuilder As New SqlCommandBuilder(da)
        da.Update(ds, "DefaultScheduleFileLocations")


        Return defaultPath
    End Function


End Class


Public Class OfficeCalendarFilelocation
    Implements IreturnStorageModule
    Dim connectionString As Object
    Public Sub New()
        Try

            connectionString = ConfigurationManager.ConnectionStrings("Familyliteracy").ConnectionString
        Catch ex As Exception
            MsgBox("Cannot connect to Datasource")
        End Try
    End Sub

    Function locateFilePath(path As String) As String Implements IreturnStorageModule.locateFilePath
        Dim query As String = "SELECT * FROM DefaultScheduleFileLocations"
        Dim conn As New SqlConnection(connectionString)
        Dim cmd As New SqlCommand(query, conn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        conn.Open()
        da.Fill(ds, "DefaultScheduleFileLocations")
        conn.Close()
        Dim dt As DataTable = ds.Tables("DefaultScheduleFileLocations")

        Dim rw As DataRow

        Dim defaultPath As String = String.Empty

        For Each rw In dt.Rows
            defaultPath = rw("Office_Schedule_Location")
        Next

        Dim objCommandBuilder As New SqlCommandBuilder(da)
        da.Update(ds, "DefaultScheduleFileLocations")


        Return defaultPath
    End Function


End Class