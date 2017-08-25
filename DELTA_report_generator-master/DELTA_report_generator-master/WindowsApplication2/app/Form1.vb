Imports System.IO
Imports Microsoft.Office.Interop.Excel

Public Class Form1
    'So basically the user chooses the Date, Then file location, Then presses the button To start the process, cause that how I figured it out
    '    The process starts by making a New big array, Then fills it With the existing report, Then eventually overwrites the file And add a New column
    '    based ReadOnly the numbers taken from a parsed csv file that I add To the big array


    ''initialize the necessary variables
    Dim file_location As String = ""

    Dim day_num As String = "1"

    Dim month_num As String = ""

    Dim year_num As String = ""

    Dim day As Integer = 0

    Dim extension As String = ".xlsx"

    Dim excel As Application = New Application

    'sets the extension
    Public Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged

        If ListBox2.SelectedIndex = 0 Then
            extension = ".xls"

        ElseIf ListBox2.SelectedIndex = 1 Then
            extension = ".xlsx"

        End If

        MsgBox(extension)
    End Sub

    'sets the location
    Public Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        If ListBox1.SelectedIndex = 0 Then
            file_location = "C:\Documents and Settings\operator\My Documents\DELTA_reports\"
        End If

        If ListBox1.SelectedIndex = 1 Then
            file_location = "C:\Users\Jbrisson\Documents\Visual Studio 2015\Projects\DELTA_report_generator-master\DELTA_report_generator-master\app\app\bin\Debug\"
        End If

        If ListBox1.SelectedIndex = 2 Then
            file_location = "C:\Documents and Settings\Maintenance\My Documents\DELTA_reports\"
        End If

        MsgBox(file_location)
    End Sub

    'assigns values to the date variables
    Public Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        'when the date is picked, the values are parsed, and stored into the variables day_num, month_num, year_num for later use

        Dim Datestring As String = DateTimePicker1.Value
        Dim DATE_array As String() = Strings.Split(Datestring, "/")
        Dim index As Integer
        While index < 3
            index += 1
        End While

        day_num = DATE_array(0)
        month_num = MonthName(CInt(DATE_array(1)))
        year_num = (Strings.Split(DATE_array(2), " ")(0))

        '^ puts the year only as the variable, because the string also contains the time

        day = CInt(day_num)

    End Sub

    'starts the process
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        excel_quit()

        Dim monthly_report(100, 60) As Object
        Dim daily_report(1000, 10) As Object
        Dim numbersfilename As String = file_location & month_num & year_num & "\" & day_num & extension
        Dim reportfilename As String = file_location & month_num & year_num & "\" & "monthly_report" & extension
        Dim count As Integer = 0

        If File.Exists(numbersfilename) = False Or File.Exists(reportfilename) = False Then
            MsgBox("file not found")
            Exit Sub
        End If

        daily_report = excel_into_array(daily_report, numbersfilename, 1)
        count = number_of_sheets(reportfilename)
        MsgBox(count)

        For a As Integer = 1 To count

            monthly_report = excel_into_array(monthly_report, reportfilename, a)
            add_numbers(monthly_report, daily_report)
            array_into_excel(monthly_report, reportfilename, a)

        Next

        excel_quit()

        MsgBox("The numbers of report " & day_num & " have been added succesfully")

    End Sub

    Private Function excel_into_array(report(,) As Object, filename As String, i As Integer)

        If File.Exists(filename) = False Then
            MsgBox("file not found")
            Exit Function
        Else

            Dim workbook1 As Workbook = excel.Workbooks.Open(filename)
            Dim sheet As Worksheet = workbook1.Sheets(i)
            Dim r As Range = sheet.UsedRange

            report = r.Value(XlRangeValueDataType.xlRangeValueDefault)

        End If
        Return report

    End Function

    Private Function number_of_sheets(filename As String)

        If File.Exists(filename) = False Then
            MsgBox("file not found")
            Exit Function
        Else

            Dim workbook0 As Workbook = excel.Workbooks.Open(filename)
            Return workbook0.Sheets.Count

        End If

    End Function

    Private Function add_numbers(monthly_report(,) As Object, daily_report(,) As Object)

        If daily_report Is Nothing Then
            MsgBox("No Numbers To add")
            Exit Function
        End If

        Dim bound0 As Object = daily_report.GetUpperBound(0)
        Dim bound1 As Object = monthly_report.GetUpperBound(0)
        Dim location As Integer = day + 3

        For index1 As Integer = 1 To bound1

            For index2 As Integer = 1 To bound0

                If System.String.Equals(monthly_report(index1, 3), daily_report(index2, 1)) = True And Not String.Equals(monthly_report(index1, 3), Nothing) Then

                    monthly_report(index1, location) = daily_report(index2, 5)

                End If

            Next

        Next

    End Function

    Private Function array_into_excel(report(,) As Object, filelocation As String, i As Integer)

        If report Is Nothing Then
            MsgBox("No Numbers To add")
            Exit Function
        End If

        Dim workbook2 As Workbook = excel.Workbooks.Open(filelocation)
        Dim sheet1 As Worksheet = workbook2.Sheets(i)

        sheet1.UsedRange.Value = report

        workbook2.Close(True,)

    End Function

    Sub excel_quit()
        excel.Quit()
        MsgBox("excel is dead")
    End Sub

End Class