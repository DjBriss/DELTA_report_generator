Imports System.IO
Imports Microsoft.Office.Interop.Excel

Public Class Form1
    'So basically the user chooses the Date, Then file location, Then presses the button To start the process, cause that how I figured it out
    '    The process starts by making a New big array, Then fills it With the existing report, Then eventually overwrites the file And add a New column
    '    based ReadOnly the numbers taken from a parsed csv file that I add To the big array
    Dim file_location As String = ""

    Dim day_num As String = "1"

    Dim month_num As String = ""

    Dim year_num As String = ""

    Dim day As Integer = CInt(day_num)

    Dim excel As Application = New Application

    Public Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Console.WriteLine(ListBox1.SelectedIndex)

        If ListBox1.SelectedIndex = 0 Then
            file_location = "C:\Documents and Settings\operator\My Documents\DELTA\reports\"
        End If
        If ListBox1.SelectedIndex = 1 Then

            file_location = "C:\Users\Jbrisson\Documents\Visual Studio 2015\Projects\DELTA_report_generator-master\DELTA_report_generator-master\WindowsApplication2\WindowsApplication2\bin\Debug\"
        End If

        Console.WriteLine(file_location)
    End Sub

    Public Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        'when the date is picked, the values are parsed, and stored into the variables day_num, month_num, year_num for later use

        Dim Datestring As String = DateTimePicker1.Value

        Dim DATE_array As String() = Strings.Split(Datestring, "/")
        Dim index As Integer
        While index < 3
            Console.WriteLine(DATE_array(index))
            index += 1
        End While

        day_num = DATE_array(0)

        month_num = MonthName(CInt(DATE_array(1)))

        year_num = (Strings.Split(DATE_array(2), " ")(0))
        '^ puts the year only as the variable, because the string also contains the time
        Console.WriteLine(month_num)
        Console.WriteLine(year_num)


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'starts the process

        Dim daily_report(1000, 10) As Object

        daily_report = excel_into_array(file_location, daily_report, day_num & ".xlsx", 1)

        Console.WriteLine("successcscss")

        For a As Integer = 1 To 6

            Dim monthly_report(100, 50) As Object

            monthly_report = excel_into_array(file_location, monthly_report, "monthly_report.xlsx", a)
            add_numbers(monthly_report, daily_report)
            array_into_excel(monthly_report, file_location, "monthly_report.xlsx", a)
            Console.WriteLine(CStr(a))
        Next
        excel_quit()
    End Sub

    Private Function excel_into_array(filelocation As String, report(,) As Object, file_name As String, i As Integer)

        Dim filename As String = (filelocation & month_num & year_num & "\" & file_name)

        Console.WriteLine(filename)

        Dim workbook1 As Workbook = excel.Workbooks.Open(filename)

        Dim sheet As Worksheet = workbook1.Sheets(i)

        Dim r As Range = sheet.UsedRange

        report = r.Value(XlRangeValueDataType.xlRangeValueDefault)

        If report IsNot Nothing Then

            Console.WriteLine("Length: {0}", report.Length)

            ' Get bounds of the array.
            Dim bound0 As Integer = report.GetUpperBound(0)
            Dim bound1 As Integer = report.GetUpperBound(1)

            Console.WriteLine("Dimension 0: {0}", bound0)
            Console.WriteLine("Dimension 1: {0}", bound1)

            ' Loop over all elements.
            For j As Integer = 1 To bound0
                For x As Integer = 1 To bound1
                    Dim s1 As String = report(j, x)
                    Console.Write(s1 & " ")

                Next
                Console.WriteLine()
            Next
        End If
        Return report
        workbook1.Close(True, )
    End Function

    Private Function add_numbers(monthly_report(,) As Object, daily_report(,) As Object)

        Dim bound0 As Integer = daily_report.GetUpperBound(0)

        For index1 As Integer = 1 To 46

            For index2 As Integer = 1 To bound0

                'Console.WriteLine(daily_report(index2, 1))
                'Console.WriteLine(daily_report(index2, 5))
                'Console.WriteLine(monthly_report(index1, 2))

                If monthly_report(index1, 2) = daily_report(index2, 1) Then
                    Dim location As Integer = day + 2

                    monthly_report(index1, (location)) = daily_report(index2, 5)
                    'Console.WriteLine(monthly_report(index1, 2))
                    'Console.WriteLine(daily_report(index2, 5))
                    'Console.WriteLine(monthly_report(index1, (day + 2)))
                    'Console.WriteLine("a")

                End If
                index2 += 1

            Next
            index1 += 1
            Next

    End Function

    Private Function array_into_excel(report(,) As Object, filelocation As String, file_name As String, i As Integer)

        Dim filename As String = (filelocation & month_num & year_num & "\" & file_name)

        Console.WriteLine(filename)

        Dim workbook2 As Workbook = excel.Workbooks.Open(filename)

        Dim sheet1 As Worksheet = workbook2.Sheets(i)


        sheet1.Range("A1:AJ50").Value = report

        workbook2.Close(True,)

    End Function

    Sub excel_quit()
        excel.Quit()
    End Sub


End Class