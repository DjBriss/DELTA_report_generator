﻿Imports System.IO
Imports Microsoft.Office.Interop.Excel

Public Class Form1
    'So basically the user chooses the Date, Then file location, Then presses the button To start the process, cause that how I figured it out
    '    The process starts by making a New big array, Then fills it With the existing report, Then eventually overwrites the file And add a New column
    '    based ReadOnly the numbers taken from a parsed csv file that I add To the big array
    Dim file_location As String = ""
    Dim day_num As Integer = 1

    Dim month_num As String = ""

    Dim year_num As String = ""

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Console.WriteLine(ListBox1.SelectedIndex)

        Dim file_location As String = ""

        If ListBox1.SelectedIndex = 0 Then
            file_location = "C:\Documents and Settings\operator\My Documents\DELTA\reports\"
        End If
        If ListBox1.SelectedIndex = 1 Then

            file_location = file_location = "C:\Users\Jbrisson\Documents\Visual Studio 2015\Projects\WindowsApplication2\WindowsApplication2\bin\Debug\"
        End If

        Console.WriteLine(file_location)
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        'when the date is picked, the values are parsed, and stored into the variables day_num, month_num, year_num for later use

        Dim Datestring As String = DateTimePicker1.Value

        Dim DATE_array As String() = Strings.Split(Datestring, "/")
        Dim index As Integer
        While index < 3
            Console.WriteLine(DATE_array(index))
            index += 1
        End While

        day_num = CInt(DATE_array(0))

        month_num = MonthName(CInt(DATE_array(1)))

        year_num = (Strings.Split(DATE_array(2), " ")(0))
        '^ puts the year only as the variable, because the string also contains the time
        Console.WriteLine(month_num)
        Console.WriteLine(year_num)


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'starts the process

        Dim old_report(10000, 10000) As String

        read_file(file_location, old_report)
        '^uses the function to fill the array with existing numbers from the report file in appropriate location

        ' write_tofile(file_location, old_report)

    End Sub

    Private Function read_file(ByVal filelocation As String, report As String(,))
        'to get the old report numbers

        Console.WriteLine(filelocation & month_num & year_num & "\" & "report" & month_num & year_num & ".txt")

        Dim filereader_1 As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(filelocation & month_num & year_num & "\" & "report" & month_num & year_num & ".txt")
        'Dim filereader_1 As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader("C:\Users\Matt\Documents\GitHub\DELTA_report_generator\reportJuly2017.txt")
        'starts to open the present version of the report, based on where it is stored

        Dim i As Integer = 0
        Dim j As Integer = 1
        Do
            Do

                report(i, j) = Strings.Split(filereader_1.ReadLine, "    ")(j)

                j += 1
                'inputs the numbers already present in the reports

            Loop Until Strings.Split(filereader_1.ReadLine, "    ")(j) = ""

            i += 1
        Loop Until filereader_1.Peek = -1

        'To test

        Dim filereader_2 As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(filelocation & month_num & year_num & "\" & "report" & month_num & year_num & ".txt")
        Do
            Do

                Console.WriteLine(report(i, j))
                'Console.WriteLine(TAB)

                j += 1
                'inputs the numbers already present in the reports

            Loop Until report(i, j) = ""

            i += 1
        Loop Until filereader_1.Peek = -1

        Return report

    End Function

    'Private Function write_tofile(ByVal filelocation As String, ParamArray report As String()())

    '    'to write the new report, with the added column
    '    Dim filewriter As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(filelocation)

    '    Dim countercolumn As Integer = 0
    '    Dim counterline As Integer = 0
    '    Do
    '        Do
    '            filewriter.WriteLine(report(counterline)(countercolumn) + TAB())
    '            countercolumn += 1
    '            'iterate to write the first line of the report
    '        Loop Until countercolumn = 33

    '        'iterate to write all lines
    '        counterline += 1

    '    Loop Until counterline = 100

    'End Function

    Private Function get_csv_data(csv_file_path As String)
        Dim source() As String = File.ReadAllLines(csv_file_path)
        Dim fields(source.Length)
        Dim index = 0
        For Each element In source
            fields(index) = Split(source(index), Chr(34) + "," + Chr(34))
            index = index + 1
        Next
        Return fields
    End Function

End Class