'Module Module1
'Private Function read_file(ByVal filelocation As String, ParamArray report As String()())
'        'to get the old report numbers


'        Dim filereader_1 As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(filelocation)
'        'staarts to open the present version of the report, based on where it is stored
'        Do
'            report(i)() = Strings.Split(filereader_1.ReadLine, "    ")
'            'inputs the numbers already present in the reports

'        Loop Until filereader_1.Peek = -1

'        Return report

'    End Function

'    Private Function write_tofile(ByVal filelocation As String, ParamArray report As String()())

'        'to write the new report, with the added column
'        Dim filewriter As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(filelocation)

'        Dim countercolumn As Integer = 0
'        Dim counterline As Integer = 0
'        Do
'            Do
'                filewriter.WriteLine(report(counterline)(countercolumn) + TAB())
'                countercolumn += 1
'                'iterate to write the first line of the report
'            Loop Until countercolumn = 33

'            'iterate to write all lines
'            counterline += 1

'        Loop Until counterline = 100

'    End Function
'End Module
