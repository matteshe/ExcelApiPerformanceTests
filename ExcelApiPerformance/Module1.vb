'Copyright (c) 2018 Mathias.Herkt@sqs.com
Imports NetOffice.ExcelApi

Module Module1

    Sub Main()
        Dim netOfficeThread = New Threading.Thread(AddressOf NetOfficeWay)
        Dim lateBindingThread = New Threading.Thread(AddressOf LateBindingWay)

        netOfficeThread.Start()
        lateBindingThread.Start()

        While netOfficeThread.IsAlive OrElse lateBindingThread.IsAlive
            ' wait
            Threading.Thread.Sleep(200)
        End While

        Console.WriteLine("press any key to close")
        Console.ReadKey()
    End Sub

    Private Sub LateBindingWay()
        Dim excelApp As Object = CreateObject("Excel.Application")
        excelApp.DisplayAlerts = False

        Dim fi As New IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim fs As System.IO.FileStream = New System.IO.FileStream(fi.DirectoryName & "\DataSheetLateBinding.xlsx", IO.FileMode.Create)

        fs.Write(My.Resources.DataSheet, 0, My.Resources.DataSheet.Length)
        fs.Close()

        Dim book = Nothing
        Dim newBook = Nothing
        Try
            book = excelApp.Workbooks.Open(fs.Name)
            Dim inputSheet = book.Sheets().Item(1)
            Dim usedRange = inputSheet.UsedRange
            Dim cols As Integer = usedRange.Columns.Count / 2
            Dim rows As Integer = usedRange.Rows.Count / 2
            Console.WriteLine(String.Format("colums:{0}, rows:{1}", usedRange.Columns.Count, usedRange.Rows.Count))

            newBook = excelApp.Workbooks.Add()
            Dim newInputSheet = newBook.Sheets.Add()
            newInputSheet.Name = "InputData1"
            Dim readThread As New Threading.Thread(AddressOf PrintSeconds)
            readThread.Start("LateBindingWay 1/3")

            Console.WriteLine("with inputsheet")
            For colIdx As Integer = 1 To cols
                For rowIdx As Integer = 1 To rows
                    newInputSheet.Cells(colIdx, rowIdx).Value = inputSheet.Cells(colIdx, rowIdx).Value
                    newInputSheet.Cells(colIdx, rowIdx).FormulaLocal = inputSheet.Cells(colIdx, rowIdx).FormulaLocal
                Next
            Next
            readThread.Abort()

            Console.WriteLine("with used range")
            newInputSheet = newBook.Sheets.Add()
            newInputSheet.Name = "InputData2"
            newInputSheet.Select()
            readThread = New Threading.Thread(AddressOf PrintSeconds)
            readThread.Start("LateBindingWay 2/3")
            Dim fillRange = newInputSheet.Range(newInputSheet.Cells(1, 1), newInputSheet.Cells(cols, rows))
            For colIdx As Integer = 1 To cols
                For rowIdx As Integer = 1 To rows
                    fillRange.Cells(colIdx, rowIdx).Value = usedRange.Cells(colIdx, rowIdx).Value
                    fillRange.Cells(colIdx, rowIdx).FormulaLocal = usedRange.Cells(colIdx, rowIdx).FormulaLocal
                Next
            Next
            readThread.Abort()

            Console.WriteLine("with cell once")
            newInputSheet = newBook.Sheets.Add()
            newInputSheet.Name = "InputData3"
            newInputSheet.Select()
            readThread = New Threading.Thread(AddressOf PrintSeconds)
            readThread.Start("LateBindingWay 3/3")
            For colIdx As Integer = 1 To cols
                For rowIdx As Integer = 1 To rows
                    Dim cell = inputSheet.Cells(colIdx, rowIdx)
                    Dim newCell = newInputSheet.Cells(colIdx, rowIdx)

                    newCell.Value = cell.Value
                    newCell.FormulaLocal = cell.FormulaLocal
                Next
            Next
            readThread.Abort()

            newBook.SaveAs(fs.Name.Replace(".xlsx", "new.xlsx"))
            Console.WriteLine("done")
        Finally
            If (book IsNot Nothing) Then
                book.Close()
            End If

            If (newBook IsNot Nothing) Then
                newBook.Close()
            End If
            excelApp.Quit()
            IO.File.Delete(fs.Name)
        End Try
    End Sub

    Private Sub NetOfficeWay()
        Dim excelApp As New Application()
        excelApp.DisplayAlerts = False
        Dim fi As New IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim fs As System.IO.FileStream = New System.IO.FileStream(fi.DirectoryName & "\DataSheetNetOffice.xlsx", IO.FileMode.Create)

        fs.Write(My.Resources.DataSheet, 0, My.Resources.DataSheet.Length)
        fs.Close()
        Dim book As Workbook = Nothing
        Dim newBook As Workbook = Nothing
        Try
            book = excelApp.Workbooks.Open(fs.Name)
            Dim inputSheet As Worksheet = book.Sheets().Item(1)
            Dim usedRange As Range = inputSheet.UsedRange
            Dim cols As Integer = usedRange.Columns.Count / 2
            Dim rows As Integer = usedRange.Rows.Count / 2
            Console.WriteLine(String.Format("colums:{0}, rows:{1}", usedRange.Columns.Count, usedRange.Rows.Count))

            newBook = excelApp.Workbooks.Add()
            Dim newInputSheet As Worksheet = newBook.Sheets.Add()
            newInputSheet.Name = "InputData1"
            Dim readThread As New Threading.Thread(AddressOf PrintSeconds)
            readThread.Start("NetOfficeWay 1/3")

            Console.WriteLine("with inputsheet")
            For colIdx As Integer = 1 To cols
                For rowIdx As Integer = 1 To rows
                    newInputSheet.Cells(colIdx, rowIdx).Value = inputSheet.Cells(colIdx, rowIdx).Value
                    newInputSheet.Cells(colIdx, rowIdx).FormulaLocal = inputSheet.Cells(colIdx, rowIdx).FormulaLocal
                Next
            Next
            readThread.Abort()

            Console.WriteLine("with used range")
            newInputSheet = newBook.Sheets.Add()
            newInputSheet.Name = "InputData2"
            newInputSheet.Select()
            readThread = New Threading.Thread(AddressOf PrintSeconds)
            readThread.Start("NetOfficeWay 2/3")
            Dim fillRange As Range = newInputSheet.Range(newInputSheet.Cells(1, 1), newInputSheet.Cells(cols, rows))
            For colIdx As Integer = 1 To cols
                For rowIdx As Integer = 1 To rows
                    fillRange.Cells(colIdx, rowIdx).Value = usedRange.Cells(colIdx, rowIdx).Value
                    fillRange.Cells(colIdx, rowIdx).FormulaLocal = usedRange.Cells(colIdx, rowIdx).FormulaLocal
                Next
            Next
            readThread.Abort()

            Console.WriteLine("with cell once")
            newInputSheet = newBook.Sheets.Add()
            newInputSheet.Name = "InputData3"
            newInputSheet.Select()
            readThread = New Threading.Thread(AddressOf PrintSeconds)
            readThread.Start("NetOfficeWay 3/3")
            For colIdx As Integer = 1 To cols
                For rowIdx As Integer = 1 To rows
                    Dim cell As Range = inputSheet.Cells(colIdx, rowIdx)
                    Dim newCell As Range = newInputSheet.Cells(colIdx, rowIdx)

                    newCell.Value = cell.Value
                    newCell.FormulaLocal = cell.FormulaLocal
                Next
            Next
            readThread.Abort()

            newBook.SaveAs(fs.Name.Replace(".xlsx", "new.xlsx"))
            Console.WriteLine("done")
        Finally
            If (book IsNot Nothing) Then
                book.Close()
            End If

            If (newBook IsNot Nothing) Then
                newBook.Close()
            End If
            excelApp.Dispose()
            IO.File.Delete(fs.Name)
        End Try
    End Sub

    Public Sub PrintSeconds(ByVal caller As String)
        Dim sec As Integer = 1
        While True
            Console.WriteLine(caller & ": " & sec)
            Threading.Thread.Sleep(1000)
            sec += 1
        End While
    End Sub
End Module
