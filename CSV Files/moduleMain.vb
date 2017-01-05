Imports System.Runtime.CompilerServices
Imports CSV_Files.Exporting
Imports CSV_Files.Importing.CSV

Module moduleMain
    Dim pth As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "data.csv")

    Sub Main()
        Dim nums = Enumerable.Range(1, 100000)
        Dim sw As New Stopwatch
        Task.Run(Async Function()

                     sw.Start()
                     Await nums.SaveCSV
                     sw.Stop()

                     Console.WriteLine($"It took {sw.Elapsed} to write the csv.")
                 End Function).Wait()

        Dim imp As New NumberDataImporter(pth)

        Task.Run(Async Function()
                     Dim sw2 As New Stopwatch
                     sw2.Start()
                     Dim nd = Await imp.Import
                     sw2.Stop()
                     Console.WriteLine($"It took {sw2.Elapsed} to read {nd.Count:N0} records from the csv.")
                     'Stop
                     'examine nd  See that we have number data objects in it
                 End Function).Wait()
        Console.WriteLine($"All fields for number data objects {NumberDataFields.AllFields}")
        Console.WriteLine($"Required fields for number data objects {NumberDataFields.RequiredFields}")
        Console.WriteLine("Press any key to quit...")

        Console.ReadKey()

    End Sub

    <Extension>
    Function SaveCSV(nums As IEnumerable(Of Integer)) As Task
        Return Task.Run(Sub()
                            Using str As New IO.FileStream(pth, IO.FileMode.Create)
                                'Write required fields only
                                nums.ToCsv(str,
                                           {
                                           NumberDataFields.Number,
                                           NumberDataFields.EvenOdd},
                                           Function(n) n,
                                           Function(n) If(n Mod 2 = 0, NumberData.EvenOdd.Even, NumberData.EvenOdd.Odd)
                                          )
                                'Write all data
                                'nums.ToCsv(str,
                                '           {
                                '           NumberDataFields.Number,
                                '           NumberDataFields.EvenOdd,
                                '           NumberDataFields.SquareRoot
                                '           },
                                '           Function(n) n,
                                '           Function(n) If(n Mod 2 = 0, NumberData.EvenOdd.Even, NumberData.EvenOdd.Odd),
                                '           Function(n) n ^ 0.5)

                            End Using
                        End Sub)

    End Function

End Module
