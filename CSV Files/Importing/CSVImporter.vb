Imports Microsoft.VisualBasic.FileIO
Imports System.Threading.Tasks

Namespace Importing.CSV
    Public MustInherit Class CSVImporter(Of T, RawType)
        'RawType is a data model class that is represented by the data in the CSV
        'T is the type you want to use in the program.

        'T can either inherit from RawType or you can override the ConvertRaw function
        'to generate a T object from a RawType object.
        Inherits FileImporter(Of T)
        Public Sub New(FilePath As String)
            MyBase.New(FilePath)
        End Sub
        Public Sub New(FilePath As String, CancelSource As Threading.CancellationTokenSource)
            MyBase.New(FilePath, CancelSource)
        End Sub
        Protected Function ParseCSV(ByRef Line As Long) As List(Of T)
            Dim splitResults As New List(Of List(Of RawType))

            splitResults.Add(New List(Of RawType))
            Dim currentResults = splitResults(0)
            Dim fields As String()
            Dim delimiter() As String = {","}
            OnImportProgress(New ImportProgressEventArgs("Parsing file", 0.0))
            Using parser As New TextFieldParser(FilePath)
                parser.SetDelimiters(delimiter)

                parser.HasFieldsEnclosedInQuotes = True
                Dim keys() As String = Nothing
                While Not parser.EndOfData
                    Line = parser.LineNumber
                    If CancelToken.IsCancellationRequested Then Exit While
                    ' Read in the fields for the current line

                    If keys Is Nothing Then

                        Dim keysLine As String = parser.ReadLine()
                        keysLine = keysLine.Replace(ControlChars.Quote, "")

                        keys = keysLine.Split(","c)

                        Continue While

                    Else

                        fields = parser.ReadFields()

                        If fields(0).StartsWith("#") Then

                            splitResults.Add(New List(Of RawType))
                            currentResults = splitResults(splitResults.Count - 1)
                            Continue While
                        End If

                    End If
                    If fields.Count > 0 AndAlso Not fields.
                        All(Function(f) String.IsNullOrEmpty((f & "").Trim)) Then 'ignore blank lines
                        'or lines with just ,,,,
                        Dim tmp As RawType = ParseLine(fields, keys)
                        If tmp IsNot Nothing Then currentResults.Add(tmp)
                    End If
                End While
            End Using
            OnImportProgress(New ImportProgressEventArgs("Merge and convert objects.", 0.75))

            If splitResults.Count = 1 Then
                _Result = ConvertRaw(splitResults(0))
                Return _Result.ToList
            End If
            Dim results As IEnumerable(Of RawType) = MergeResults(splitResults)
            _Result = ConvertRaw(results)
            OnImportProgress(New ImportProgressEventArgs(String.Format("File parsing complete for {0}", FilePath), 1.0))
            Return _Result.ToList
        End Function
        'Used to merge multiple ienumerable(of ienumerable(of RawType))
        'Think of this as if you had multiple files merged into one and you put a # to mark where one file stopped.
        'Useful for when you have same data but different sections that you want to merge together
        Protected MustOverride Function MergeResults(SplitResults As IEnumerable(Of IEnumerable(Of RawType))) As IEnumerable(Of RawType)
        'Override to tell the importer how to get a value of type T from a RawType
        Protected MustOverride Function ConvertRaw(value As RawType) As T
        Protected Overridable Function ConvertRaw(values As IEnumerable(Of RawType)) As IEnumerable(Of T)
            Dim tmp As RawType
            If TypeOf tmp Is T Then Return values.Cast(Of T)
            Return From r In values Select ConvertRaw(r)
        End Function
        'Parse a line of the CSV and create a RawType object from it.
        Protected MustOverride Function ParseLine(Line As String(), Keys As String()) As RawType
        'helper function when importing timespans.  Since timespans are hard to parse
        'yet a Date can easily store a timespan inside it.
        Protected Function GetTime(fields As String(), key As String, keys As String()) As TimeSpan?
            If Not String.IsNullOrEmpty(fields(Array.IndexOf(keys, key))) Then
                Dim tmpDate As Date
                If Date.TryParse(fields(Array.IndexOf(keys, key)), tmpDate) Then Return tmpDate.TimeOfDay
            End If
            Return Nothing
        End Function

        Public Overrides Function Import() As Task(Of IEnumerable(Of T))
            SetupImport()

            Dim tsk As New Task(Of IEnumerable(Of T))(Function()
                                                          Dim results As List(Of T)
                                                          Dim line As Long = -1
                                                          Try
                                                              results = ParseCSV(line)

                                                          Catch ex As IndexOutOfRangeException
                                                              'Check the CSV file headers against the expected headers
                                                              Cancel()
                                                              Dim impTYpe As Type = Me.GetType
                                                              ex.Data.Add("Action", "Parsing with " & impTYpe.Name)
                                                              ex.Data.Add("LineNumber", line)
                                                              Throw

                                                          Catch ex As Exception
                                                              Cancel()
                                                              Dim impTYpe As Type = Me.GetType
                                                              ex.Data.Add("Action", "Parsing with " & impTYpe.Name)
                                                              ex.Data.Add("LineNumber", line)
                                                              Throw
                                                          End Try

                                                          Return results
                                                      End Function
                                                  )

            tsk.Start()
            Return tsk
        End Function
    End Class
End Namespace
