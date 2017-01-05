Namespace Importing.CSV
    Public Class NumberDataImporter
        Inherits Importing.CSV.CSVImporter(Of NumberData, NumberData)

        Public Sub New(FilePath As String)
            MyBase.New(FilePath)
        End Sub

        Protected Overrides Function ConvertRaw(value As NumberData) As NumberData
            Return value
        End Function

        Protected Overrides Function MergeResults(SplitResults As IEnumerable(Of IEnumerable(Of NumberData))) As IEnumerable(Of NumberData)
            Throw New NotImplementedException()
        End Function

        Protected Overrides Function ParseLine(Line() As String, Keys() As String) As NumberData
            Dim tmp As New NumberData With {
                .Number = CInt(Line(Array.IndexOf(Keys, NumberDataFields.Number))),
                .Even = CType(
                [Enum].Parse(GetType(NumberData.EvenOdd),
                             Line(Array.IndexOf(Keys, NumberDataFields.EvenOdd))),
                       NumberData.EvenOdd)}

            With tmp
                If Array.IndexOf(Keys, NumberDataFields.SquareRoot) > -1 Then
                    .Sqrt = CDbl(Line(Array.IndexOf(Keys, NumberDataFields.SquareRoot)))
                End If
            End With
            Return tmp
        End Function

    End Class
End Namespace
