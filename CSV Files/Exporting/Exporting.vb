Imports System.Runtime.CompilerServices
Imports System.IO
Imports System.Text
Imports System.Linq.Expressions

Namespace Exporting
    Module CSV
        <Extension()>
        Public Sub ToCsv(Of T)(ByVal objects As IEnumerable(Of T),
                                ByVal outputStream As Stream, ByVal encoding As Encoding,
                                ByVal columnSeparator As Char, ByVal lineTerminator As String,
                                ByVal encapsulationCharacter As Char, ByVal autoGenerateColumns As Boolean,
                                ByVal columnHeaders() As String,
                                ByVal ParamArray outputValues() As Expression(Of Func(Of T, Object)))
            Using writer As StreamWriter = New StreamWriter(outputStream, encoding, 1024, True)

                WriteColumnHeaders(writer, columnSeparator, lineTerminator,
                               encapsulationCharacter, autoGenerateColumns, columnHeaders, outputValues)
                WriteData(objects, writer, columnSeparator, lineTerminator,
                          encapsulationCharacter, outputValues)
                writer.Flush()
            End Using

        End Sub

        <Extension()>
        Public Sub ToCsv(Of T)(ByVal objects As IEnumerable(Of T),
                                ByVal outputStream As Stream, ByVal columnHeaders() As String,
                                ByVal ParamArray outputValues() _
                                   As Expression(Of Func(Of T, Object)))
            objects.ToCsv(outputStream, Encoding.Default, ","c, Environment.NewLine, ControlChars.Quote,
                          False, columnHeaders, outputValues)
        End Sub

        Private Sub WriteColumnHeaders(Of T)(ByVal writer As StreamWriter, ByVal columnSeparator As Char,
                                              ByVal lineTerminator As String, ByVal encapsulationCharacter As Char,
                                              ByVal autoGenerateColumns As Boolean, ByVal columnHeaders() As String,
                                              ByVal ParamArray outputValues() As Expression(Of Func(Of T, Object)))
            If autoGenerateColumns Then
                For i As Integer = 0 To outputValues.Length - 1
                    Dim expression As Expression(Of Func(Of T, Object)) = outputValues(i)
                    Dim columnHeader As String
                    If TypeOf expression.Body Is MemberExpression Then
                        Dim body As MemberExpression = DirectCast(expression.Body, MemberExpression)
                        columnHeader = body.Member.Name
                    ElseIf TypeOf expression.Body Is UnaryExpression Then
                        Dim body As UnaryExpression = DirectCast(expression.Body, UnaryExpression)
                        If TypeOf body.Operand Is MemberExpression Then
                            Dim operand As MemberExpression = DirectCast(body.Operand, MemberExpression)
                            columnHeader = operand.Member.Name
                        Else
                            columnHeader = body.ToString()
                        End If
                    Else
                        columnHeader = expression.Body.ToString()
                    End If
                    If i < outputValues.Length - 1 Then
                        writer.Write("{0}{1}",
                                     columnHeader.EncapsulateIfRequired(columnSeparator, encapsulationCharacter),
                                     columnSeparator)
                    Else
                        writer.Write(columnHeader.EncapsulateIfRequired(columnSeparator, encapsulationCharacter))
                    End If
                Next
                writer.Write(lineTerminator)
            Else
                If Not columnHeaders Is Nothing And columnHeaders.Length > 0 Then
                    If columnHeaders.Length = outputValues.Length Then
                        For i As Integer = 0 To columnHeaders.Length - 1
                            If i < columnHeaders.Length - 1 Then
                                writer.Write(String.Format("{0}{1}",
                                                           columnHeaders(i).EncapsulateIfRequired(columnSeparator,
                                                                                                  encapsulationCharacter),
                                                           columnSeparator))
                            Else
                                writer.Write(columnHeaders(i).EncapsulateIfRequired(columnSeparator,
                                                                                    encapsulationCharacter))
                            End If
                        Next
                        writer.Write(lineTerminator)
                    Else
                        Throw _
                            New ArgumentException(
                                "The number of column headers does not match the number of output values.")
                    End If
                End If
            End If
        End Sub

        Private Sub WriteData(Of T)(ByVal objects As IEnumerable(Of T), ByVal writer As StreamWriter,
                                     ByVal columnSeparator As Char, ByVal lineTerminator As String,
                                     ByVal encapsulationCharacter As Char,
                                     ByVal ParamArray outputValues() As Expression(Of Func(Of T, Object)))
            Debug.WriteLine(objects.Count)

            Dim outputs(outputValues.Length) As Func(Of T, Object)
            For i As Integer = 0 To outputValues.Length - 1
                outputs(i) = outputValues(i).Compile()
            Next
            For Each obj As T In objects
                If Not obj Is Nothing Then
                    For i As Integer = 0 To outputValues.Length - 1
                        Dim valueFunc As Func(Of T, Object) = outputs(i)
                        Dim value As Object = valueFunc(obj)
                        If Not value Is Nothing Then
                            Dim valueString As String = value.ToString()
                            writer.Write(valueString.EncapsulateIfRequired(columnSeparator, encapsulationCharacter))
                        End If
                        If i < outputValues.Length - 1 Then
                            writer.Write(columnSeparator)
                        End If
                    Next
                    writer.Write(lineTerminator)
                End If
            Next
        End Sub

        <Extension()>
        Private Function EncapsulateIfRequired(ByVal theString As String, ByVal columnSeparator As Char,
                                               ByVal encapsulationCharacter As Char) As String
            If theString.Contains(columnSeparator) Then
                If theString.Contains(encapsulationCharacter) Then
                    theString = theString.Replace(encapsulationCharacter.ToString(),
                                                  New String(encapsulationCharacter, 2))
                End If
                Return String.Format("{1}{0}{1}", theString, encapsulationCharacter)
            Else
                Return theString
            End If
        End Function
    End Module
End Namespace
