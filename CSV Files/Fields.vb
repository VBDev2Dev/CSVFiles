Option Infer On
'The point of this file is to make it so we can define the fields we will work with when exporting/importing
'We can also mark which fields are necessary and which ones can be left out of the file when importing.
'Did not implement this in this sample but you can use this to tell the user x fields are required for import
'when an exception is thrown while importing.

Imports System.Text
Imports System.Reflection
<AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property, AllowMultiple:=False, Inherited:=False)>
Public Class RequiredFieldAttribute
    Inherits Attribute
    'used to mark fields for parsing that are required
End Class

Public Class Fields
    Protected Function GetFields(Optional OnlyRequired As Boolean = True) As String

        Dim fieldInfos As FieldInfo() = Me.GetType.GetFields(BindingFlags.[Public] Or
                                                                       BindingFlags.[Static] Or
                                                                       BindingFlags.FlattenHierarchy)

        Dim constants As IEnumerable(Of FieldInfo) =
                From f In fieldInfos Where f.IsLiteral AndAlso Not f.IsInitOnly Order By f.GetValue(Nothing)

        Dim requiredOnly = From c In constants Where c.IsDefined(GetType(RequiredFieldAttribute))

        Dim fields As IEnumerable(Of FieldInfo)
        If OnlyRequired Then
            fields = requiredOnly
        Else
            fields = constants
        End If

        Return String.Join(","c, fields.Select(Function(f) f.GetValue(Nothing)))
    End Function
End Class

Public NotInheritable Class NumberDataFields
    Inherits Fields
    Shared _ndf As NumberDataFields

    Private Sub New()
    End Sub

    Shared Function AllFields() As String
        If _ndf Is Nothing Then _ndf = New NumberDataFields
        Return _ndf.GetFields(False)
    End Function

    Shared Function RequiredFields() As String
        If _ndf Is Nothing Then _ndf = New NumberDataFields
        Return _ndf.GetFields(True)
    End Function

    <RequiredField>
    Public Const Number As String = "Number"
    <RequiredField>
    Public Const EvenOdd As String = "Even/Odd"

    Public Const SquareRoot As String = "Sqrt"

End Class
