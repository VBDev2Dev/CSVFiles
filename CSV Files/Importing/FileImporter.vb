Imports System.IO
Namespace Importing
    Public MustInherit Class FileImporter(Of T)
        Inherits Importer(Of T)
        Private _filePath As String = ""
        Property FilePath As String
            Get
                Return _filePath
            End Get
            Set(value As String)
                _filePath = value
            End Set
        End Property
        Public Sub New(FilePath As String)
            MyBase.New()
            Me.FilePath = FilePath
        End Sub
        Public Sub New(FilePath As String, CancelSource As Threading.CancellationTokenSource)
            MyBase.New(CancelSource)
            Me.FilePath = FilePath
        End Sub
        Protected Overrides Sub SetupImport()
            MyBase.SetupImport()
            If Not File.Exists(_filePath) Then Throw New FileNotFoundException("Cannot import file.", _filePath)
        End Sub
    End Class

End Namespace
