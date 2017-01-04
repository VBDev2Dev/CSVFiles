Imports System.Threading.Tasks
Namespace Importing
    Public MustInherit Class Importer(Of T)
        Implements IDisposable

        Public Sub New()
            MyClass.New(New Threading.CancellationTokenSource)
        End Sub
        Public Sub New(CancelSource As Threading.CancellationTokenSource)
            Me.CancelSource = CancelSource
        End Sub
        Protected _cancelSource As Threading.CancellationTokenSource
        Public Property CancelSource As Threading.CancellationTokenSource
            Get
                Return _cancelSource
            End Get
            Set(value As Threading.CancellationTokenSource)
                If value Is Nothing Then value = New Threading.CancellationTokenSource
                _cancelSource = value
            End Set
        End Property
        Private _Token As Threading.CancellationToken
        Public ReadOnly Property CancelToken As Threading.CancellationToken
            Get
                Return _Token
            End Get
        End Property
        Public Sub Cancel()
            If _cancelSource IsNot Nothing Then
                If Not _cancelSource.IsCancellationRequested Then _cancelSource.Cancel(True)
                _cancelSource.Dispose()
                _cancelSource = New Threading.CancellationTokenSource
            End If

        End Sub
        Protected _Result As IEnumerable(Of T)
        Public ReadOnly Property Result As IEnumerable(Of T)
            Get
                Return _Result
            End Get
        End Property
        Public MustOverride Function Import() As Task(Of IEnumerable(Of T))
        Event ImportProgress(sender As Object, e As ImportProgressEventArgs)
        Protected Overridable Sub OnImportProgress(e As ImportProgressEventArgs)
            RaiseEvent ImportProgress(Me, e)
        End Sub
        Protected Function PadID(ByVal id As String) As String
            Do Until id.Length >= 7
                id = "0" & id
            Loop
            Return id
        End Function
        ''' <summary>
        ''' Call this before you do anything else in the Import function
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overridable Sub SetupImport()
            If _cancelSource Is Nothing Then _cancelSource = New Threading.CancellationTokenSource
            _Token = _cancelSource.Token
        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    If CancelSource IsNot Nothing Then
                        CancelSource.Dispose()
                        _cancelSource = Nothing
                    End If
                    _Result = Nothing

                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

    Public Class ImportProgressEventArgs
        Inherits EventArgs
        Private _percentComplete As Double
        ReadOnly Property PercentComplete As Double
            Get
                Return _percentComplete
            End Get
        End Property
        Private _action As String
        ReadOnly Property Action As String
            Get
                Return _action
            End Get
        End Property
        Public Sub New(Action As String, PercentComplete As Double)
            _action = Action
            _percentComplete = PercentComplete
        End Sub

    End Class
End Namespace