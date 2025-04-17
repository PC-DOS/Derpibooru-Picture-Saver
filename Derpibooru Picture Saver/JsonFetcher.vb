Imports System.Threading.Tasks

Public Class JsonFetcher
    Private _HttpReq As SimpleHttpRequest.HttpJsonRequest

    Private _JsonUrl As String
    Private _JsonText As String

    Public Sub New(Optional JsonTargetUrl As String = "")
        _HttpReq = New SimpleHttpRequest.HttpJsonRequest
        _JsonUrl = JsonTargetUrl
    End Sub

    Public Function DoGet(JsonTargetUrl As String) As String
        _JsonUrl = JsonTargetUrl
        _JsonText = _HttpReq.DoGet(JsonTargetUrl)
        Return _JsonText
    End Function

    Public Async Function DoGetAsync(JsonTargetUrl As String) As Task(Of String)
        Await Task.Factory.StartNew(Sub()
                                        Me.DoGet(JsonTargetUrl)
                                    End Sub)
        Return _JsonText
    End Function

    Public Property JsonUrl As String
        Get
            Return _JsonUrl
        End Get
        Set(value As String)
            _JsonUrl = value
        End Set
    End Property

    Public ReadOnly Property JsonText As String
        Get
            Return _JsonUrl
        End Get
    End Property
End Class

