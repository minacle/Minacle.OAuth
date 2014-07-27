Public Class OAuthCredentials

  Public Property Key As String

  Public Property Secret As String

  Public Sub New(key As String, secret As String)
    Me.Key = key
    Me.Secret = secret
  End Sub

  Public Overrides Function Equals(obj As Object) As Boolean
    If TypeOf obj Is OAuthCredentials Then
      Return Me.Key = DirectCast(obj, OAuthCredentials).Key AndAlso Me._Secret = DirectCast(obj, OAuthCredentials)._Secret
    Else
      Return MyBase.Equals(obj)
    End If
  End Function
End Class
