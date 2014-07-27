'RFC 5849 
Namespace OAuthClient

  Public Class OAuth1

    Private Const OAUTH_VERSION = "1.0"

    Public Property ConsumerCredentials As OAuthCredentials
    Public Property RequestCredentials As OAuthCredentials
    Public Property AccessCredentials As OAuthCredentials
    Public Property DefaultSignatureMethod As HMAC

    Public Function Initiate(method As String, uri As Uri, Optional realm As String = Nothing, Optional oauth_consumer_key As String = Nothing, Optional oauth_signature_method As String = Nothing, Optional oauth_timestamp As String = Nothing, Optional oauth_nonce As String = Nothing, Optional oauth_callback As String = Nothing) As Boolean
      Dim req = MakeRequest(uri)
      req.Method = method
      req.Headers.Add("Authorization", MakeOAuthAuthorisationString(method, uri, realm:=realm, oauth_consumer_key:=oauth_consumer_key, oauth_signature_method:=oauth_signature_method, oauth_timestamp:=oauth_timestamp, oauth_nonce:=oauth_nonce, oauth_callback:=oauth_callback))
      Dim res As HttpWebResponse = Nothing
      Try
        res = req.GetResponse
        Dim q = Web.HttpUtility.ParseQueryString(GetResponseString(res))
        RequestCredentials = New OAuthCredentials(q("oauth_token"), q("oauth_token_secret"))
      Catch
      End Try
      If res IsNot Nothing Then res.Dispose()
      Return RequestCredentials IsNot Nothing
    End Function

    Public Overridable Sub Authorise(uri As Uri)
      Throw New NotImplementedException
    End Sub

    Public Function Token(method As String, uri As Uri, Optional realm As String = Nothing, Optional oauth_consumer_key As String = Nothing, Optional oauth_token As String = Nothing, Optional oauth_signature_method As String = Nothing, Optional oauth_timestamp As String = Nothing, Optional oauth_nonce As String = Nothing, Optional oauth_callback As String = Nothing, Optional oauth_verifier As String = Nothing) As Boolean
      Dim req = MakeRequest(uri)
      req.Method = method
      req.Headers.Add("Authorization", MakeOAuthAuthorisationString(method, uri, realm:=realm, oauth_consumer_key:=oauth_consumer_key, oauth_token:=RequestCredentials.Key, oauth_signature_method:=oauth_signature_method, oauth_timestamp:=oauth_timestamp, oauth_nonce:=oauth_nonce, oauth_callback:=oauth_callback, oauth_verifier:=oauth_verifier))
      Dim res As HttpWebResponse = Nothing
      Try
        res = req.GetResponse
        Dim q = Web.HttpUtility.ParseQueryString(GetResponseString(res))
        AccessCredentials = New OAuthCredentials(q("oauth_token"), q("oauth_token_secret"))
      Catch
      End Try
      If res IsNot Nothing Then res.Dispose()
      Return AccessCredentials IsNot Nothing
    End Function

    Public Function MakeOAuthAuthorisationString(method As String, uri As Uri, Optional realm As String = Nothing, Optional oauth_consumer_key As String = Nothing, Optional oauth_token As String = Nothing, Optional oauth_signature_method As String = Nothing, Optional oauth_timestamp As String = Nothing, Optional oauth_nonce As String = Nothing, Optional oauth_callback As String = Nothing, Optional oauth_verifier As String = Nothing, Optional oauth_signature As String = Nothing, Optional content As Specialized.NameValueCollection = Nothing) As String
      Dim signatureMethod = Me.DefaultSignatureMethod
      If oauth_consumer_key Is Nothing Then oauth_consumer_key = ConsumerCredentials.Key
      If oauth_token Is Nothing AndAlso AccessCredentials IsNot Nothing Then oauth_token = AccessCredentials.Key
      If oauth_timestamp Is Nothing Then oauth_timestamp = Date.UtcNow.ToUnixTimestamp
      If oauth_signature Is Nothing Then oauth_signature = MakeOAuthSignature(method, uri, realm, oauth_consumer_key, oauth_token, oauth_signature_method, oauth_timestamp, oauth_nonce, oauth_callback, oauth_verifier, content)
      Dim result As New StringBuilder
      With result
        .Append("OAuth")
        If realm IsNot Nothing Then .AppendFormat(" realm=""{0}"",", realm)
        If oauth_consumer_key IsNot Nothing Then .AppendFormat(" oauth_consumer_key=""{0}"",", oauth_consumer_key)
        If oauth_token IsNot Nothing Then .AppendFormat(" oauth_token=""{0}"",", oauth_token)
        If oauth_timestamp IsNot Nothing Then .AppendFormat(" oauth_timestamp=""{0}"",", oauth_timestamp)
        If oauth_nonce IsNot Nothing Then .AppendFormat(" oauth_nonce=""{0}"",", oauth_nonce)
        If oauth_callback IsNot Nothing Then .AppendFormat(" oauth_callback=""{0}"",", oauth_callback)
        If oauth_verifier IsNot Nothing Then .AppendFormat(" oauth_verifier=""{0}"",", oauth_verifier)
        If signatureMethod Is Nothing Then
          .AppendFormat(" oauth_signature_method=""{0}"",", "PLAINTEXT")
        Else
          .AppendFormat(" oauth_signature_method=""{0}"",", "HMAC-" & signatureMethod.HashName)
        End If
        If oauth_signature IsNot Nothing Then .AppendFormat(" oauth_signature=""{0}"",", oauth_signature)
        .AppendFormat(" oauth_version=""{0}"",", OAUTH_VERSION)
        .Remove(.Length - 1, 1)
        Return .ToString
      End With
    End Function

    Public Function MakeOAuthQueryString(method As String, uri As Uri, Optional realm As String = Nothing, Optional oauth_consumer_key As String = Nothing, Optional oauth_token As String = Nothing, Optional oauth_signature_method As String = Nothing, Optional oauth_timestamp As String = Nothing, Optional oauth_nonce As String = Nothing, Optional oauth_callback As String = Nothing, Optional oauth_verifier As String = Nothing, Optional oauth_signature As String = Nothing, Optional content As Specialized.NameValueCollection = Nothing) As String
      Dim signatureMethod = Me.DefaultSignatureMethod
      If oauth_consumer_key Is Nothing Then oauth_consumer_key = ConsumerCredentials.Key
      If oauth_token Is Nothing AndAlso AccessCredentials IsNot Nothing Then oauth_token = AccessCredentials.Key
      If oauth_timestamp Is Nothing Then oauth_timestamp = Date.UtcNow.ToUnixTimestamp
      If oauth_signature Is Nothing Then oauth_signature = MakeOAuthSignature("GET", uri, realm, oauth_consumer_key, oauth_token, oauth_signature_method, oauth_timestamp, oauth_nonce, oauth_callback, oauth_verifier, content)
      Dim q As New Specialized.NameValueCollection
      If realm IsNot Nothing Then q.Add("realm", realm)
      If oauth_consumer_key IsNot Nothing Then q.Add("oauth_consumer_key", oauth_consumer_key)
      If oauth_token IsNot Nothing Then q.Add("oauth_token", oauth_token)
      If oauth_timestamp IsNot Nothing Then q.Add("oauth_timestamp", oauth_timestamp)
      If oauth_nonce IsNot Nothing Then q.Add("oauth_nonce=", oauth_nonce)
      If oauth_callback IsNot Nothing Then q.Add("oauth_callback=", oauth_callback)
      If oauth_verifier IsNot Nothing Then q.Add("oauth_verifier=", oauth_verifier)
      If signatureMethod Is Nothing Then
        q.Add("oauth_signature_method=", "PLAINTEXT")
      Else
        q.Add("oauth_signature_method=", "HMAC-" & signatureMethod.HashName)
      End If
      If oauth_signature IsNot Nothing Then q.Add("oauth_signature=", oauth_signature)
      q.Add("oauth_version=", OAUTH_VERSION)
      Return q.ToQueryString
    End Function

    Protected Function MakeOAuthSignature(method As String, uri As Uri, Optional realm As String = Nothing, Optional oauth_consumer_key As String = Nothing, Optional oauth_token As String = Nothing, Optional oauth_signature_method As String = Nothing, Optional oauth_timestamp As String = Nothing, Optional oauth_nonce As String = Nothing, Optional oauth_callback As String = Nothing, Optional oauth_verifier As String = Nothing, Optional content As Specialized.NameValueCollection = Nothing) As String
      Dim signatureMethod = Me.DefaultSignatureMethod
      If oauth_consumer_key Is Nothing Then oauth_consumer_key = ConsumerCredentials.Key
      If oauth_timestamp Is Nothing Then oauth_timestamp = Date.UtcNow.ToUnixTimestamp
      Dim signatureBase As New StringBuilder
      If oauth_signature_method IsNot Nothing Then signatureMethod = HMAC.Create(oauth_signature_method.Replace("-", ""))
      With signatureBase
        Dim parameters As New SortedDictionary(Of String, String)
        parameters("oauth_version") = OAUTH_VERSION
        .Append(method.ToUpperInvariant.ToPercentEncoded & "&")
        .Append(String.Format("{0}{1}{2}{3}", uri.Scheme, uri.SchemeDelimiter, uri.Host, uri.AbsolutePath).ToPercentEncoded & "&")
        If Not String.IsNullOrEmpty(uri.Query) Then
          Dim queryString = Web.HttpUtility.ParseQueryString(uri.Query.Substring(1))
          For Each k In queryString.AllKeys
            parameters(k) = queryString(k)
          Next
        End If
        If content IsNot Nothing Then
          For Each k In content.AllKeys
            parameters(k) = content(k)
          Next
        End If
        If realm IsNot Nothing Then parameters("realm") = realm
        If oauth_consumer_key IsNot Nothing Then parameters("oauth_consumer_key") = oauth_consumer_key
        If oauth_token IsNot Nothing Then parameters("oauth_token") = oauth_token
        If oauth_timestamp IsNot Nothing Then parameters("oauth_timestamp") = oauth_timestamp
        If oauth_nonce IsNot Nothing Then parameters("oauth_nonce") = oauth_nonce
        If oauth_callback IsNot Nothing Then parameters("oauth_callback") = oauth_callback
        If oauth_verifier IsNot Nothing Then parameters("oauth_verifier") = oauth_verifier
        If signatureMethod Is Nothing Then
          parameters("oauth_signature_method") = "PLAINTEXT"
        Else
          parameters("oauth_signature_method") = "HMAC-" & signatureMethod.HashName.ToPercentEncoded
        End If
        For i = 0 To parameters.Count - 1
          Dim k = parameters.Keys(i)
          .Append(String.Format("{0}={1}", k.ToPercentEncoded, parameters(k).ToPercentEncoded).ToPercentEncoded)
          If i < parameters.Count - 1 Then .Append("&".ToPercentEncoded)
        Next
        If signatureMethod Is Nothing Then
          MakeOAuthSignature = (ConsumerCredentials.Secret & "&" & If(AccessCredentials IsNot Nothing, AccessCredentials.Secret, String.Empty)).ToPercentEncoded
        Else
          signatureMethod.Key = Encoding.ASCII.GetBytes(ConsumerCredentials.Secret & "&" & If(AccessCredentials IsNot Nothing, AccessCredentials.Secret, String.Empty))
          MakeOAuthSignature = Convert.ToBase64String(signatureMethod.ComputeHash(Encoding.ASCII.GetBytes(.ToString))).ToPercentEncoded
        End If
      End With
    End Function

    Protected Function MakeRequest(requestUriString As String) As HttpWebRequest
      Return MakeRequest(New Uri(requestUriString))
    End Function

    Protected Function MakeRequest(requestUri As Uri) As HttpWebRequest
      MakeRequest = WebRequest.Create(requestUri)
    End Function
  End Class

End Namespace