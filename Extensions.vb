Imports System.Runtime.CompilerServices
Imports System.ComponentModel

<Extension>
<EditorBrowsable(EditorBrowsableState.Never)>
Public Module Extensions

  Private ReadOnly unreservedChars As Byte() = Encoding.ASCII.GetBytes("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.~")

  <Extension>
  <EditorBrowsable(EditorBrowsableState.Advanced)>
  Public Function GetResponseString(webResponse As WebResponse) As String
    Dim stream As Stream = webResponse.GetResponseStream
    Dim data As New MemoryStream
    Do
      Dim buf(4095) As Byte
      Dim len = stream.Read(buf, 0, buf.Length)
      If len = 0 Then Exit Do
      data.Write(buf, 0, len)
    Loop
    data.Position = 0
    Using reader As New StreamReader(data)
      GetResponseString = reader.ReadToEnd
    End Using
    data.Dispose()
    stream.Dispose()
  End Function

  <Extension>
  <EditorBrowsable(EditorBrowsableState.Advanced)>
  Public Function ToPercentEncoded(stringToEncode As String) As String
    Dim result As New StringBuilder
    For Each b In Encoding.UTF8.GetBytes(stringToEncode)
      If unreservedChars.Contains(b) Then
        result.Append(Chr(b))
      Else
        result.Append(String.Format("%{0:X2}", b))
      End If
    Next
    Return result.ToString
  End Function

  <Extension>
  <EditorBrowsable(EditorBrowsableState.Advanced)>
  Public Function ToQueryString(queryCollection As Specialized.NameValueCollection) As String
    Dim sb As New StringBuilder
    For Each k In queryCollection.AllKeys
      sb.Append(String.Format("{0}={1}&", k.ToPercentEncoded, queryCollection(k).ToPercentEncoded))
    Next
    If sb.Length > 0 Then sb.Remove(sb.Length - 1, 1)
    Return sb.ToString
  End Function

  <Extension>
  <EditorBrowsable(EditorBrowsableState.Advanced)>
  Public Function ToUnixTimestamp([date] As Date) As Integer
    Return Convert.ToInt32(([date] - #1/1/1970#).TotalSeconds)
  End Function

  <Extension>
  <EditorBrowsable(EditorBrowsableState.Advanced)>
  Public Sub UpdateQuery(ByRef uri As Uri, queryCollection As Specialized.NameValueCollection)
    Dim uriString = uri.GetLeftPart(UriPartial.Path)
    Dim q As Specialized.NameValueCollection
    If String.IsNullOrEmpty(uri.Query) Then
      q = Web.HttpUtility.ParseQueryString(String.Empty)
    Else
      q = Web.HttpUtility.ParseQueryString(uri.Query.Substring(1))
    End If
    For Each k In queryCollection.AllKeys
      q.Add(k, queryCollection(k))
    Next
    If q.Count > 0 Then
      uriString &= "?" & q.ToString
    End If
    uri = New Uri(uriString)
  End Sub
End Module
