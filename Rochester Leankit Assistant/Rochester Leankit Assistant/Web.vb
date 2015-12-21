Imports System.Net
Imports System.Text
Imports System.IO

Public Class Web
    Public Function Request(ByVal Address As String, ByVal Method As String, Optional ByVal body As String = Nothing, Optional ByVal Credentials As NetworkCredential = Nothing)
        Try
            Dim wrequest As WebRequest = WebRequest.Create(Address)
            wrequest.Method = Method
            wrequest.Credentials = Credentials
            wrequest.PreAuthenticate = True
            If Method = "POST" Then
                If Not String.IsNullOrEmpty(body) Then
                    Dim byteData As Byte() = Encoding.UTF8.GetBytes(body)
                    wrequest.ContentLength = byteData.Length
                    wrequest.ContentType = "application/x-www-form-urlencoded"
                    'request.ContentType = "application/json"
                    Dim dataStream As Stream = wrequest.GetRequestStream()
                    dataStream.Write(byteData, 0, byteData.Length)
                    dataStream.Close()
                Else
                    wrequest.ContentLength = 0
                End If
            End If
            wrequest.Timeout = 15000
            wrequest.CachePolicy = New Cache.RequestCachePolicy(Cache.RequestCacheLevel.BypassCache)

            Dim response As WebResponse = wrequest.GetResponse()
            Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
            Dim reader As New StreamReader(response.GetResponseStream(), Encoding.GetEncoding(1252))
            Dim responseFromServer As String = reader.ReadToEnd()
            reader.Close()
            response.Close()
            Return responseFromServer
        Catch ex As WebException
            Return ex.Message
            MsgBox("There was a problem trying to access the Webserver. Please try again later. Error Message:  " & ex.Message)
        End Try
    End Function
End Class
