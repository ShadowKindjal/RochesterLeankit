Imports System.Environment
Imports System.Xml

Public Class MediaFire

    'Public Name As String
    'Public Key As String
    Private ErrorMessage As String

    Function getToken()
        Try
            Dim XmlDoc As New XmlDocument
            Dim Web As New Web
            XmlDoc.LoadXml(Web.Request("https://www.mediafire.com/api/user/get_session_token.php?email=nck.harmon@gmail.com&password=j112499h&application_id=45463&signature=2bf74cd0297a1b6ba8345d87e344b3766bbf9b7b", "GET"))
            Return XmlDoc.SelectSingleNode("//response/session_token").InnerText
        Catch ex As Exception
            ErrorMessage = "There was a problem trying to access the Webserver. Please try again later. Error Message: " & ex.Message
            Return ErrorMessage
        End Try
    End Function

    Sub getFile(ByVal Key As String, ByVal Name As String, ByVal Update As Boolean, ByVal Extension As String)
        Try
            Dim XmlDoc As New XmlDocument
            Dim XMLDir As String = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\" & Name & Extension
            Dim Web As New Web
            If Update Then
                If My.Computer.FileSystem.FileExists(XMLDir) Then My.Computer.FileSystem.DeleteFile(XMLDir)
                XmlDoc.LoadXml(Web.Request("http://www.mediafire.com/api/file/get_links.php?link_type=direct_download&session_token=" & getToken() & "&quick_key=" & Key & "&response_format=xml", "GET"))
                Dim DownloadLink As String = XmlDoc.SelectSingleNode("//response/links/link/direct_download").InnerText
                My.Computer.Network.DownloadFile(DownloadLink, XMLDir)
            Else
                If Not My.Computer.FileSystem.FileExists(XMLDir) Then
                    XmlDoc.LoadXml(Web.Request("http://www.mediafire.com/api/file/get_links.php?link_type=direct_download&session_token=" & getToken() & "&quick_key=" & Key & "&response_format=xml", "GET"))
                    Dim DownloadLink As String = XmlDoc.SelectSingleNode("//response/links/link/direct_download").InnerText
                    My.Computer.Network.DownloadFile(DownloadLink, XMLDir)
                End If
            End If
        Catch ex As Exception
            If ErrorMessage = Nothing Then MsgBox("There was a problem trying to download program resources. Please try again later. Error Message: " & ex.Message)
            MsgBox(ErrorMessage)
            Application.Exit()
        End Try
    End Sub

End Class
