Imports System.Net

Public Class Leankit

    Public Class Account
        Public Shared Name As String
        Public Shared FullName As String
        Public Shared Email As String
        Public Shared Credentials As NetworkCredential

        Public Shared Sub SetUser(ByVal ServerResponse As String)
            Try
                Dim LastIndex As Integer = ServerResponse.IndexOf(""",""UserName"":""" & Email)
                Dim Substring As String = ServerResponse.Substring(0, LastIndex)
                Dim StartIndex As Integer = Substring.LastIndexOf("""FullName"":""") + """FullName"":""".Length
                FullName = ServerResponse.Substring(StartIndex, LastIndex - StartIndex)
            Catch ex As Exception

            End Try
        End Sub
    End Class

    Public Class Board
        Public Shared Title As String
        Public Shared Id As String
        Public Shared Description As String

        Public Shared Function Retrieve()
            Try
                Dim Leankit As New Web
                Return (Leankit.Request("https://" & Account.Name & ".leankitkanban.com/Kanban/Api/Boards/" & Id, "Get", String.Empty, Account.Credentials))
            Catch ex As Exception
                MsgBox("There was an error while attempting to load the board. This user may not have permsision to use this software.")
                Return Nothing
            End Try
        End Function
    End Class

    Public Class Card
        Public SystemType As String
        Public BoardId As String
        Public BoardTitle As String
        Public LaneId As String
        Public LaneTitle As String
        Public Title As String
        Public Description As String
        Public TypeId As String
        Public Priority As String
        Public PriorityText As String
        Public TypeName As String
        Public TypeColorHex As Color
        Public Size As Integer
        Public Color As Color
        Public AssignedUsers As String
        Public IsBlocked As Boolean
        Public Index As String
        Public StartDate As Date
        Public DueDate As Date
        Public ExternalSystemName As String
        Public ExternalSystemUrl As String
        Public ExternalCardID As String
        Public ExternalCardIdPrefix As String
        Public Tags As String
        Public ParentBoardId As String
        Public LastMove As Date
        Public LastActivity As Date
        Public Id As String
        Public ClassOfServiceId As String
        Public ClassOfServiceTitle As String
        Public LotNumber As String
    End Class

    Public Class Lane
        Public Panel As Panel
        Public Id As String
        Public Description As String
        Public Index As String
        Public Title As String
        Public Type As String
        Public Width As Integer
        Public ParentLaneId As String
        Public Orientation As String
        Public ChildLaneIds As String
        Public SiblingLaneIds As String
    End Class

    Public Class Priority
        Public Id As String
        Public Name As String
    End Class

    Public Class Type
        Public Id As String
        Public Name As String
    End Class

    Public Class ClassofService
        Public Id As String
        Public Name As String
    End Class

    Public Sub MoveCard(ByVal Card As Card, ByVal Lane As Lane, ByVal Position As String)
        Dim Leankit As New Web
        Leankit.Request("https://" & Account.Name & ".leankit.com/kanban/api/board/" & Board.Id & "/MoveCard/" & Card.Id & "/lane/" & Lane.Id & "/position/" & Position, "POST", String.Empty, Account.Credentials)
    End Sub

    Public Sub DeleteCard(ByVal Card As Card)
        Dim Leankit As New Web
        Leankit.Request("https://" & Account.Name & ".leankit.com/kanban/api/board/" & Board.Id & "/deletecard/" & Card.Id, "POST", String.Empty, Account.Credentials)
    End Sub

    Public Sub AddCard(ByVal Card As Card, ByVal Lane As Lane)
        Dim Answer As MsgBoxResult = MsgBox("Submition Confirmation: Are you sure you would like to continue?", MsgBoxStyle.YesNo)
        If Answer = vbYes Then
            Dim postData As String
            Try
                postData = "&Title=" & Card.Title
                postData += "&Description=" & Card.Description
                postData += "&TypeId=" & Card.TypeId
                postData += "&Priority=" & Card.Priority
                postData += "&Size=" & Card.Size
                postData += "&DueDate=" & Card.DueDate
                postData += "&ExternalSystemName=" & Card.ExternalSystemName
                postData += "&ExternalSystemUrl=" & Card.ExternalSystemUrl
                postData += "&Tags=" & Card.Tags
                postData += "&ClassOfServiceId=" & Card.ClassOfServiceId
                postData += "&ExternalCardID=" & Card.ExternalCardID

                Dim Leankit As New Web
                MsgBox(Leankit.Request("https://" & Account.Name & ".leankit.com/kanban/api/board/186012224/AddCard/lane/" & Lane.Id & "/position/1", "POST", postData, Account.Credentials))
            Catch ex As Exception
                MsgBox("There was an error processing your request. Please try again later.")
            End Try

        End If
    End Sub

    Public Sub UpdateCard(ByVal Card As Card)
        Dim Answer As MsgBoxResult = MsgBox("Submition Confirmation: Are you sure you would like to continue?", MsgBoxStyle.YesNo)
        If Answer = vbYes Then
            Dim postData As String
            Try
                postData = "&CardId=" & Card.Id
                postData += "&Title=" & Card.Title
                postData += "&Description=" & Card.Description
                postData += "&TypeId=" & Card.TypeId
                postData += "&Priority=" & Card.Priority
                postData += "&Size=" & Card.Size
                postData += "&DueDate=" & Card.DueDate
                postData += "&ExternalSystemName=" & Card.ExternalSystemName
                postData += "&ExternalSystemUrl=" & Card.ExternalSystemUrl
                postData += "&Tags=" & Card.Tags
                postData += "&ClassOfServiceId=" & Card.ClassOfServiceId
                postData += "&ExternalCardID=" & Card.ExternalCardID

                Dim Leankit As New Web
                MsgBox(Leankit.Request("https://" & Account.Name & ".leankit.com/kanban/api/card/update", "POST", postData, Account.Credentials))
            Catch ex As Exception
                MsgBox("There was an error processing your request. Please try again later.")
            End Try

        End If
    End Sub
End Class