Imports System.Xml
Imports System.Environment
Imports Microsoft.Office.Interop

Public Class Manager
    Public Cards() As Leankit.Card = Nothing
    Public Lanes() As Leankit.Lane = Nothing
    Public Priorities() As Leankit.Priority = Nothing
    Public CardTypes() As Leankit.Type = Nothing
    Public Shared CurrentForm As Form

    Private Sub Manager_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Board As New Leankit.Board
        Dim ServerResponse As String = Board.Retrieve
        Leankit.Account.SetUser(ServerResponse)
        GUISetup(ServerResponse)
        Populate(ServerResponse)
        ToolbarPopulate()
    End Sub

    Sub callPopulate()
        Dim Board As New Leankit.Board
        Populate(Board.Retrieve)
    End Sub

    Sub Populate(Optional ByVal ServerResponse As String = "")
        Me.SuspendLayout()
        Dim Board As Panel = Me.Controls(Leankit.Board.Id)
        Board.Controls.Clear()

        CardPopulate(ServerResponse)
        LanePopulate(ServerResponse)
        PrioritiesPopulate()
        CardTypesPopulate(ServerResponse)
        Me.ResumeLayout()
    End Sub

    Sub GUISetup(ByVal ServerResponse As String)
        Me.Controls.Clear()
        Me.Text = "Rochester Leankit Assistant"
        Me.Size = New Size(1366, 768)
        Me.MinimumSize = New Size(720, 480)
        Me.Top = (Screen.PrimaryScreen.Bounds.Height - Me.Height) / 2
        Me.Left = (Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.AutoScroll = True
        Me.WindowState = FormWindowState.Maximized

        Dim Board As New Panel
        Board.Name = Leankit.Board.Id
        Board.Dock = DockStyle.Fill
        Board.BackColor = Color.FromArgb(197, 201, 190)
        Board.AutoScroll = True
        Me.Controls.Add(Board)


        Dim Toolbar As New Panel
        Toolbar.Name = "Toolbar"
        Toolbar.Dock = DockStyle.Right
        Toolbar.Width = 250
        Toolbar.AutoScroll = True
        Toolbar.BackColor = Color.White
        Me.Controls.Add(Toolbar)

        Dim TitleBar As New Panel
        TitleBar.Name = "TitleBar"
        TitleBar.Location = New Point(0, 0)
        TitleBar.Dock = DockStyle.Top
        TitleBar.Height = 50
        TitleBar.BackColor = Color.FromArgb(39, 43, 36)
        Me.Controls.Add(TitleBar)

        Dim Picture As New PictureBox
        Picture.Name = "Picture"
        Picture.Size = New Size(75, 25)
        Picture.Location = New Point(25, 11)
        Picture.Anchor = AnchorStyles.Left Or AnchorStyles.Top
        Picture.Cursor = Cursors.Hand
        Picture.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\Leankit-White.png"
        Picture.BorderStyle = BorderStyle.None
        Picture.SizeMode = PictureBoxSizeMode.StretchImage
        TitleBar.Controls.Add(Picture)
        Picture.Load()
        AddHandler Picture.Click, AddressOf CloseForm

        Dim Title As New Label
        Title.Name = "Title"
        Title.AutoSize = False
        Title.Location = New Point(125, 12.5)
        Title.Size = New Size(TitleBar.Width - Title.Location.X - 275, 25)
        Title.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        Title.Text = StrConv(ServerResponse.Substring(ServerResponse.IndexOf("""Title"":""", 0) + """Title"":""".Length, (ServerResponse.IndexOf(""",""Description"":", 0)) - (ServerResponse.IndexOf("""Title"":""", 0) + """Title"":""".Length)), VbStrConv.Uppercase)
        Title.Font = New Font("Helvetica", 12, FontStyle.Bold)
        Title.TextAlign = ContentAlignment.MiddleLeft
        Title.ForeColor = Color.White
        TitleBar.Controls.Add(Title)

        Dim Button As New PictureBox
        Button.Name = "Button"
        Button.Size = New Size(25, 25)
        Button.Location = New Point(TitleBar.Width - 50, 12.5)
        Button.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        Button.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\Triangle.png"
        Button.BorderStyle = BorderStyle.None
        Button.SizeMode = PictureBoxSizeMode.StretchImage
        Button.Cursor = Cursors.Hand
        TitleBar.Controls.Add(Button)
        Button.Load()
        AddHandler Button.Click, AddressOf HideShow

        Dim Logout As New PictureBox
        Logout.Name = "Logout"
        Logout.Size = New Size(25, 25)
        Logout.Location = New Point(TitleBar.Width - 100, 12.5)
        Logout.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        Logout.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\Logout.png"
        Logout.BorderStyle = BorderStyle.None
        Logout.SizeMode = PictureBoxSizeMode.StretchImage
        Logout.Cursor = Cursors.Hand
        TitleBar.Controls.Add(Logout)
        Logout.Load()
        AddHandler Logout.Click, AddressOf UserToolbar

        Dim Add As New PictureBox
        Add.Name = "Add"
        Add.Size = New Size(25, 25)
        Add.Location = New Point(TitleBar.Width - 150, 12.5)
        Add.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        Add.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\PlusSign.png"
        Add.BorderStyle = BorderStyle.None
        Add.SizeMode = PictureBoxSizeMode.StretchImage
        Add.Cursor = Cursors.Hand
        TitleBar.Controls.Add(Add)
        Add.Load()
        AddHandler Add.Click, AddressOf AddToolbar

        Dim Export As New PictureBox
        Export.Name = "Export"
        Export.Size = New Size(25, 25)
        Export.Location = New Point(TitleBar.Width - 200, 12.5)
        Export.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        Export.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\Export.png"
        Export.BorderStyle = BorderStyle.None
        Export.SizeMode = PictureBoxSizeMode.StretchImage
        Export.Cursor = Cursors.Hand
        TitleBar.Controls.Add(Export)
        Export.Load()
        AddHandler Export.Click, AddressOf ExportToolbar

        Dim Refresh As New PictureBox
        Refresh.Name = "Refresh"
        Refresh.Size = New Size(25, 25)
        Refresh.Location = New Point(TitleBar.Width - 250, 12.5)
        Refresh.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        Refresh.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\refresh.png"
        Refresh.BorderStyle = BorderStyle.None
        Refresh.SizeMode = PictureBoxSizeMode.StretchImage
        Refresh.Cursor = Cursors.Hand
        TitleBar.Controls.Add(Refresh)
        Refresh.Load()
        AddHandler Refresh.Click, AddressOf callPopulate
        Toolbar.Hide()
    End Sub

    Sub ToolbarPopulate()
        Dim Toolbar As Panel = Me.Controls("Toolbar")
        Toolbar.Controls.Clear()
        Toolbar.AutoScroll = True

        Dim Title As New Label
        Title.Name = "Title"
        Title.AutoSize = False
        Title.Location = New Point(10, 12.5)
        Title.Size = New Size(Toolbar.Width - 20, 25)
        Title.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        Title.Text = "Select an option from above"
        Title.TextAlign = ContentAlignment.MiddleCenter
        Title.ForeColor = Color.DarkGray
        Toolbar.Controls.Add(Title)
    End Sub

    Sub CardPopulate(ByRef ServerResponse As String)
        Dim Count As Integer = 0
        Dim position As Integer = 0
        While position <> -1
            Try
                position = ServerResponse.IndexOf("""Card"",""LaneId"":", position) + """Card"",""LaneId"":".Length
                If position <> -1 Then
                    Dim endPosition As Integer = ServerResponse.IndexOf(",""Title"":", position + 1)
                    If endPosition <> -1 Then
                        ReDim Preserve Cards(Count)
                        Dim Card As New Leankit.Card
                        Card.LaneId = ServerResponse.Substring(position, endPosition - position)
                        Card.Id = ServerResponse.Substring((ServerResponse.IndexOf(",""DrillThroughBoardId", position)) - 9, 9)
                        Card.Color = ColorTranslator.FromHtml(ServerResponse.Substring(ServerResponse.IndexOf("""Color"":""", position) + """Color"":""".Length, (ServerResponse.IndexOf(""",""Version"":", position)) - (ServerResponse.IndexOf("""Color"":""", position) + """Color"":""".Length)))
                        Card.Title = ServerResponse.Substring(ServerResponse.IndexOf("""Title"":""", position) + """Title"":""".Length, (ServerResponse.IndexOf(""",""Description"":", position)) - (ServerResponse.IndexOf("""Title"":""", position) + """Title"":""".Length))
                        Card.ExternalCardID = ServerResponse.Substring(ServerResponse.IndexOf("""ExternalCardID"":""", position) + """ExternalCardID"":""".Length, (ServerResponse.IndexOf(""",""Tags"":", position)) - (ServerResponse.IndexOf("""ExternalCardID"":""", position) + """ExternalCardID"":""".Length))
                        Card.PriorityText = ServerResponse.Substring(ServerResponse.IndexOf("""PriorityText"":""", position) + """PriorityText"":""".Length, (ServerResponse.IndexOf(""",""TypeName"":", position)) - (ServerResponse.IndexOf("""PriorityText"":""", position) + """PriorityText"":""".Length))
                        Card.TypeName = ServerResponse.Substring(ServerResponse.IndexOf("""TypeName"":""", position) + """TypeName"":""".Length, (ServerResponse.IndexOf(""",""TypeIconPath"":", position)) - (ServerResponse.IndexOf("""TypeName"":""", position) + """TypeName"":""".Length))
                        Card.Size = CInt(ServerResponse.Substring(ServerResponse.IndexOf(""",""Size"":", position) + """,""Size"":".Length, (ServerResponse.IndexOf(",""Active""", position)) - (ServerResponse.IndexOf(""",""Size"":", position) + """,""Size"":".Length)))
                        Card.LotNumber = Card.ExternalCardID.Substring(Card.ExternalCardID.Length - 6, 6)
                        Cards(Count) = Card
                        Count += 1
                    End If
                    position = endPosition
                End If
            Catch ex As Exception
                ReDim Preserve Cards(Count - 1)
                Exit While
            End Try
        End While
    End Sub

    Sub LanePopulate(ByRef ServerResponse As String)
        Dim Count As Integer = 0
        Dim LaneCheck As Boolean
        Dim position As Integer = 0
        Dim lastposition As Integer = 0
        While position <> -1
            position = ServerResponse.IndexOf("""Id"":", position) + """Id"":".Length
            If position < lastposition Then Exit While
            If ServerResponse.Substring(position + 11, 11) = "Description" Then
                LaneCheck = True
            Else
                LaneCheck = False
            End If
            If position <> -1 Then
                Dim endposition As Integer = ServerResponse.IndexOf(",""", position + 1)
                If endposition <> -1 And LaneCheck Then
                    ReDim Preserve Lanes(Count)
                    Dim Lane As New Leankit.Lane
                    Lane.ParentLaneId = ServerResponse.Substring(ServerResponse.IndexOf(",""ParentLaneId"":", position) + ",""ParentLaneId"":".Length, (ServerResponse.IndexOf(",""Cards"":", position)) - (ServerResponse.IndexOf(",""ParentLaneId"":", position) + ",""ParentLaneId"":".Length))
                    Lane.ChildLaneIds = ServerResponse.Substring(ServerResponse.IndexOf(",""ChildLaneIds"":", position) + ",""ChildLaneIds"":".Length, (ServerResponse.IndexOf(",""SiblingLaneIds"":", position)) - (ServerResponse.IndexOf(",""ChildLaneIds"":", position) + ",""ChildLaneIds"":".Length))
                    Lane.Title = ServerResponse.Substring(ServerResponse.IndexOf("""Title"":""", position) + """Title"":""".Length, (ServerResponse.IndexOf(""",""CardLimit"":", position)) - (ServerResponse.IndexOf("""Title"":""", position) + """Title"":""".Length))
                    Lane.Index = ServerResponse.Substring(ServerResponse.IndexOf(",""Index"":", position) + ",""Index"":".Length, (ServerResponse.IndexOf(",""Active"":", position)) - (ServerResponse.IndexOf(",""Index"":", position) + ",""Index"":".Length))
                    Lane.Width = ServerResponse.Substring(ServerResponse.IndexOf(",""Width"":", position) + ",""Width"":".Length, (ServerResponse.IndexOf(",""ParentLaneId"":", position)) - (ServerResponse.IndexOf(",""Width"":", position) + ",""Width"":".Length))
                    Lane.Id = ServerResponse.Substring(position, endposition - position)
                    Lanes(Count) = Lane
                    Count += 1
                End If
                position = endposition
                lastposition = position
            End If
        End While
        Count = 0
        BoardPopulate(0, New Point(0, 0))
    End Sub

    Sub BoardPopulate(ByVal Parent As String, ByRef start As Point, Optional ByVal Width As Integer = 0)
        Dim Board As Panel = Me.Controls(Leankit.Board.Id)
        Dim bool As Boolean = True

        Dim Count As Integer = 0
        Dim cCount As Integer = 0
        Dim Counter As Integer = 0
        Dim Index As Integer = 0
        Dim cIndex As Integer = 0
        Dim StartX As Integer = start.X
        Dim StartY As Integer = start.Y
        Dim PanelH As Integer = 0
        Dim PanelL As Integer = 0

        While Count < 200 'And StartY < 40
            'MsgBox(Count)
            For Index = 0 To Lanes.Count - 1
                With Lanes(Index)
                    If .Index = Count And .ParentLaneId = Parent Then
                        Dim Title As New Label
                        Title.Text = .Title
                        Title.Name = .Id
                        Title.Location = New Point(StartX + 10, StartY + 10)
                        Title.Size = New Size(200 + ((CType(.Width, Integer) - 1) * 210), 15)
                        Title.Font = New Font("Helvetica", 8, FontStyle.Bold)
                        Title.TextAlign = ContentAlignment.MiddleCenter
                        Title.BackColor = Color.White
                        Board.Controls.Add(Title)


                        If .Title.ToUpper = "NOT STARTED - FUTURE WORK" Then
                            Title.Width = 210 + ((CType(.Width, Integer) - 1) * 200)
                        Else
                            Title.Width = 200 + ((CType(.Width, Integer) - 1) * 190)
                        End If

                        If .ChildLaneIds <> "[]" Then
                            BoardPopulate(.Id, New Point(Title.Location.X - 10, Title.Location.Y + Title.Height), .Width)
                        Else

                            Dim Panel As New Panel
                            Panel.Name = .Id
                            Panel.Location = New Point(Title.Location.X, Title.Location.Y + 25)
                            Panel.Anchor = AnchorStyles.Left Or AnchorStyles.Top
                            Panel.Size = New Size(Title.Width, 120)
                            'Panel.MinimumSize = New Size(200, Me.ClientSize.Height - Me.Location.Y)
                            Panel.BackColor = Color.FromArgb(221, 223, 216)
                            Panel.AutoScroll = False
                            Panel.AllowDrop = True
                            Board.Controls.Add(Panel)
                            'AddHandler Panel.DragEnter, AddressOf cDragEnter
                            'AddHandler Panel.DragDrop, AddressOf cDragDrop

                            PanelL = Panel.Location.Y
                            PanelH = Panel.Height

                            Dim cStartX As Integer = 0
                            Dim cStarty As Integer = 0
                            Dim cardbool As Boolean = True


                            For cIndex = 0 To Cards.Count - 1
                                With Cards(cIndex)
                                    If .LaneId = Lanes(Index).Id Then

                                        Dim Card As New Panel
                                        Card.Size = New Size(180, 100)
                                        Card.Name = .Id
                                        'MsgBox(Card.Name)
                                        Card.Location = New Point(cStartX + 10, cStarty + 10)
                                        Card.BackColor = .Color
                                        Card.Cursor = Cursors.Hand
                                        'AddHandler Card.Click, AddressOf LoadCard
                                        'AddHandler Card.MouseDown, AddressOf cMouseDown
                                        'AddHandler Card.MouseUp, AddressOf cMouseUp
                                        'AddHandler Card.MouseMove, AddressOf cMouseMove

                                        Dim CardTitle As New Label
                                        CardTitle.Text = .ExternalCardID
                                        'MsgBox(CardTitle.Text)
                                        CardTitle.Dock = DockStyle.Top
                                        CardTitle.Height = 15
                                        CardTitle.Font = New Font("Helvetica", 8, FontStyle.Bold)
                                        CardTitle.TextAlign = ContentAlignment.MiddleCenter
                                        CardTitle.BackColor = ControlPaint.Dark(Card.BackColor)
                                        CardTitle.ForeColor = Color.White
                                        Card.Controls.Add(CardTitle)
                                        'AddHandler CardTitle.Click, AddressOf LoadCardSend
                                        'AddHandler CardTitle.MouseDown, AddressOf cMouseDown
                                        'AddHandler CardTitle.MouseUp, AddressOf cMouseUp
                                        'AddHandler CardTitle.MouseMove, AddressOf pMouseMove

                                        Dim CardMessage As New Label
                                        CardMessage.AutoSize = False
                                        CardMessage.Text = .Title
                                        'MsgBox(CardMessage.Text)
                                        CardMessage.Size = New Size(Card.Width - 10, Card.Height - 25)
                                        CardMessage.Location = New Point(5, 20)
                                        Card.Anchor = AnchorStyles.Top Or AnchorStyles.Left
                                        CardMessage.Font = New Font("Helvetica", 8, FontStyle.Regular)
                                        CardMessage.TextAlign = ContentAlignment.TopLeft
                                        CardMessage.BackColor = Color.Transparent
                                        Card.Controls.Add(CardMessage)
                                        'AddHandler CardMessage.Click, AddressOf LoadCardSend
                                        'AddHandler CardMessage.MouseDown, AddressOf cMouseDown
                                        'AddHandler CardMessage.MouseUp, AddressOf cMouseUp
                                        'AddHandler CardMessage.MouseMove, AddressOf pMouseMove

                                        Panel.Controls.Add(Card)
                                        Panel.Height = Card.Location.Y + Card.Height + 10
                                        PanelL = Panel.Location.Y
                                        PanelH = Panel.Height

                                        If Lanes(Index).Width > 1 Then
                                            If cardbool Then
                                                cStartX = Card.Location.X + Card.Width
                                                cardbool = False
                                            Else
                                                cStartX = 0
                                                cStarty = Card.Location.Y + Card.Height
                                                cardbool = True
                                            End If
                                        Else
                                            cStarty = Card.Location.Y + Card.Height
                                            cStartX = 0
                                        End If
                                    End If
                                End With

                            Next

                        End If

                        If .Width = Width Then
                            StartY = PanelL + PanelH
                        Else
                            StartX = Title.Location.X + Title.Width
                        End If
                        Count += 1
                        bool = False
                    End If
                End With
            Next
            If bool Then Count += 1
            bool = True
        End While
    End Sub

    Sub PrioritiesPopulate()
        Dim Web As New Web
        Dim ServerResponse As String = Web.Request("https://" & Leankit.Account.Name & ".leankitkanban.com/Kanban/Api/Board/" & Leankit.Board.Id & "/GetBoardIdentifiers", "GET", String.Empty, Leankit.Account.Credentials)
        Dim Count As Integer = 0
        Dim position As Integer = ServerResponse.IndexOf("""Priorities"":[") + """Priorities"":[".Length
        While position >= ServerResponse.IndexOf("""Priorities"":[") + """Priorities"":[".Length
            Try
                position = ServerResponse.IndexOf("""Id"":", position) + """Id"":".Length
                If position <> -1 Then
                    Dim endPosition As Integer = ServerResponse.IndexOf(",", position + 1)
                    If endPosition <> -1 Then
                        ReDim Preserve Priorities(Count)
                        Dim Priority As New Leankit.Priority
                        Priority.Id = ServerResponse.Substring(position, endPosition - position)
                        Priority.Name = ServerResponse.Substring(ServerResponse.IndexOf("""Name"":""", position) + """Name"":""".Length, (ServerResponse.IndexOf("""}", position)) - (ServerResponse.IndexOf("""Name"":""", position) + """Name"":""".Length))
                        Priorities(Count) = Priority
                        Count += 1
                    End If
                    position = endPosition
                End If
            Catch ex As Exception
                MsgBox(Count)
                ReDim Preserve Priorities(Count - 1)
                Exit While
            End Try
        End While
    End Sub

    Sub CardTypesPopulate(ByRef ServerResponse As String)
        Dim Count As Integer = 0
        Dim position As Integer = ServerResponse.IndexOf("""CardTypes"":[") + """CardTypes"":[".Length
        While position <> -1
            Try
                position = ServerResponse.IndexOf("""Id"":", position) + """Id"":".Length
                If position <> -1 Then
                    Dim endPosition As Integer = ServerResponse.IndexOf(",""", position + 1)
                    If endPosition <> -1 Then
                        ReDim Preserve CardTypes(Count)
                        Dim CardType As New Leankit.Type
                        CardType.Id = ServerResponse.Substring(position, endPosition - position)
                        CardType.Name = ServerResponse.Substring(ServerResponse.IndexOf("""Name"":""", position) + """Name"":""".Length, (ServerResponse.IndexOf(""",""ColorHex"":", position)) - (ServerResponse.IndexOf("""Name"":""", position) + """Name"":""".Length))
                        CardTypes(Count) = CardType
                        Count += 1
                    End If
                    position = endPosition
                End If
            Catch ex As Exception
                ReDim Preserve CardTypes(Count - 1)
                Exit While
            End Try
        End While
    End Sub

    'Toolbar Button Actions

    Sub HideShow()
        Dim Button As PictureBox = Me.Controls("TitleBar").Controls("Button")
        Dim Toolbar As Control = Me.Controls("Toolbar")
        Dim Bitmap As Bitmap = CType(Bitmap.FromFile(Button.ImageLocation), Bitmap)
        Button.Image = Bitmap
        If Toolbar.Visible Then
            Button.Image.RotateFlip(RotateFlipType.RotateNoneFlipNone)
            Toolbar.Hide()
        ElseIf Not Toolbar.Visible Then
            Button.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            Toolbar.Show()
        End If
    End Sub

    Sub LogoutForm()
        Login.Show()
        Me.Hide()
        Me.Dispose()
    End Sub

    Sub UserToolbar()
        Dim Toolbar As Panel = Me.Controls("Toolbar")
        Toolbar.Controls.Clear()
        Toolbar.AutoScroll = True

        Dim Title As New Label
        Title.Name = "Title"
        Title.AutoSize = False
        Title.Location = New Point(37.5, 12.5)
        Title.Size = New Size(Toolbar.Width - 75, 25)
        Title.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        Title.Text = "User Information"
        Title.TextAlign = ContentAlignment.MiddleCenter
        Title.ForeColor = Color.DarkGray
        Toolbar.Controls.Add(Title)

        Dim Close As New PictureBox
        Close.Name = "Close"
        Close.Size = New Size(25, 25)
        Close.Location = New Point(Toolbar.Width - Close.Width - 12.5, Title.Location.Y)
        Close.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        Close.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\BlackX.png"
        Close.BorderStyle = BorderStyle.None
        Close.SizeMode = PictureBoxSizeMode.StretchImage
        Close.Cursor = Cursors.Hand
        Toolbar.Controls.Add(Close)
        Close.Load()
        AddHandler Close.Click, AddressOf ToolbarPopulate

        Dim User As New PictureBox
        User.Name = "User"
        User.Size = New Size(100, 100)
        User.Location = New Point((Toolbar.Width - User.Width) / 2, Title.Location.Y + Title.Height + 15)
        User.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        User.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\User.png"
        User.BorderStyle = BorderStyle.None
        User.SizeMode = PictureBoxSizeMode.StretchImage
        User.Cursor = Cursors.Hand
        Toolbar.Controls.Add(User)
        User.Load()

        Dim Account As New Label
        Account.Name = "Title"
        Account.AutoSize = False
        Account.TextAlign = ContentAlignment.TopCenter
        Account.Size = New Size(Toolbar.Width - 20, 25)
        Account.Location = New Point((Toolbar.Width - Account.Width) / 2, User.Location.Y + User.Height + 30)
        Account.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        Account.Text = "Logged in as" & vbCrLf & Leankit.Account.FullName
        'Account.TextAlign = ContentAlignment.TopCenter
        Account.ForeColor = Color.Black
        Toolbar.Controls.Add(Account)

        Dim Logout As New Button
        Logout.Size = New Size(Toolbar.Width - 20, 25)
        Logout.Location = New Point((Toolbar.Width - Logout.Width) / 2, Account.Location.Y + Account.Height + 15)
        Logout.Name = "Logout"
        Logout.Text = "Logout"
        Logout.Cursor = Cursors.Hand
        Logout.FlatStyle = FlatStyle.Flat
        Logout.BackColor = Color.FromArgb(150, 201, 61)
        Logout.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        Toolbar.Controls.Add(Logout)

        AddHandler Logout.Click, AddressOf LogoutForm

        If Not Toolbar.Visible Then
            HideShow()
        End If
    End Sub

    Sub AddToolbar()
        Dim Toolbar As Panel = Me.Controls("Toolbar")
        Toolbar.Controls.Clear()
        Toolbar.AutoScroll = True

        Dim Title As New Label
        Title.Name = "Title"
        Title.AutoSize = False
        Title.Location = New Point(37.5, 12.5)
        Title.Size = New Size(Toolbar.Width - 75, 25)
        Title.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        Title.Text = "Add New Card"
        Title.TextAlign = ContentAlignment.MiddleCenter
        Title.ForeColor = Color.DarkGray
        Toolbar.Controls.Add(Title)

        Dim Close As New PictureBox
        Close.Name = "Close"
        Close.Size = New Size(25, 25)
        Close.Location = New Point(Toolbar.Width - Close.Width - 12.5, Title.Location.Y)
        Close.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        Close.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\BlackX.png"
        Close.BorderStyle = BorderStyle.None
        Close.SizeMode = PictureBoxSizeMode.StretchImage
        Close.Cursor = Cursors.Hand
        Toolbar.Controls.Add(Close)
        Close.Load()
        AddHandler Close.Click, AddressOf ToolbarPopulate

        Dim Search As New TextBox
        Search.Name = "Search"
        Search.Text = "Search"
        Search.Size = New Size(Toolbar.Width - 50, 25)
        Search.Location = New Point((Toolbar.Width - Search.Width) / 2, Title.Location.Y + Title.Height + 12.5)
        Search.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Toolbar.Controls.Add(Search)
        AddHandler Search.TextChanged, AddressOf Filter
        AddHandler Search.Click, AddressOf FilterTextSelection
        AddHandler Search.Enter, AddressOf FilterTextSelection

        Dim Tree As New TreeView
        Tree.Name = "Tree"
        Tree.Location = New Point(Search.Location.X, Search.Location.Y + Search.Height + 12.5)
        Tree.Size = New Size(Search.Width, 500)
        Tree.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Tree.BackColor = Color.White
        Toolbar.Controls.Add(Tree)
        'AddHandler Tree.AfterSelect, AddressOf PartPopulate
        TreePopulate("")

        Dim Cancel As New Button
        Cancel.Size = New Size(Search.Width / 2 - 6.25, 25)
        Cancel.Location = New Point(Search.Location.X, Tree.Location.Y + Tree.Height + 15)
        Cancel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Cancel.Name = "Cancel"
        Cancel.Text = "Cancel"
        Cancel.Cursor = Cursors.Hand
        Cancel.FlatStyle = FlatStyle.Flat
        Cancel.BackColor = Color.FromArgb(150, 201, 61)
        Cancel.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        Toolbar.Controls.Add(Cancel)
        AddHandler Cancel.Click, AddressOf ToolbarPopulate

        Dim bNext As New Button
        bNext.Size = Cancel.Size
        bNext.Location = New Point(Search.Location.X + (Search.Width / 2 + 6.25), Tree.Location.Y + Tree.Height + 15)
        bNext.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        bNext.Name = "Next"
        bNext.Text = "Next"
        bNext.Cursor = Cursors.Hand
        bNext.FlatStyle = FlatStyle.Flat
        bNext.BackColor = Color.FromArgb(150, 201, 61)
        bNext.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        Toolbar.Controls.Add(bNext)
        AddHandler bNext.Click, AddressOf AddCardToolbar

        If Not Toolbar.Visible Then
            HideShow()
        End If
    End Sub

    Sub AddCardToolbar()
        Dim Toolbar As Panel = Me.Controls("Toolbar")
        Dim Tree As TreeView = Me.Controls("Toolbar").Controls("Tree")
        Dim oTitle As Label = Me.Controls("Toolbar").Controls("Title")
        Dim sNode As String = Nothing
        Toolbar.AutoScroll = True

        If Tree.SelectedNode Is Nothing Then
            oTitle.Text = "Please make a selection"
            oTitle.ForeColor = Color.Red
        Else
            sNode = Tree.SelectedNode.Text
            Toolbar.Controls.Clear()

            Dim nTitle As New Label
            nTitle.Name = "Title"
            nTitle.AutoSize = False
            nTitle.Location = New Point(37.5, 12.5)
            nTitle.Size = New Size(Toolbar.Width - 75, 25)
            nTitle.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
            nTitle.Text = "Add New Card"
            nTitle.TextAlign = ContentAlignment.MiddleCenter
            nTitle.ForeColor = Color.DarkGray
            Toolbar.Controls.Add(nTitle)

            Dim Close As New PictureBox
            Close.Name = "Close"
            Close.Size = New Size(25, 25)
            Close.Location = New Point(Toolbar.Width - Close.Width - 12.5, nTitle.Location.Y)
            Close.Anchor = AnchorStyles.Right Or AnchorStyles.Top
            Close.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\BlackX.png"
            Close.BorderStyle = BorderStyle.None
            Close.SizeMode = PictureBoxSizeMode.StretchImage
            Close.Cursor = Cursors.Hand
            Toolbar.Controls.Add(Close)
            Close.Load()
            AddHandler Close.Click, AddressOf ToolbarPopulate

            Dim Layout As New TableLayoutPanel
            Layout.Name = "Layout"
            Layout.Size = New Size(Toolbar.Width - 50, 500)
            Layout.Location = New Point(12.5, nTitle.Location.Y + nTitle.Height + 12.5)
            Layout.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top Or AnchorStyles.Bottom
            Layout.RowCount = 10
            Layout.ColumnCount = 1
            Toolbar.Controls.Add(Layout)


            Dim Title As New TextBox
            Title.Name = "cTitle"
            Title.Multiline = True
            Title.Size = New Size(Layout.Width - 5, 35)
            Title.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
            Layout.Controls.Add(Title)

            Dim TitleLabel As New Label
            TitleLabel.Size = New Size(Title.Width, 15)
            TitleLabel.Text = "Title"
            TitleLabel.ForeColor = Color.LightGray
            TitleLabel.BackColor = Color.Transparent
            Layout.Controls.Add(TitleLabel)

            Dim Description As New RichTextBox
            Description.Name = "Description"
            Description.Multiline = True
            Description.Size = New Size(Title.Width, 65)
            Description.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
            Layout.Controls.Add(Description)

            Dim DescriptionLabel As New Label
            DescriptionLabel.Size = New Size(Title.Width, 15)
            DescriptionLabel.Text = "Description"
            DescriptionLabel.ForeColor = Color.LightGray
            DescriptionLabel.BackColor = Color.Transparent
            Layout.Controls.Add(DescriptionLabel)

            Dim CardType As New ComboBox
            CardType.Name = "Type"
            CardType.Size = New Size(Title.Width, 20)
            CardType.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
            CardType.DropDownStyle = ComboBoxStyle.DropDownList
            Layout.Controls.Add(CardType)

            For x = 0 To CardTypes.Count - 1
                CardType.Items.Add(CardTypes(x))
            Next

            Dim CardTypeLabel As New Label
            CardTypeLabel.Size = New Size(Title.Width, 15)
            CardTypeLabel.Text = "Card Type"
            CardTypeLabel.ForeColor = Color.LightGray
            CardTypeLabel.BackColor = Color.Transparent
            Layout.Controls.Add(CardTypeLabel)

            Dim Quantity As New TextBox
            Quantity.Name = "Size"
            Quantity.Multiline = True
            Quantity.Size = New Size(Title.Width, 20)
            Quantity.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
            Layout.Controls.Add(Quantity)

            Dim QuantityLabel As New Label
            QuantityLabel.Size = New Size(Title.Width, 15)
            QuantityLabel.Text = "Quantity"
            QuantityLabel.ForeColor = Color.LightGray
            QuantityLabel.BackColor = Color.Transparent
            Layout.Controls.Add(QuantityLabel)

            Dim Finish As New TextBox
            Finish.Name = "DueDate"
            Finish.Multiline = True
            Finish.Size = New Size(Title.Width, 20)
            Finish.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
            Layout.Controls.Add(Finish)

            Dim FinishLabel As New Label
            FinishLabel.Size = New Size(Title.Width, 15)
            FinishLabel.Text = "Due Date"
            FinishLabel.ForeColor = Color.LightGray
            FinishLabel.BackColor = Color.Transparent
            Layout.Controls.Add(FinishLabel)

            Dim Tags As New TextBox
            Tags.Name = "Tags"
            Tags.Multiline = True
            Tags.Size = New Size(Title.Width, 20)
            Tags.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
            Layout.Controls.Add(Tags)

            Dim TagsLabel As New Label
            TagsLabel.Size = New Size(Title.Width, 15)
            TagsLabel.Text = "Tags"
            TagsLabel.ForeColor = Color.LightGray
            TagsLabel.BackColor = Color.Transparent
            Layout.Controls.Add(TagsLabel)

            Dim CardID As New TextBox
            CardID.Name = "ExternalCardID"
            CardID.Multiline = True
            CardID.Size = New Size(Title.Width, 20)
            CardID.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
            Layout.Controls.Add(CardID)

            Dim CardIDLabel As New Label
            CardIDLabel.Size = New Size(Title.Width, 15)
            CardIDLabel.Text = "Lot - Run"
            CardIDLabel.ForeColor = Color.LightGray
            CardIDLabel.BackColor = Color.Transparent
            Layout.Controls.Add(CardIDLabel)

            Dim Priority As New ComboBox
            Priority.Name = "Priority"
            Priority.Size = New Size(Title.Width, 20)
            Priority.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
            Priority.DropDownStyle = ComboBoxStyle.DropDownList
            Layout.Controls.Add(Priority)

            PrioritiesPopulate()

            For x = 0 To Priorities.Count - 1
                Priority.Items.Add(Priorities(x))
            Next

            Dim PriorityLabel As New Label
            PriorityLabel.Size = New Size(Title.Width, 15)
            PriorityLabel.Text = "Priority"
            PriorityLabel.ForeColor = Color.LightGray
            PriorityLabel.BackColor = Color.Transparent
            Layout.Controls.Add(PriorityLabel)

            Dim Lane As New ComboBox
            Lane.Name = "LaneTitle"
            Lane.Size = New Size(Title.Width, 20)
            Lane.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
            Lane.DropDownStyle = ComboBoxStyle.DropDownList
            Layout.Controls.Add(Lane)

            For x = 0 To Lanes.Count - 1
                Lane.Items.Add(Lanes(x).Title)
            Next

            Dim LaneLabel As New Label
            LaneLabel.Size = New Size(Title.Width, 15)
            LaneLabel.Text = "Lane"
            LaneLabel.ForeColor = Color.LightGray
            LaneLabel.BackColor = Color.Transparent
            Layout.Controls.Add(LaneLabel)

            Dim Submit As New Button
            Submit.Size = New Size(Title.Width, 25)
            Submit.Name = "Submit"
            Submit.Text = "Submit"

            Submit.Cursor = Cursors.Hand
            Submit.FlatStyle = FlatStyle.Flat
            Submit.BackColor = Color.FromArgb(150, 201, 61)
            Submit.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
            Layout.Controls.Add(Submit, 0, 18)
        End If
    End Sub

    Sub ExportToolbar()
        Dim Toolbar As Panel = Me.Controls("Toolbar")
        Toolbar.Controls.Clear()
        Toolbar.AutoScroll = True

        Dim Title As New Label
        Title.Name = "Title"
        Title.AutoSize = False
        Title.Location = New Point(37.5, 12.5)
        Title.Size = New Size(Toolbar.Width - 75, 25)
        Title.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        Title.Text = "User Information"
        Title.TextAlign = ContentAlignment.MiddleCenter
        Title.ForeColor = Color.DarkGray
        Toolbar.Controls.Add(Title)

        Dim Close As New PictureBox
        Close.Name = "Close"
        Close.Size = New Size(25, 25)
        Close.Location = New Point(Toolbar.Width - Close.Width - 12.5, Title.Location.Y)
        Close.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        Close.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\BlackX.png"
        Close.BorderStyle = BorderStyle.None
        Close.SizeMode = PictureBoxSizeMode.StretchImage
        Close.Cursor = Cursors.Hand
        Toolbar.Controls.Add(Close)
        Close.Load()
        AddHandler Close.Click, AddressOf ToolbarPopulate

        Dim Labels As New Button
        Labels.Size = New Size(Toolbar.Width - 20, 25)
        Labels.Location = New Point((Toolbar.Width - Labels.Width) / 2, Title.Location.Y + Title.Height + 15)
        Labels.Name = "Labels"
        Labels.Text = "Export Labels"
        Labels.Cursor = Cursors.Hand
        Labels.FlatStyle = FlatStyle.Flat
        Labels.BackColor = Color.FromArgb(150, 201, 61)
        Labels.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        Toolbar.Controls.Add(Labels)
        AddHandler Labels.Click, AddressOf LabelExport

        Dim sWeight As New Button
        sWeight.Size = New Size(Toolbar.Width - 20, 25)
        sWeight.Location = New Point((Toolbar.Width - Labels.Width) / 2, Labels.Location.Y + Labels.Height + 15)
        sWeight.Name = "Weights"
        sWeight.Text = "Export Shipping Weights"
        sWeight.Cursor = Cursors.Hand
        sWeight.FlatStyle = FlatStyle.Flat
        sWeight.BackColor = Color.FromArgb(150, 201, 61)
        sWeight.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        Toolbar.Controls.Add(sWeight)
        AddHandler sWeight.Click, AddressOf WeightExport

        If Not Toolbar.Visible Then
            HideShow()
        End If
    End Sub

    Sub RunSelect()
        Dim Toolbar As Panel = Me.Controls("Toolbar")
        Toolbar.Controls.Clear()
        Toolbar.AutoScroll = False

        Dim Title As New Label
        Title.Name = "Title"
        Title.AutoSize = False
        Title.Location = New Point(37.5, 12.5)
        Title.Size = New Size(Toolbar.Width - 75, 25)
        Title.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right
        Title.Text = "Select the run and cards to export"
        Title.TextAlign = ContentAlignment.MiddleCenter
        Title.ForeColor = Color.DarkGray
        Toolbar.Controls.Add(Title)

        Dim Close As New PictureBox
        Close.Name = "Close"
        Close.Size = New Size(25, 25)
        Close.Location = New Point(Toolbar.Width - Close.Width - 12.5, Title.Location.Y)
        Close.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        Close.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\BlackX.png"
        Close.BorderStyle = BorderStyle.None
        Close.SizeMode = PictureBoxSizeMode.StretchImage
        Close.Cursor = Cursors.Hand
        Toolbar.Controls.Add(Close)
        Close.Load()
        AddHandler Close.Click, AddressOf ToolbarPopulate

        Dim Run As New ComboBox
        Run.Name = "Run"
        Run.Text = "Select a Run"
        Run.Size = New Size(Toolbar.Width - 50, 25)
        Run.Location = New Point((Toolbar.Width - Run.Width) / 2, Title.Location.Y + Title.Height + 12.5)
        Run.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        Toolbar.Controls.Add(Run)
        AddHandler Run.TextChanged, AddressOf LabelPopulate
        Dim Count As Integer = 0

        For x = CardTypes.Count - 1 To 0 Step -1
            If CardTypes(x).Name.Substring(0, 3) = "Run" And Count < 4 Then
                Run.Items.Add(CardTypes(x).Name)
                Count += 1
            End If
        Next

        Dim List As New TreeView
        List.Name = "List"
        List.Location = New Point(Run.Location.X, Run.Location.Y + Run.Height + 12.5)
        'List.Size = New Size(Run.Width, 500)
        List.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        List.BackColor = Color.White
        List.CheckBoxes = True
        Toolbar.Controls.Add(List)

        Dim Cancel As New Button
        Cancel.Size = New Size(Run.Width / 2 - 6.25, 25)
        Cancel.Location = New Point(Run.Location.X, Toolbar.Height - Cancel.Height - 12.5)
        List.Size = New Size(Run.Width, Toolbar.Height - List.Location.Y - Cancel.Height - 25)
        Cancel.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        Cancel.Name = "Cancel"
        Cancel.Text = "Cancel"
        Cancel.Cursor = Cursors.Hand
        Cancel.FlatStyle = FlatStyle.Flat
        Cancel.BackColor = Color.FromArgb(150, 201, 61)
        Cancel.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        Toolbar.Controls.Add(Cancel)
        AddHandler Cancel.Click, AddressOf ToolbarPopulate

        Dim Export As New Button
        Export.Size = Cancel.Size
        Export.Location = New Point(Run.Location.X + (Run.Width / 2 + 6.25), Cancel.Location.Y)
        Export.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        Export.Name = "Export"
        Export.Text = "Export"
        Export.Cursor = Cursors.Hand
        Export.FlatStyle = FlatStyle.Flat
        Export.BackColor = Color.FromArgb(150, 201, 61)
        Export.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        Toolbar.Controls.Add(Export)
    End Sub

    Sub LabelExport()
        RunSelect()
        Dim Export As Button = Me.Controls("Toolbar").Controls("Export")
        AddHandler Export.Click, AddressOf ExcelLables
    End Sub

    Sub WeightExport()
        RunSelect()
        Dim Export As Button = Me.Controls("Toolbar").Controls("Export")
        AddHandler Export.Click, AddressOf ExcelWeights
    End Sub

    Sub LabelPopulate()
        Dim Toolbar As Panel = Me.Controls("Toolbar")
        Dim List As TreeView = Toolbar.Controls("List")
        Dim Run As ComboBox = Toolbar.Controls("Run")
        Dim Part As TreeNode
        List.Nodes.Clear()
        List.CheckBoxes = True
        Try
            For x = 0 To Cards.Count - 1
                If Cards(x).ExternalCardID.Length > 5 Then
                    If Cards(x).TypeName.Substring(4, 5) = Run.Text.Substring(4, 5) Then
                        Part = List.Nodes.Add(Cards(x).Title)
                        Part.Tag = x
                        Part.Checked = True
                    End If
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    Sub ExcelWeights()
        Dim Toolbar As Panel = Me.Controls("Toolbar")
        Dim List As TreeView = Toolbar.Controls("List")
        Dim RunList As ComboBox = Toolbar.Controls("Run")
        Dim PartList(List.Nodes.Count) As String
        Dim Count As Integer = 0
        Dim Part As TreeNode

        Dim xlApp As Excel.Application = New Excel.Application()

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
            Return
        End If

        Dim SFD As New SaveFileDialog
        SFD.Filter = "Excel Documents | *.xls"
        SFD.DefaultExt = "xls"
        SFD.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        SFD.OverwritePrompt = True
        SFD.FileName = RunList.Text & " Shipping Weights"

        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\ShippingWeights.xlsx")
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Worksheets("sheet1")
        Dim xlListObject As Excel.ListObject = xlWorkSheet.ListObjects("Table1")
        Dim xlRange As Excel.Range = xlListObject.DataBodyRange
        Dim WeightArray(xlListObject.ListRows.Count, xlListObject.ListColumns.Count)
        WeightArray = xlRange.Value


        Try
            If SFD.ShowDialog = DialogResult.OK Then
                ReDim Preserve PartList(Count)
                For Each Part In List.Nodes
                    If Part.Checked Then
                        For x = 0 To Cards.Count - 1
                            If x = Part.Tag Then
                                ReDim Preserve PartList(Count)
                                PartList(Count) = Part.Tag
                                Count += 1
                            End If
                        Next
                    End If
                Next

                Dim DataArray(PartList.Count - 1, 7)
                Dim r As Integer
                For r = 0 To PartList.Count - 1
                    DataArray(r, 0) = Cards(PartList(r)).Title
                    DataArray(r, 1) = Cards(PartList(r)).LotNumber 'ExternalCardID.Substring(Cards(PartList(r)).ExternalCardID.Length - 6, 6)
                    DataArray(r, 2) = CardSize(Cards(PartList(r)).Title, Cards(PartList(r)).LotNumber, Cards(PartList(r)).Size)
                    'DataArray(r, 2) = sECardID(PartList(r)).Substring(sECardID(PartList(r)).Length - 6, 6)
                    For n = 1 To xlListObject.ListRows.Count
                        If InStr(Cards(PartList(r)).Title, WeightArray(n, 1)) > 0 Then
                            DataArray(r, 3) = WeightArray(n, 2)
                            DataArray(r, 4) = WeightArray(n, 3)
                            DataArray(r, 5) = WeightArray(n, 4)
                            DataArray(r, 6) = WeightArray(n, 5)
                        End If
                    Next
                Next

                xlApp.Quit()
                xlApp = New Excel.Application()
                xlRange = Nothing
                xlWorkBook = xlApp.Workbooks.Add
                xlWorkSheet = xlWorkBook.Worksheets(1)
                xlWorkSheet.Name = RunList.Text

                xlWorkSheet.Range("A1").Value = "Part Number"
                xlWorkSheet.Range("B1").Value = "Lot Number"
                xlWorkSheet.Range("C1").Value = "Quantity"
                xlWorkSheet.Range("D1").Value = "Wire Type"
                xlWorkSheet.Range("E1").Value = "Needles / Set"
                xlWorkSheet.Range("F1").Value = "Sets / Bag"
                xlWorkSheet.Range("G1").Value = "Weight / Bag"
                xlWorkSheet.Range("H1").Value = "Total Needles"
                xlWorkSheet.Range("I1").Value = "Total Bags"
                xlWorkSheet.Range("J1").Value = "Total Weight"

                xlRange = xlWorkSheet.Range("A1").Resize(PartList.Count + 1, 10)
                xlWorkSheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, xlRange,, Excel.XlYesNoGuess.xlYes).Name = "CurrentRunLabels"
                xlListObject = xlWorkSheet.ListObjects("CurrentRunLabels")

                xlListObject.ListColumns(8).DataBodyRange.Formula = "=IFERROR(SUMPRODUCT([@[Quantity]],[@[Needles / Set]]),)"
                xlListObject.ListColumns(9).DataBodyRange.Formula = "=IFERROR(QUOTIENT([@[Quantity]],[@[Sets / Bag]]),)"
                xlListObject.ListColumns(10).DataBodyRange.Formula = "=IFERROR(PRODUCT([@[Total Bags]],[@[Weight / Bag]]),)"

                xlRange = xlWorkSheet.Range("A2").Resize(PartList.Count, 7)
                xlRange.Value = DataArray

                xlWorkSheet.Range("A:A").ColumnWidth = 40
                xlWorkSheet.Range("B:I").ColumnWidth = 15
                With xlListObject
                    .Range.HorizontalAlignment = Excel.Constants.xlLeft
                    .Sort.SortFields.Clear()
                    .Sort.SortFields.Add(.ListColumns(5).DataBodyRange, Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending, Excel.XlSortDataOption.xlSortNormal)
                    .Sort.SortFields.Add(.ListColumns(4).DataBodyRange, Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlAscending, Excel.XlSortDataOption.xlSortNormal)
                    .Sort.SortFields.Add(.ListColumns(2).DataBodyRange, Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlAscending, Excel.XlSortDataOption.xlSortNormal)
                    .Sort.Header = Excel.XlYesNoGuess.xlYes
                    .Sort.MatchCase = False
                    .Sort.Orientation = Excel.XlSortOrientation.xlSortColumns
                    .Sort.SortMethod = Excel.XlSortMethod.xlPinYin
                    .Sort.Apply()
                End With

                xlWorkSheet = xlWorkBook.Worksheets.Add
                xlWorkSheet.Name = "Shipping Configuration"

                WeightArray = xlListObject.DataBodyRange.Value

                xlWorkSheet.Range("A1").Value = "Box Number"
                xlWorkSheet.Range("B1").Value = "Part Number"
                xlWorkSheet.Range("C1").Value = "Lot Number"
                xlWorkSheet.Range("D1").Value = "Bags in Box"
                xlWorkSheet.Range("E1").Value = "Weight in Box"

                xlWorkSheet.Range("A:A").ColumnWidth = 15
                xlWorkSheet.Range("B:B").ColumnWidth = 40
                xlWorkSheet.Range("C:E").ColumnWidth = 15

                Run.Reset()
                Dim Box As Integer = 1
                Dim Row As Integer = 0

                ReDim DataArray(1000, 4)

                For n = 1 To WeightArray.GetLength(0)
                    Run.Weight += WeightArray(n, 10)
                Next
                If Run.Weight > Run.WeightTrigger Then Run.BoxCount = 30
                Run.AverageWeight = Run.Weight / Run.BoxCount

                Dim Tolerance As New ShippingTolerance
                If Tolerance.ShowDialog() = DialogResult.OK Then

                    DataArray(0, 0) = "Box 1"
                    For x = 1 To PartList.Count
                        If WeightArray(x, 9) <> 0 Then
                            For y = 1 To WeightArray(x, 9)
                                If (WeightArray(x, 1) <> DataArray(Row, 1)) Or (CStr(WeightArray(x, 2)) <> CStr(DataArray(Row, 2))) Then
                                    Row += 1
                                End If
                                DataArray(Row, 1) = WeightArray(x, 1)
                                DataArray(Row, 2) = WeightArray(x, 2)
                                DataArray(Row, 3) += 1
                                DataArray(Row, 4) += WeightArray(x, 7)
                                Run.BoxWeight += WeightArray(x, 7)
                                '((AverageWeight - BoxWeight) < 0.5 Or (BoxWeight - AverageWeight) > 2)
                                If Run.BoxWeight > Run.MaxWeight Then
                                    DataArray(Row, 4) -= WeightArray(x, 7)
                                    Run.BoxWeight -= WeightArray(x, 7)

                                    If DataArray(Row, 3) > 1 Then
                                        DataArray(Row, 3) -= 1
                                        Row += 1
                                    Else
                                        DataArray(Row, 1) = Nothing
                                        DataArray(Row, 2) = Nothing
                                    End If

                                    DataArray(Row, 3) = "Total"
                                    DataArray(Row, 4) = Run.BoxWeight
                                    Box += 1
                                    Row += 1
                                    DataArray(Row, 0) = "Box " & Box
                                    Run.BoxWeight = 0
                                    y -= 1
                                ElseIf ((((Run.AverageWeight - Run.BoxWeight) < Run.LowTolerance And (Run.AverageWeight - Run.BoxWeight) > 0) Or (Run.BoxWeight - Run.AverageWeight) > Run.HighTolerance) And Run.BoxWeight >= Run.MinWeight) And (x <> PartList.Count And y <> WeightArray(x, 9)) Then
                                    Row += 1
                                    DataArray(Row, 3) = "Total"
                                    DataArray(Row, 4) = Run.BoxWeight
                                    Box += 1
                                    Row += 1
                                    DataArray(Row, 0) = "Box " & Box
                                    Run.BoxWeight = 0
                                End If
                                'End If
                            Next
                        End If
                    Next

                    Row += 1
                    DataArray(Row, 3) = "Total"
                    DataArray(Row, 4) = Run.BoxWeight

                    xlRange = xlWorkSheet.Range("A2").Resize(DataArray.GetLength(0), 5)
                    xlRange.Value = DataArray
                    xlRange = xlWorkSheet.Range("A2").Resize(DataArray.GetLength(0), 8)
                    xlRange.HorizontalAlignment = Excel.Constants.xlLeft

                    xlWorkSheet.Rows(1).Insert()
                    xlWorkSheet.Rows(1).Insert()
                    xlWorkSheet.Rows(1).Insert()
                    xlWorkSheet.Rows(1).Insert()

                    xlWorkSheet.Range("A1").Value = "Total Weight"
                    xlWorkSheet.Range("A2").Value = "Box Count"
                    xlWorkSheet.Range("A3").Value = "Average Weight"

                    xlWorkSheet.Range("B1:B3").NumberFormat = "0.00"
                    xlWorkSheet.Range("E:E").NumberFormat = "0.00"

                    xlWorkSheet.Range("B1").Value = Run.Weight
                    xlWorkSheet.Range("B2").Value = Run.BoxCount
                    xlWorkSheet.Range("B3").Value = Run.AverageWeight

                    xlRange = xlWorkSheet.Range("A1:A3")
                    With xlRange
                        With .Interior
                            .Pattern = Excel.XlPattern.xlPatternSolid
                            .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
                            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
                            .TintAndShade = 0.399975585192419
                            .PatternTintAndShade = 0
                        End With
                        With .Borders(Excel.XlBordersIndex.xlEdgeBottom)
                            .LineStyle = Excel.XlLineStyle.xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = Excel.XlBorderWeight.xlThin
                        End With
                    End With

                    xlRange = xlWorkSheet.Range("B1:B3")
                    With xlRange
                        With .Interior
                            .Pattern = Excel.XlPattern.xlPatternSolid
                            .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
                            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
                            .TintAndShade = 0.599993896298105
                            .PatternTintAndShade = 0
                        End With
                        With .Borders(Excel.XlBordersIndex.xlEdgeBottom)
                            .LineStyle = Excel.XlLineStyle.xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = Excel.XlBorderWeight.xlThin
                        End With
                    End With

                    xlRange = xlWorkSheet.Range("A5:E5")
                    With xlRange.Resize(1, 5)
                        With .Interior
                            .Pattern = Excel.XlPattern.xlPatternSolid
                            .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
                            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
                            .TintAndShade = 0.399975585192419
                            .PatternTintAndShade = 0
                        End With
                        With .Borders(Excel.XlBordersIndex.xlEdgeBottom)
                            .LineStyle = Excel.XlLineStyle.xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = Excel.XlBorderWeight.xlThin
                        End With
                    End With

                    xlRange = xlWorkSheet.Range("A6").Resize(DataArray.GetLength(0), 1)
                    Dim xltRange As Excel.Range
                    For Each xltRange In xlRange.Cells
                        If InStr(xltRange.Value, "Box") > 0 Then
                            With xltRange.Resize(1, 5)
                                With .Interior
                                    .Pattern = Excel.XlPattern.xlPatternSolid
                                    .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
                                    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
                                    .TintAndShade = 0.599993896298105
                                    .PatternTintAndShade = 0
                                End With
                                With .Borders(Excel.XlBordersIndex.xlEdgeBottom)
                                    .LineStyle = Excel.XlLineStyle.xlContinuous
                                    .ColorIndex = 0
                                    .TintAndShade = 0
                                    .Weight = Excel.XlBorderWeight.xlThin
                                End With
                            End With
                        End If
                    Next


                    xlApp.PrintCommunication = False
                    xlWorkSheet.PageSetup.FitToPagesWide = 1
                    xlWorkSheet.PageSetup.FitToPagesTall = False
                    xlApp.PrintCommunication = True

                    xlApp.DisplayAlerts = False
                    xlWorkBook.SaveAs(SFD.FileName)
                    xlApp.Visible = True
                    xlApp.Workbooks.Open(SFD.FileName)
                    xlApp.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized
                End If
            End If
        Catch ex As Exception
            MsgBox("There was an error processing the file.")
        End Try

        xlWorkSheet = Nothing
        xlWorkBook = Nothing
        'xlApp.Quit()
        xlApp = Nothing
        GC.Collect()

    End Sub

    Sub ExcelLables()
        Dim Toolbar As Panel = Me.Controls("Toolbar")
        Dim List As TreeView = Toolbar.Controls("List")
        Dim Run As ComboBox = Toolbar.Controls("Run")
        Dim PartList(List.Nodes.Count) As String
        Dim Count As Integer = 0
        Dim Part As TreeNode

        Dim xlApp As Excel.Application = New Excel.Application()

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
            Return
        End If

        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim xlListObject As Excel.ListObject
        Dim xlRange As Excel.Range
        Dim SFD As New SaveFileDialog
        SFD.Filter = "Excel Documents | *.xls"
        SFD.DefaultExt = "xls"
        SFD.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        SFD.OverwritePrompt = True
        SFD.FileName = Run.Text & " Labels"

        Try
            If SFD.ShowDialog = DialogResult.OK Then
                'ReDim Preserve PartList(Count)
                For Each Part In List.Nodes
                    If Part.Checked Then
                        For x = 0 To Cards.Count - 1
                            If x = Part.Tag Then
                                ReDim Preserve PartList(Count)
                                PartList(Count) = Part.Tag
                                Count += 1
                            End If
                        Next
                    End If
                Next

                xlWorkBook = xlApp.Workbooks.Add
                xlWorkSheet = xlWorkBook.Worksheets(1)

                Dim DataArray(PartList.Count - 1, 6)
                Dim r As Integer
                For r = 0 To PartList.Count - 1
                    DataArray(r, 0) = Cards(PartList(r)).Title
                    For s = 0 To Lanes.Count - 1
                        If Lanes(s).Id = Cards(PartList(r)).LaneId Then DataArray(r, 1) = Lanes(s).Title
                    Next
                    DataArray(r, 2) = Cards(PartList(r)).LotNumber 'Cards(PartList(r)).ExternalCardID.Substring(Cards(PartList(r)).ExternalCardID.Length - 6, 6)
                    DataArray(r, 3) = Cards(PartList(r)).TypeName
                    DataArray(r, 4) = Cards(PartList(r)).Size
                    DataArray(r, 5) = CardSize(Cards(PartList(r)).Title, Cards(PartList(r)).LotNumber, Cards(PartList(r)).Size)
                Next


                xlWorkSheet.Range("A1").Value = "Part Number"
                xlWorkSheet.Range("B1").Value = "Lane"
                xlWorkSheet.Range("C1").Value = "Lot Number"
                xlWorkSheet.Range("D1").Value = "Priority"
                xlWorkSheet.Range("E1").Value = "Size"
                xlWorkSheet.Range("F1").Value = "Label Quantity"

                xlRange = xlWorkSheet.Range("A2").Resize(PartList.Count, 6)
                xlRange.Value = DataArray
                xlRange = xlWorkSheet.Range("A1").Resize(PartList.Count + 1, 6)
                xlWorkSheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, xlRange,, Excel.XlYesNoGuess.xlYes).Name = "CurrentRunLabels"
                xlListObject = xlWorkSheet.ListObjects("CurrentRunLabels")
                xlWorkSheet.Range("A:A").ColumnWidth = 40
                xlWorkSheet.Range("B:B").ColumnWidth = 20
                xlWorkSheet.Range("C:E").ColumnWidth = 11
                xlListObject.Range.HorizontalAlignment = Excel.Constants.xlLeft
                'xlWorkSheet.PageSetup.PrintArea = "CurrentRunLabels"
                xlApp.DisplayAlerts = False
                xlWorkBook.SaveAs(SFD.FileName)
                'MsgBox("The file was sucessfully exported!")
                xlApp.Visible = True
                xlApp.Workbooks.Open(SFD.FileName)
                xlApp.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized

                'xlApp.PrintCommunication = False
                'xlWorkSheet.PageSetup.FitToPagesWide = 1
                'xlWorkSheet.PageSetup.FitToPagesTall = False
                'xlApp.PrintCommunication = True
            End If
        Catch ex As Exception
            MsgBox("There was an error processing the file.")
        End Try

        'xlWorkSheet.PrintOutEx(From:=1, To:=1, Copies:=1, Collate:=True)
        xlWorkSheet = Nothing
        xlWorkBook = Nothing
        'xlApp.Quit()
        xlApp = Nothing
        GC.Collect()

    End Sub

    Function CardSize(ByVal Title As String, ByVal Lot As String, ByVal Size As Integer)
        Dim Pouches As String = String.Empty
        Dim position As Integer = 0
        Dim start As Integer = 0
        Dim BoxSize As String = String.Empty
        Dim PS As String = String.Empty
        Dim Ribbon As String = String.Empty

        Try
            If Title.Substring(Title.IndexOf("-") - 2, 2) = "90" Then
                PS = Title.Substring(Title.IndexOf("-") - 2, 5).ToUpper
            Else
                PS = Title.Substring(Title.IndexOf("-") - 1, 4).ToUpper
            End If
        Catch ex As Exception
        End Try
        Try
            If Title.Substring(0, 6).ToUpper = "S02244" Then
                Pouches = Size * 10
                Exit Try
            ElseIf Title.Substring(0, 1) = "0" Or Title.Substring(0, 8) = "S0013315" Or Lot.Substring(0, 2).ToUpper = "8F" Or Lot.Substring(0, 2).ToUpper = "3F" Or Title.Substring(0, 6) = "700998" Then
                Pouches = Size * 25
                Exit Try
            ElseIf PS = "5-PS" Or PS = "4-PS" Then
                Pouches = Size * 50
                Exit Try
            End If
            Ribbon = Title.Substring(Title.IndexOf("-") + 1, 1).ToUpper
            If Ribbon = "R" Or Title.Substring(0, 6).ToUpper = "S91072" Or Title.Substring(0, 6) = "S96960" Or Title.Substring(0, 3) = "S81" Or Title.Substring(0, 3) = "S91" Or PS = "90-PS" Or Title.Substring(Title.IndexOf("-") + 1, 1) = "T" Or Title.Substring(Title.IndexOf("-") + 1, 1) = "P" Or Title.Substring(Title.IndexOf("-") + 1, 1) = "H" Or Title.Substring(Title.IndexOf("-") + 1, 1) = "0" Then
                Pouches = Size
            End If
            Try
                BoxSize = Title.Substring(Title.IndexOf("-") + 1, 2)
                If CInt(BoxSize) <> -1 And CInt(BoxSize) <> 0 Then
                    Pouches = CInt(BoxSize) * Size
                End If
            Catch ex As Exception
                position = Title.IndexOf("-")
                BoxSize = Title.Substring(Title.IndexOf("-", position + 1) + 1, 2)
                If CInt(BoxSize) <> -1 Then
                    Pouches = CInt(BoxSize) * Size
                End If
            End Try

        Catch ex As Exception
        End Try
        'If Title.Substring(0, 6) = "700998" Then
        '    MsgBox(Lot)
        '    MsgBox(Size)
        '    MsgBox(Pouches)
        'End If
        Try
            If Title.Substring(0, 6).ToUpper = "S91072" Then
                Pouches = Size
            End If
        Catch ex As Exception

        End Try
        Return Pouches
    End Function

    Private Sub Filter(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Search As TextBox = Me.Controls("Toolbar").Controls("Search")
        If Search.Text <> "Search" Then TreePopulate(Search.Text)
    End Sub

    Private Sub FilterTextSelection(sender As Object, e As System.EventArgs)
        Dim Search As TextBox = Me.Controls("Toolbar").Controls("Search")
        If Search.Text = "Search" Then Search.SelectAll()
    End Sub

    Sub TreePopulate(Filter As String)
        Dim XMLDir As String = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\" & "Product_Library" & ".xml"
        Dim doc As XDocument
        Dim xDoc As New XmlDocument
        Dim Tree As TreeView = Me.Controls("Toolbar").Controls("Tree")
        Try
            doc = XDocument.Load(XMLDir)
        Catch ex As Exception
            Dim Resoucres As New MediaFire
            Resoucres.getFile("lrlfmcjtbwh68r9", "Product_Library", True, ".xml")
            doc = XDocument.Load(XMLDir)
        End Try
        Dim docstring As String = doc.ToString
        xDoc.LoadXml(docstring)

        Tree.Nodes.Clear()
        Dim Item As String
        With xDoc.SelectSingleNode("Parts")
            Dim Items(.ChildNodes.Count - 1) As TreeNode
            For X = 0 To .ChildNodes.Count - 1
                Item = .ChildNodes(X).Name
                If Item.ToUpper.Contains(Filter.ToUpper) Then Items(X) = Tree.Nodes.Add(Item.Replace("_", " "))
            Next
        End With
    End Sub

    Private Sub CloseForm()
        Dim sender As Object = Nothing
        Dim e As New FormClosingEventArgs(CloseReason.UserClosing, True)
        Form_FormClosing(sender, e)
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
        End If
        Dim Close As Form = ManagerClose
        If Close.Visible = False And e.CloseReason = CloseReason.UserClosing Then
            Close.ShowDialog()
        End If
    End Sub
End Class
