Imports System.Environment
Imports System.Net

Public Class Login

    Public Shared ErrorMessage As String = Nothing
    Public Shared LoggedIn As Boolean = False

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GUISetup()
        If Not LoggedIn Then
            LoginGUI()
        Else
            BoardPopulate()
        End If
    End Sub

    Private Sub GUISetup()
        Me.Text = "Rochester Leankit Assistant"
        Me.Size = New Size(500, 300)  'Sets the size of the loading screen GUI
        Me.Top = (Screen.PrimaryScreen.Bounds.Height - Me.Height) / 2
        Me.Left = (Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2
        Me.BackColor = Color.White
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
    End Sub

    Private Sub LoginGUI()
        Me.Controls.Clear()

        Dim Login As New Panel
        Login.Name = "Login"
        Login.Size = Me.ClientSize
        Login.Location = New Point(0, 0)
        Me.Controls.Add(Login)

        Dim Button As New Button
        Button.Text = "Login"
        Button.Size = New Size(400, 25)
        Button.Location = New Point((Login.Width - Button.Width) / 2, 200)
        Button.FlatStyle = FlatStyle.Flat
        Button.BackColor = Color.FromArgb(150, 201, 61)
        Button.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        Button.TabIndex = 4
        AddHandler Button.Click, AddressOf BoardPopulate
        Login.Controls.Add(Button)

        Dim PassLabel As New Label
        PassLabel.Size = New Size(400, 15)
        PassLabel.Location = New Point(Button.Location.X, 185)
        PassLabel.Text = "Password"
        PassLabel.ForeColor = Color.LightGray
        PassLabel.BackColor = Color.Transparent
        Login.Controls.Add(PassLabel)

        Dim PassText As New TextBox
        PassText.Name = "Password"
        PassText.Text = "RGuincho1"
        PassText.Size = New Size(400, 25)
        PassText.Location = New Point(Button.Location.X, 165)
        PassText.BorderStyle = BorderStyle.FixedSingle
        PassText.PasswordChar = "*"
        PassText.TabIndex = 3
        AddHandler PassText.KeyDown, AddressOf EnterCheck
        Login.Controls.Add(PassText)

        Dim UserLabel As New Label
        UserLabel.Size = New Size(400, 15)
        UserLabel.Location = New Point(Button.Location.X, 150)
        UserLabel.Text = "Username"
        UserLabel.ForeColor = Color.LightGray
        UserLabel.BackColor = Color.Transparent
        Login.Controls.Add(UserLabel)

        Dim UserText As New TextBox
        UserText.Name = "Username"
        UserText.Text = "rguincho"
        UserText.Size = New Size(400, 25)
        UserText.Location = New Point(Button.Location.X, 130)
        UserText.BorderStyle = BorderStyle.FixedSingle
        UserText.TabIndex = 2
        AddHandler UserText.KeyDown, AddressOf EnterCheck
        Login.Controls.Add(UserText)

        Dim AddressLabel As New Label
        AddressLabel.Size = New Size(400, 15)
        AddressLabel.Location = New Point(Button.Location.X, 115)
        AddressLabel.Text = "Address"
        AddressLabel.ForeColor = Color.LightGray
        AddressLabel.BackColor = Color.Transparent
        Login.Controls.Add(AddressLabel)

        Dim AddressText As New TextBox
        AddressText.Name = "Account"
        AddressText.Text = "rochestermed"
        AddressText.Size = New Size(400, 25)
        AddressText.Location = New Point(Button.Location.X, 95)
        AddressText.BorderStyle = BorderStyle.FixedSingle
        AddressText.TabIndex = 1
        AddHandler AddressText.KeyDown, AddressOf EnterCheck
        Login.Controls.Add(AddressText)

        Dim Picture As New PictureBox
        Picture.Size = New Size(150, 50)
        Picture.Location = New Point(Button.Location.X, 25)
        Picture.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\Leankit.png"
        Picture.BorderStyle = BorderStyle.None
        Picture.SizeMode = PictureBoxSizeMode.StretchImage
        Login.Controls.Add(Picture)

        Dim Logo As New PictureBox
        Logo.Size = New Size(195, 70)
        Logo.Location = New Point(255, 15)
        Logo.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\Rochester.png"
        Logo.BorderStyle = BorderStyle.None
        Logo.SizeMode = PictureBoxSizeMode.StretchImage
        Login.Controls.Add(Logo)
    End Sub

    Private Sub BoardPopulate()
        Try
            AutoFix()
            Public_Vars()
            Dim Board As New Leankit.Board
            Dim ServerResponse As String = Board.RetrieveAll

            If ServerResponse <> Nothing Then
                Controls.Clear()

                Dim Text As New Label
                Text.Text = "All Boards"
                Text.Font = New Font(Text.Font.FontFamily, 18)
                Text.Size = New Size(400, 50)
                Text.Location = New Point(25, 25)
                Me.Controls.Add(Text)

                Dim Button As New Button
                Button.Text = "Back"
                Button.Size = New Size(150, 25)
                Button.Location = New Point(300, 200)
                Button.FlatStyle = FlatStyle.Flat
                Button.BackColor = Color.FromArgb(150, 201, 61)
                Button.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
                Button.Cursor = Cursors.Hand
                AddHandler Button.Click, AddressOf LoginGUI
                Me.Controls.Add(Button)

                Dim Open As New Button
                Open.Text = "Open"
                Open.Name = "Open"
                Open.Enabled = False
                Open.Size = New Size(150, 25)
                Open.Location = New Point(135, 200)
                Open.FlatStyle = FlatStyle.Flat
                Open.BackColor = Color.FromArgb(150, 201, 61)
                Open.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
                Open.Cursor = Cursors.Hand
                AddHandler Open.Click, AddressOf openBoard
                Me.Controls.Add(Open)

                Dim Panel As New ListView
                Panel.Name = "Panel"
                Panel.Size = New Size(Button.Location.X + Button.Width - Text.Location.X, 100)
                Panel.Location = New Point(Text.Location.X, 75)
                Panel.View = View.List
                Me.Controls.Add(Panel)
                AddHandler Panel.ItemSelectionChanged, AddressOf ListViewEnableButton

                Dim SubMessage, IDString As String
                Dim StartIndex As Integer = 0
                Dim DelimiterStart, DelimiterEnd As Integer
                DelimiterEnd = ServerResponse.Length
                IDString = Nothing

                While DelimiterStart >= StartIndex
                    DelimiterStart = ServerResponse.IndexOf("""Id"":", StartIndex) + 4
                    SubMessage = ServerMessage(ServerResponse, DelimiterStart, DelimiterEnd)
                    If StartIndex < DelimiterStart Then
                        StartIndex = DelimiterStart
                        IDString = (SubMessage.Substring(1, SubMessage.IndexOf(",") - 1))
                    End If
                    DelimiterStart = ServerResponse.IndexOf("Title", StartIndex) + 8
                    SubMessage = ServerMessage(ServerResponse, DelimiterStart, DelimiterEnd)
                    If StartIndex < DelimiterStart Then
                        StartIndex = DelimiterStart
                        Dim listItem As New ListViewItem
                        listItem.Text = ServerMessage(SubMessage, 0, SubMessage.IndexOf(""""))
                        listItem.Tag = IDString
                        Panel.Items.Add(listItem)
                    End If
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            LoginGUI()
        End Try
    End Sub

    Private Sub ListViewEnableButton()
        Dim List As ListView = Controls("Panel")
        Dim Open As Button = Controls("Open")
        Dim Item As ListViewItem
        Try
            For Each Item In List.Items
                If Item.Selected Then
                    Open.Enabled = True
                    Exit Sub
                End If
            Next
            Open.Enabled = False
        Catch ex As Exception
        End Try
    End Sub

    Private Sub openBoard()
        Dim List As ListView = Controls("Panel")
        Dim Item As ListViewItem
        Dim Board As New Leankit.Board

        For Each Item In List.Items
            If Item.Selected Then
                Leankit.Board.Id = Item.Tag
                Try
                    Board.getIdentifiers()
                    Dim Manager As New Manager
                    Manager.Show()
                    Manager.CurrentForm = Manager
                    Hide()
                    Dispose()
                Catch ex As Exception
                    MsgBox("There was an error while attempting to load the board. This user may not have permsision to use this software.")
                    BoardPopulate()
                End Try
                Exit Sub
            End If
        Next
    End Sub

    Private Sub EnterClick(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then BoardPopulate()
    End Sub

    Private Sub EnterCheck(ByVal sender As System.Object, ByVal e As EventArgs)
        Call EnterClick(sender, e)
    End Sub

    Private Sub Form_Closing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Application.Exit()
        End If
    End Sub

    Private Sub AutoFix()
        If Me.Controls("Login").Controls("Username").Text.Contains("@rochestermed.com") <> True Then Me.Controls("Login").Controls("Username").Text = Me.Controls("Login").Controls("Username").Text & "@rochestermed.com"
        If Me.Controls("Login").Controls("Account").Text.Contains(".leankit.com") Then Me.Controls("Login").Controls("Account").Text = Me.Controls("Login").Controls("Account").Text.Replace(".leankit.com", "")
    End Sub

    Private Sub Public_Vars()
        Leankit.Account.Name = Me.Controls("Login").Controls("Account").Text
        Leankit.Account.Credentials = New NetworkCredential(Me.Controls("Login").Controls("Username").Text, Me.Controls("Login").Controls("Password").Text)
        Leankit.Account.Email = Me.Controls("Login").Controls("Username").Text
    End Sub

    Private Function ServerMessage(ByVal Message As String, ByVal DelimiterStart As Integer, ByVal DelimiterEnd As Integer)
        Return Message.Substring(DelimiterStart, DelimiterEnd - DelimiterStart)
    End Function

End Class