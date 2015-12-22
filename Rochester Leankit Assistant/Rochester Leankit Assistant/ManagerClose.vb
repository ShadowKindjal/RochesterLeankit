Imports System.Environment

Public Class ManagerClose

    Private Sub ManagerClose_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Controls.Clear()
        Me.Text = "Rochester Leankit Assistant"
        Me.Size = New Size(500, 300)
        Me.Top = ((Manager.CurrentForm.Height - Me.Height) / 2) + Manager.CurrentForm.Top
        Me.Left = ((Manager.CurrentForm.Width - Me.Width) / 2) + Manager.CurrentForm.Left
        Me.BackColor = Color.White
        Me.FormBorderStyle = FormBorderStyle.None

        Dim Panel As New Panel
        Panel.Name = "Panel"
        Panel.Width = 400
        Panel.Height = 100
        Panel.Location = New Point((Me.Width - Panel.Width) / 2, 0)
        Panel.BackColor = Color.White
        Me.Controls.Add(Panel)

        Dim Picture As New PictureBox
        Picture.Name = "Picture"
        Picture.Size = New Size(150, 50)
        Picture.Location = New Point(0, 25)
        Picture.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\Leankit.png"
        Picture.BorderStyle = BorderStyle.None
        Picture.SizeMode = PictureBoxSizeMode.StretchImage
        Panel.Controls.Add(Picture)
        Picture.Load()

        Dim Logo As New PictureBox
        Logo.Name = "Logo"
        Logo.Size = New Size(195, 70)
        Logo.Location = New Point(205, 15)
        Logo.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\Rochester.png"
        Logo.BorderStyle = BorderStyle.None
        Logo.SizeMode = PictureBoxSizeMode.StretchImage
        Panel.Controls.Add(Logo)
        Logo.Load()

        Dim CloseApp As New Button
        CloseApp.Text = "Close Application"
        CloseApp.Size = New Size(400, 25)
        CloseApp.Top = Me.ClientSize.Height - 50
        CloseApp.Left = (Me.Width - CloseApp.Width) / 2
        CloseApp.FlatStyle = FlatStyle.Flat
        CloseApp.BackColor = Color.FromArgb(150, 201, 61)
        CloseApp.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        AddHandler CloseApp.Click, AddressOf AppExit
        Me.Controls.Add(CloseApp)

        Dim UserLogout As New Button
        UserLogout.Text = "Logout " & Leankit.Account.FullName
        UserLogout.Size = New Size(400, 25)
        UserLogout.Top = CloseApp.Location.Y - 37.5
        UserLogout.Left = (Me.Width - CloseApp.Width) / 2
        UserLogout.FlatStyle = FlatStyle.Flat
        UserLogout.BackColor = Color.FromArgb(150, 201, 61)
        UserLogout.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        AddHandler UserLogout.Click, AddressOf Logout
        Me.Controls.Add(UserLogout)

        Dim sBoard As New Button
        sBoard.Text = "Board Selection"
        sBoard.Size = New Size(400, 25)
        sBoard.Top = UserLogout.Location.Y - 37.5
        sBoard.Left = (Me.Width - CloseApp.Width) / 2
        sBoard.FlatStyle = FlatStyle.Flat
        sBoard.BackColor = Color.FromArgb(150, 201, 61)
        sBoard.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        AddHandler sBoard.Click, AddressOf selectBoard
        Me.Controls.Add(sBoard)

        Dim cReturn As New Button
        cReturn.Text = "Back to Current Board"
        cReturn.Size = New Size(400, 25)
        cReturn.Top = sBoard.Location.Y - 37.5
        cReturn.Left = (Me.Width - CloseApp.Width) / 2
        cReturn.FlatStyle = FlatStyle.Flat
        cReturn.BackColor = Color.FromArgb(150, 201, 61)
        cReturn.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        AddHandler cReturn.Click, AddressOf BoardReturn
        Me.Controls.Add(cReturn)
    End Sub

    Private Sub AppExit()
        Application.Exit()
    End Sub

    Sub Logout()
        Dim ManagerForm As Form = Manager
        Login.LoggedIn = False
        Me.Hide()
        Me.Dispose()
        Manager.CurrentForm.Hide()
        Login.Show()
    End Sub

    Sub selectBoard()
        Dim ManagerForm As Form = Manager
        Login.LoggedIn = True
        Me.Hide()
        Me.Dispose()
        Manager.CurrentForm.Hide()
        Login.Show()
    End Sub

    Sub BoardReturn()
        Me.Hide()
    End Sub
End Class