Imports System.Environment
Imports System.Threading

Public Class Load
    Private Resoucres As New MediaFire

    Private Sub FormLoad(sender As Object, e As EventArgs) Handles MyBase.Load
        Resoucres.getFile("ysm7h70z9ir5ll2", "Leankit", False, ".png")
        Resoucres.getFile("lxo9gq2ty0on7cb", "Rochester", False, ".png")
        Resoucres.getFile("hmd8rj4iq818hm1", "Load", False, ".gif")
        GUISetup()
    End Sub

    Private Async Sub FormShown(sender As Object, e As EventArgs) Handles Me.Shown
        Refresh()
        'Dim Download = New Thread(AddressOf resourceDownload)
        'Download.Start()
        Await (Task.Run(Sub() resourceDownload()))
        Login.Show()
        Hide()
    End Sub

    Private Sub resourceDownload()
        Resoucres.getFile("b664di6zs3bkhbl", "Leankit-White", False, ".png")
        Resoucres.getFile("07bijl970315i2k", "Triangle", False, ".png")
        Resoucres.getFile("q6n20gw2i0yvxxu", "Logout", False, ".png")
        Resoucres.getFile("711t3wkm1atv4c1", "PlusSign", False, ".png")
        Resoucres.getFile("j6djsawf8ni76m3", "BlackX", False, ".png")
        Resoucres.getFile("b7lq5b8vvigk8nz", "Export", False, ".png")
        Resoucres.getFile("w5veb6a7bbwjwza", "User", False, ".png")
        Resoucres.getFile("bi38hh8w33af0ie", "refresh", False, ".png")
        Resoucres.getFile("edk4ty8jgi6f2r4", "Product_Library", True, ".xml")
        Resoucres.getFile("jjy6g479dd2p5ub", "Version", False, ".xml")
        Resoucres.getFile("89nb12db7z75q17", "ShippingWeights", True, ".xlsx")
    End Sub

    Private Sub GUISetup()
        Me.Text = "Rochester Leankit Assistant"
        Me.Size = New Size(500, 150)  'Sets the size of the loading screen GUI
        Me.Top = (Screen.PrimaryScreen.Bounds.Height - Me.Height) / 2
        Me.Left = (Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2
        Me.BackColor = Color.White
        Me.FormBorderStyle = FormBorderStyle.None

        Dim Panel As New Panel
        Panel.Name = "Panel"
        Panel.Width = 400
        Panel.Height = 100
        Panel.Location = New Point((Me.Width - Panel.Width) / 2, (Me.Height - Panel.Height) / 2)
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

        Dim Load As New PictureBox
        Load.Name = "Load"
        Load.Size = New Size(25, 25)
        Load.Location = New Point(Me.Width - 30, Me.Height - 30)
        Load.ImageLocation = GetFolderPath(SpecialFolder.ApplicationData) & "\Rochester\Load.gif"
        Load.BorderStyle = BorderStyle.None
        Load.SizeMode = PictureBoxSizeMode.StretchImage
        Controls.Add(Load)
        Load.Load()
    End Sub

End Class