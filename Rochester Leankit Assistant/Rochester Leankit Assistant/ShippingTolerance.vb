Imports System.Environment

Public Class ShippingTolerance
    Private Sub ShippingTolerance_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Controls.Clear()
        Me.Text = "Rochester Leankit Assistant"
        Me.Size = New Size(500, 400)
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

        Dim calcRun As New Button
        calcRun.Text = "Calculate Run"
        calcRun.Size = New Size(400, 25)
        calcRun.Top = Me.ClientSize.Height - 50
        calcRun.Left = (Me.Width - calcRun.Width) / 2
        calcRun.FlatStyle = FlatStyle.Flat
        calcRun.BackColor = Color.FromArgb(150, 201, 61)
        calcRun.FlatAppearance.BorderColor = Color.FromArgb(150, 201, 61)
        calcRun.DialogResult = DialogResult.OK
        AddHandler calcRun.Click, AddressOf UpdateRun

        Dim lLowTolerance As New Label
        lLowTolerance.Text = "Low Weight Tolerance:"
        Dim LowTolerance As New NumericUpDown
        LowTolerance.Minimum = 0
        LowTolerance.Maximum = 5
        LowTolerance.DecimalPlaces = 2
        LowTolerance.Increment = 0.1
        LowTolerance.Value = Run.LowTolerance
        LowTolerance.Name = "LowTol"

        Dim lHighTolerance As New Label
        lHighTolerance.Text = "High Weight Tolerance:"
        Dim Hightolerance As New NumericUpDown
        Hightolerance.Minimum = 0
        Hightolerance.Maximum = 5
        Hightolerance.DecimalPlaces = 2
        Hightolerance.Increment = 0.1
        Hightolerance.Value = Run.HighTolerance
        Hightolerance.Name = "HighTol"

        Dim lMinWeight As New Label
        lMinWeight.Text = "Minimum Box Weight:"
        Dim MinWeight As New NumericUpDown
        MinWeight.Minimum = 0
        MinWeight.Maximum = 100
        MinWeight.DecimalPlaces = 2
        MinWeight.Increment = 0.1
        MinWeight.Value = Run.MinWeight
        MinWeight.Name = "MinWt"

        Dim lMaxWeight As New Label
        lMaxWeight.Text = "Maximum Box Weight:"
        Dim MaxWeight As New NumericUpDown
        MaxWeight.Minimum = 0
        MaxWeight.Maximum = 100
        MaxWeight.DecimalPlaces = 2
        MaxWeight.Increment = 0.1
        MaxWeight.Value = Run.MaxWeight
        MaxWeight.Name = "MaxWt"

        Dim lBoxCount As New Label
        lBoxCount.Text = "Box Count:"
        Dim BoxCount As New NumericUpDown
        BoxCount.Minimum = 10
        BoxCount.Maximum = 100
        BoxCount.DecimalPlaces = 0
        BoxCount.Increment = 10
        BoxCount.Value = Run.BoxCount
        BoxCount.Name = "BoxCt"
        AddHandler BoxCount.ValueChanged, AddressOf calculateAvgWeight

        Dim Weight As New Label
        Weight.Text = "Total Weight: " & Run.Weight & " lb"
        Dim AvgWeight As New Label
        AvgWeight.Name = "AvgWeight"
        AvgWeight.Text = "Average Weight: " & Math.Round(Run.AverageWeight, 2) & " lb"

        lLowTolerance.Width = calcRun.Width / 2
        lHighTolerance.Width = calcRun.Width / 2
        lMinWeight.Width = calcRun.Width / 2
        lMaxWeight.Width = calcRun.Width / 2
        lBoxCount.Width = calcRun.Width / 2
        lLowTolerance.Left = calcRun.Left
        lHighTolerance.Left = calcRun.Left
        lMinWeight.Left = calcRun.Left
        lMaxWeight.Left = calcRun.Left
        Weight.Left = calcRun.Left
        lBoxCount.Left = calcRun.Left
        LowTolerance.Width = calcRun.Width / 2
        Hightolerance.Width = calcRun.Width / 2
        MinWeight.Width = calcRun.Width / 2
        MaxWeight.Width = calcRun.Width / 2
        Weight.Width = calcRun.Width / 2
        AvgWeight.Width = calcRun.Width / 2
        BoxCount.Width = calcRun.Width / 2
        LowTolerance.Left = calcRun.Left + calcRun.Width / 2
        Hightolerance.Left = calcRun.Left + calcRun.Width / 2
        MinWeight.Left = calcRun.Left + calcRun.Width / 2
        MaxWeight.Left = calcRun.Left + calcRun.Width / 2
        AvgWeight.Left = calcRun.Left + calcRun.Width / 2
        BoxCount.Left = calcRun.Left + calcRun.Width / 2
        lLowTolerance.Top = calcRun.Location.Y - lLowTolerance.Height - 10
        LowTolerance.Top = lLowTolerance.Top
        lHighTolerance.Top = lLowTolerance.Location.Y - lHighTolerance.Height - 10
        Hightolerance.Top = lHighTolerance.Top
        lMinWeight.Top = lHighTolerance.Location.Y - lMinWeight.Height - 10
        MinWeight.Top = lMinWeight.Top
        lMaxWeight.Top = lMinWeight.Location.Y - lMaxWeight.Height - 10
        MaxWeight.Top = lMaxWeight.Top
        lBoxCount.Top = lMaxWeight.Location.Y - lBoxCount.Height - 10
        BoxCount.Top = lBoxCount.Top
        Weight.Top = BoxCount.Location.Y - BoxCount.Height - 10
        AvgWeight.Top = Weight.Top

        Me.Controls.Add(calcRun)
        Me.Controls.Add(lLowTolerance)
        Me.Controls.Add(LowTolerance)
        Me.Controls.Add(lHighTolerance)
        Me.Controls.Add(Hightolerance)
        Me.Controls.Add(lMinWeight)
        Me.Controls.Add(MinWeight)
        Me.Controls.Add(lMaxWeight)
        Me.Controls.Add(MaxWeight)
        Me.Controls.Add(lBoxCount)
        Me.Controls.Add(BoxCount)
        Me.Controls.Add(Weight)
        Me.Controls.Add(AvgWeight)

    End Sub

    Private Sub calculateAvgWeight()
        Dim BoxCount As NumericUpDown = Me.Controls("BoxCt")
        If BoxCount.Value <> 0 Then
            Run.BoxCount = BoxCount.Value
            Run.AverageWeight = Run.Weight / Run.BoxCount
        End If
        Me.Controls("AvgWeight").Text = "Average Weight: " & Math.Round(Run.AverageWeight, 2) & " lb"
    End Sub

    Sub UpdateRun()
        Run.LowTolerance = Me.Controls("LowTol").Text
        Run.HighTolerance = Me.Controls("HighTol").Text
        Run.MinWeight = Me.Controls("MinWt").Text
        Run.MaxWeight = Me.Controls("MaxWt").Text
        Run.BoxCount = Me.Controls("BoxCt").Text
        If Run.AverageWeight < Run.MinWeight Then MsgBox("The avergae weight of this run is below the minimum box weight!", MsgBoxStyle.Critical, "Warning")
    End Sub
End Class
