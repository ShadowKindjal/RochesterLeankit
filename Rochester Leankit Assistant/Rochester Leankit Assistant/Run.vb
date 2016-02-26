Public Class Run
    Public Shared BoxWeight As Double = 0
    Public Shared BoxCount As Double = 20
    Public Shared AverageWeight As Double = 28.6
    Public Shared MinWeight As Double = 28
    Public Shared MaxWeight As Double = 32
    Public Shared LowTolerance As Double = 0.85
    Public Shared HighTolerance As Double = 0.5
    Public Shared Weight As Double = 0
    Public Shared WeightTrigger As Double = 680

    Public Shared Sub Reset()
        BoxWeight = 0
        BoxCount = 20
        AverageWeight = 28.6
        MinWeight = 28
        MaxWeight = 32
        Weight = 0
        WeightTrigger = 680
    End Sub

    Public Class Part
        Public PartNumber As String
        Public Lot As String
        Public Quantity As Integer
        Public WireType As String
        Public NeedlesPerSet As Integer
        Public SetsPerBag As Integer
        Public BagWeight As Double
        Public TotalNeedles As Integer
        Public TotalBags As Integer
        Public TotalWeight As Double

        Public Sub Calculate()
            Try
                TotalNeedles = Quantity * NeedlesPerSet
                TotalBags = Quantity / SetsPerBag
                TotalWeight = TotalBags * BagWeight
            Catch ex As Exception
                TotalNeedles = 0
                TotalBags = 0
                TotalWeight = 0
            End Try
        End Sub
    End Class

    Public Class Box
        Public Weight As Double
        Public Parts As New List(Of Part)

        Public Class Part
            Public PartNumber As String
            Public Lot As String
            Public Bags As Integer
            Public BagWeight As Double
        End Class
    End Class
End Class