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
End Class