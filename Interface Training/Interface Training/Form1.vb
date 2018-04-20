Public Class Form1

    Dim Gen As Generate = New Generate()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Gen.LoadCsv()
    End Sub
End Class
