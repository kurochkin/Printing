Public Class Form1
    Dim xlApp

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        xlApp = CreateObject("SerialPortDataProcessor.PrintLabelProcessor")
        xlApp.Init()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        xlApp = CreateObject("SerialPortDataProcessor.PrintLabelProcessor")
        xlApp.Init()
        xlApp.PrintLabel(100)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        xlApp.StopReading()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim res
        res = xlApp.ReadedData()
        TextBox1.Text = res
    End Sub
End Class
