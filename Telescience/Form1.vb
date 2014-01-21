Imports System.Math
Public Class Form1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Distance As Double
        Dim POFF As Double
        Dim BOFF As Double
        Dim X As Double
        Dim Y As Double
        X = Abs(TextBox1.Text - TextBox3.Text)
        Y = Abs(TextBox2.Text - TextBox4.Text)
        Distance = Sqrt(X ^ 2 + Y ^ 2)
        POFF = 20 - Sqrt(Distance * 10)
        Label7.Text = Math.Round(POFF)

        BOFF = Atan(X / Y) * 57.3

        If TextBox3.Text < TextBox1.Text Then
            Label8.Text = Math.Round(BOFF * -1)
        Else
            Label8.Text = Math.Round(BOFF)
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.Enabled = False
            TextBox2.Enabled = False
        Else
            TextBox1.Enabled = True
            TextBox2.Enabled = True
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox3.Enabled = False
            TextBox4.Enabled = False
        Else
            TextBox3.Enabled = True
            TextBox4.Enabled = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Elevation As Double
        Dim Bearing As Double
        Dim D As Double
        Dim Dmax As Double
        Dim X As Double
        Dim Y As Double

        X = Abs(TextBox1.Text - TextBox5.Text)
        Y = Abs(TextBox2.Text - TextBox6.Text)

        D = Sqrt((X ^ 2) + (Y ^ 2))
        Dmax = ((ComboBox1.SelectedItem - Label7.Text) ^ 2) / 10

        Elevation = (Asin(D / Dmax) * 57.3) / 2
        Label14.Text = Math.Round(Elevation)

        If ComboBox2.SelectedItem = "NE" Then
            Bearing = (Atan(X / Y) * 57.3) + 0 - Label8.Text
        ElseIf ComboBox2.SelectedItem = "SE" Then
            Bearing = (Atan(Y / X) * 57.3) + 90 - Label8.Text
        ElseIf ComboBox2.SelectedItem = "SW" Then
            Bearing = (Atan(X / Y) * 57.3) + 180 - Label8.Text
        ElseIf ComboBox2.SelectedItem = "NW" Then
            Bearing = (Atan(Y / X) * 57.3) + 270 - Label8.Text
        End If

        Label15.Text = Math.Round(Bearing, 2)
    End Sub
End Class
