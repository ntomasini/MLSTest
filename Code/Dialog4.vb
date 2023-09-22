Imports System.ComponentModel
Imports System.Windows.Forms

Public Class Dialog4

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK


        For i = 0 To Form1.GridView1.Rows.Count - 1
            Dim info As New ProcessStartInfo()
            info.FileName = "PowerShell.exe"
            info.WorkingDirectory = System.IO.Path.GetDirectoryName(Form1.GridView1.Item(1, i).Value)
            Dim a As String = "./fasttree.exe " & TextBox3.Text & " " & Form1.GridView1.Item(0, i).Value & ".fasta > " & Form1.GridView1.Item(0, i).Value & ".tree"
            info.Arguments = a
            'info.UseShellExecute = False
            info.WindowStyle = ProcessWindowStyle.Hidden
            Process.Start(info)

        Next




        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub


End Class
