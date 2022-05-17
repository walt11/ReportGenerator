Public Class MainForm

    Private CM As New ConfigManager("Config.cfg")

    Private Sub SettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem.Click
        Dim settingsForm As New SettingsForm(CM)
        settingsForm.ShowDialog()
    End Sub

    Private Sub OpenPDFToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenPDFToolStripMenuItem.Click
        Dim ofd As New OpenFileDialog
        If ofd.ShowDialog() = DialogResult.OK Then
            AxAcroPDF1.LoadFile(ofd.FileName)
            AxAcroPDF1.setView("Fit")
        End If
    End Sub
End Class
