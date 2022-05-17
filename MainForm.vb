Public Class MainForm

    Private CM As New ConfigManager("Config.cfg")

    Private Sub SettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem.Click
        Dim settingsForm As New SettingsForm(CM)
        settingsForm.ShowDialog()
    End Sub
End Class
