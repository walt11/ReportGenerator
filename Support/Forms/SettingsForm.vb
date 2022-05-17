Public Class SettingsForm

    Private CM As ConfigManager

    Public Sub New(ByRef CM As ConfigManager)

        ' This call is required by the designer.
        InitializeComponent()

        Me.CM = CM

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub SaveButton_Click(sender As Object, e As EventArgs) Handles SaveButton.Click
        CM.Save()
        Me.Close()
    End Sub
End Class