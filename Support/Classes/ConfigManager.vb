Imports System.IO
Public Class ConfigManager

    Private DSoptions As DataSet
    Private mConfigFileName As String

    Public ReadOnly Property ConfigFileName() As String
        Get
            Return mConfigFileName
        End Get
    End Property

    Public Sub New(ByVal ConfigFile As String)
        mConfigFileName = ConfigFile
        DSoptions = New DataSet("ConfigOpt")
        If File.Exists(ConfigFile) Then
            DSoptions.ReadXml(ConfigFile)
        Else
            Dim dt As New DataTable("ConfigValues")
            dt.Columns.Add("OptionName", System.Type.GetType("System.String"))
            dt.Columns.Add("OptionValue", System.Type.GetType("System.String"))
            DSoptions.Tables.Add(dt)
        End If
    End Sub

    Public Sub Save()
        Save(mConfigFileName)
    End Sub

    Public Sub Save(ByVal ConfigFile As String)
        mConfigFileName = ConfigFile
        DSoptions.WriteXml(ConfigFile)
    End Sub

    Public Function GetProperty(ByVal OptionName As String) As String
        Dim dv As DataView = DSoptions.Tables("ConfigValues").DefaultView
        dv.RowFilter = "OptionName='" & OptionName & "'"
        If dv.Count > 0 Then
            Return CStr(dv.Item(0).Item("OptionValue"))
        Else
            Return ""
        End If
    End Function

    Public Sub SetProperty(ByVal OptionName _
             As String, ByVal OptionValue As String)
        Dim dv As DataView = DSoptions.Tables("ConfigValues").DefaultView
        dv.RowFilter = "OptionName='" & OptionName & "'"
        If dv.Count > 0 Then
            dv.Item(0).Item("OptionValue") = OptionValue
        Else
            Dim dr As DataRow = DSoptions.Tables("ConfigValues").NewRow()
            dr("OptionName") = OptionName
            dr("OptionValue") = OptionValue
            DSoptions.Tables("ConfigValues").Rows.Add(dr)
        End If
    End Sub
End Class
