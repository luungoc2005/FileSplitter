Module SaveLoadSettings
    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, _
ByVal lpKeyName As String, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Int32, _
ByVal lpFileName As String) As Int32
    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32.dll" _
    Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, _
                                        ByVal lpKeyName As String, ByVal lpString As String, _
                                        ByVal lpFileName As String) As Int32

    Public Sub SaveAppSettings()
        With Form1
            SetSettingString(SettingFile, "SETTI_CHK1", CInt(.CheckBox6.Checked))
            SetSettingString(SettingFile, "SETTI_NBOX1", .NumericUpDown3.Value)
            SetSettingString(SettingFile, "SETTI_CBOX1", .ComboBox2.SelectedIndex)
            SetSettingString(SettingFile, "SETTI_CBOX2", .ComboBox3.SelectedIndex)
            SetSettingString(SettingFile, "POS_TOP", .Location.Y)
            SetSettingString(SettingFile, "POS_LEFT", .Location.X)
        End With
    End Sub

    Public Sub LoadAppSettings()
        With Form1
            .CheckBox6.Checked = CBool(GetSettingString(SettingFile, "SETTI_CHK1", True))
            .NumericUpDown3.Value = CInt(GetSettingString(SettingFile, "SETTI_NBOX1", 20))
            .ComboBox2.SelectedIndex = CInt(GetSettingString(SettingFile, "SETTI_CBOX1", 2))
            If .ComboBox3.Items.Count > 0 Then .ComboBox3.SelectedIndex = CInt(GetSettingString(SettingFile, "SETTI_CBOX2", .ComboBox3.Items.Count - 1))
            .Location = New Point(CInt(GetSettingString(SettingFile, "POS_LEFT", .Location.X)), CInt(GetSettingString(SettingFile, "POS_TOP", .Location.Y)))
        End With
    End Sub

    Private Function SettingFile() As String
        Return GetSettingFile("config.ini")
    End Function

    Private Function GetSettingFile(ByVal FileName As String) As String
        Return Application.StartupPath & Slash(Application.StartupPath) & FileName
    End Function

    Private Function GetSettingString(ByVal FileName As String, ByVal KeyName As String, ByVal sDefault As String) As String
        Dim n As Int32
        Dim sData As String
        sData = Space$(1024)
        n = GetPrivateProfileString(UCase(PROGNAME), KeyName, sDefault, sData, sData.Length, FileName)
        If n > 0 Then
            Return sData.Substring(0, n)
        Else
            Return vbNullString
        End If
    End Function

    Private Sub SetSettingString(ByVal FileName As String, ByVal KeyName As String, ByVal Value As String)
        WritePrivateProfileString(UCase(PROGNAME), KeyName, Value, FileName)
    End Sub
End Module
