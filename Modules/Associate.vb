Imports Microsoft.Win32.Registry
Module Associate
    Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Int32, ByVal uFlags As Int32, ByVal dwItem1 As Int32, ByRef dwItem2 As Int32)
    Private Const SHCNE_ASSOCCHANGED As Int32 = &H8000000I
    Private Const SHCNF_IDLIST As Int32 = &H0&

    Public Sub AssociateMerge()
        Try
            Dim _KeyName As String
            Dim _KeyValue As String
            _KeyName = PROGNAME & ".Merge"
            _KeyValue = PROGNAME & "'s splitted file"
            ClassesRoot.CreateSubKey(_KeyName).SetValue(vbNullString, _KeyValue, Microsoft.Win32.RegistryValueKind.String)
            _KeyName = ".001"
            _KeyValue = PROGNAME & ".Merge"
            ClassesRoot.CreateSubKey(_KeyName).SetValue("", _KeyValue, Microsoft.Win32.RegistryValueKind.String)
            _KeyName = PROGNAME & ".Merge"
            _KeyValue = Application.ExecutablePath & " /m %1"
            ClassesRoot.CreateSubKey(_KeyName).CreateSubKey("shell\open\command").SetValue("", _KeyValue, Microsoft.Win32.RegistryValueKind.String)
            SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0)
        Catch ex As Exception
        End Try
    End Sub

    Public Sub AssociateHandler()
        Try
            With My.Application.CommandLineArgs
                If .Count < 2 Then Exit Sub
                Select Case .Item(0)
                    Case "/m"
                        Form1.TabControl1.SelectedIndex = 1
                        Dim _FileName As String = .Item(1)
                        _FileName = IO.Path.GetFullPath(.Item(1))
                        If Not IO.File.Exists(_FileName) Then Exit Sub
                        Dim _FileNames As String() = Form1.getMergeFiles(GetMergeFileName(_FileName))
                        If IsNothing(_FileNames) Then Exit Sub
                        For Each _mFileName As String In _FileNames
                            Form1.AddItemToList(_mFileName)
                        Next
                        Form1.CheckOptions()
                    Case Else
                End Select
            End With
        Catch ex As Exception
        End Try
    End Sub
End Module
