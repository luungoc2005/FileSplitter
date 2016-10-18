Imports System.Security
Imports System.IO

Public Class Form1
    Private IsSplitValid As Boolean
    Private IsSplitPrg As Boolean
    Private IsMergePrg As Boolean
    Private StopSlit As Boolean
    Private StopMerge As Boolean
    Private LangFiles() As String = New String() {}
    Private SplitStatus As Integer '0: Idle, 1:Busy, 2:Succ.Complete, 3:Complete w Errors
    Private MergeStatus As Integer 'Same as above
    Private EncryptStatus As Integer '...
    Private ClosePending As Boolean
    Private _margins As MARGINS

#Region "Radio Buttons"
    Public Sub CheckOptions()
        NumericUpDown1.Enabled = RadioButton1.Checked
        ComboBox1.Enabled = RadioButton1.Checked
        NumericUpDown2.Enabled = Not RadioButton1.Checked
        Label4.Enabled = Not RadioButton1.Checked
        CheckBox3.Enabled = CheckBox1.Checked
        ComboBox2.Enabled = CheckBox6.Checked
        NumericUpDown3.Enabled = CheckBox6.Checked
        Label16.Enabled = CheckBox6.Checked
        TextBox7.Enabled = RadioButton3.Checked
        Label23.Enabled = RadioButton3.Checked
        CheckValidInfo()
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        CheckOptions()
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        CheckOptions()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        CheckOptions()
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If SplitStatus = 1 Then
            e.Cancel = True
            Call Button10_Click(Me, New System.EventArgs)
            ClosePending = True
        End If
        If MergeStatus = 1 Then
            e.Cancel = True
            Call Button11_Click(Me, New System.EventArgs)
            ClosePending = True
        End If
        SaveAppSettings()
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ComboBox1.SelectedIndex = 1
        ComboBox2.SelectedIndex = 1
        NumericUpDown3.Value = 10
        Label8.Text = vbNullString
        Label10.Text = vbNullString
        Label24.Text = vbNullString
        IsSplitValid = False
        CheckOptions()
        CheckLangFiles()
        ChangeLang("Vietnamese.lang")
        CheckValidInfo()
        LoadAppSettings()
        AssociateHandler()

        'Dim _Button As New Splitter_Button
        '_Button.Visible = True
        '_Button.Enabled = True
        'TabPage1.Controls.Add(_Button)
    End Sub

    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        MyBase.OnLoad(e)
        If IsGlassEnabled() Then
            _margins = New MARGINS
            _margins.cyTopHeight = 22
            _margins.cxLeftWidth = 3
            _margins.cxRightWidth = 3
            _margins.cyBottomheight = 3
            DwmExtendFrameIntoClientArea(Me.Handle, _margins)
        End If
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaintBackground(e)
        If IsGlassEnabled() Then
            e.Graphics.Clear(Color.Black)
            Dim _ClientArea As New Rectangle( _
            _margins.cxLeftWidth, _margins.cyTopHeight, _
            Me.ClientRectangle.Width - _margins.cxLeftWidth - _margins.cxRightWidth, _
            Me.ClientRectangle.Height - _margins.cyTopHeight - _margins.cyBottomheight)

            Dim _Brush As Brush = New SolidBrush(Me.BackColor)
            e.Graphics.FillRectangle(_Brush, _ClientArea)
        End If
    End Sub

    Private Sub CheckLangFiles()
        Dim _tempFiles As String() = Directory.GetFiles(Application.StartupPath)
        For Each _FileName As String In _tempFiles
            If _FileName.Substring(_FileName.Length - 5, 5).ToLower = ".lang" Then
                ComboBox3.Items.Add(GetInfoString(_FileName, "FILE_LANG_NAME", "#N/A"))
                ReDim Preserve LangFiles(LangFiles.Length)
                LangFiles(LangFiles.Length - 1) = _FileName
            End If
        Next
    End Sub

#End Region

#Region "Update Progress"
    'Split
    Public Sub ReceiveSplitProgress(ByVal ThreadIndex As Integer, ByVal Text1 As String, ByVal Text2 As String)
        Label7.Text = Text1
        Label8.Text = Text2
    End Sub
    Public Sub UpdateSplitPrgBar(ByVal BytesProcessed As Long, ByVal BytesTotal As Long)
        Dim _value As Integer : _value = Math.Round(BytesProcessed / BytesTotal * 100)
        If _value <= ProgressBar1.Maximum Then ProgressBar1.Value = _value
    End Sub
    Public Sub SplitComplete(ByVal ThreadIndex As Integer, ByVal IsError As Boolean)
        SplitStatus = IIf(IsError, 3, 2)
    End Sub

    'Merge
    Public Sub ReceiveMergeProgress(ByVal ThreadIndex As Integer, ByVal Text1 As String, ByVal Text2 As String)
        Label11.Text = Text1
        Label10.Text = Text2
    End Sub
    Public Sub UpdateMergePrgBar(ByVal BytesProcessed As Long, ByVal BytesTotal As Long)
        Dim _value As Integer : _value = Math.Round(BytesProcessed / BytesTotal * 100)
        If _value <= ProgressBar2.Maximum Then ProgressBar2.Value = _value
    End Sub
    Public Sub MergeComplete(ByVal ThreadIndex As Integer, ByVal IsError As Boolean)
        MergeStatus = IIf(IsError, 3, 2)
    End Sub

    'Encrypt
    Public Sub ReceiveEncryptProgress(ByVal ThreadIndex As Integer, ByVal Text1 As String, ByVal Text2 As String)
        Label25.Text = Text1
        Label24.Text = Text2
    End Sub
    Public Sub UpdateEncPrgBar(ByVal BytesProcessed As Long, ByVal BytesTotal As Long)
        Dim _value As Integer : _value = Math.Round(BytesProcessed / BytesTotal * 100)
        If _value <= ProgressBar3.Maximum Then ProgressBar3.Value = _value
    End Sub
    Public Sub EncComplete(ByVal ThreadIndex As Integer, ByVal IsError As Boolean)
        EncryptStatus = IIf(IsError, 3, 2)
    End Sub
#End Region

#Region "Split/Merge"
    Private Sub MoveListViewItem(ByRef lv As ListView, ByVal moveUp As Boolean)
        Dim i As Integer
        Dim cache As String
        Dim selIdx As Integer

        With lv
            selIdx = .SelectedItems.Item(0).Index
            If moveUp Then
                ' ignore moveup of row(0)
                If selIdx = 0 Then
                    Exit Sub
                End If
                ' move the subitems for the previous row
                ' to cache so we can move the selected row up
                For i = 0 To .Items(selIdx).SubItems.Count - 1
                    cache = .Items(selIdx - 1).SubItems(i).Text
                    .Items(selIdx - 1).SubItems(i).Text = .Items(selIdx).SubItems(i).Text
                    .Items(selIdx).SubItems(i).Text = cache
                Next
                ' tooltiptext
                cache = .Items(selIdx - 1).ToolTipText
                .Items(selIdx - 1).ToolTipText = .Items(selIdx).ToolTipText
                .Items(selIdx).ToolTipText = cache

                .Items(selIdx - 1).Selected = True
                .Refresh()
                .Focus()
            Else
                ' ignore move down of last row
                If selIdx = .Items.Count - 1 Then
                    Exit Sub
                End If
                ' move the subitems for the next row
                ' to cache so we can move the selected row down
                For i = 0 To .Items(selIdx).SubItems.Count - 1
                    cache = .Items(selIdx + 1).SubItems(i).Text
                    .Items(selIdx + 1).SubItems(i).Text = .Items(selIdx).SubItems(i).Text
                    .Items(selIdx).SubItems(i).Text = cache
                Next
                ' tooltiptext
                cache = .Items(selIdx + 1).ToolTipText
                .Items(selIdx + 1).ToolTipText = .Items(selIdx).ToolTipText
                .Items(selIdx).ToolTipText = cache

                .Items(selIdx + 1).Selected = True
                .Refresh()
                .Focus()
            End If
        End With
    End Sub

    Private Function CalcPartSize() As Long
        If File.Exists(TextBox1.Text) = False Then Exit Function
        If RadioButton1.Checked = True Then
            Select Case ComboBox1.SelectedIndex
                Case 0
                    Return NumericUpDown1.Value
                Case 1
                    Return NumericUpDown1.Value * 1024
                Case 2
                    Return NumericUpDown1.Value * 1048576
            End Select
        Else
            Return Math.Ceiling(FileLen(TextBox1.Text) / NumericUpDown2.Value)
        End If
    End Function

    Private Function CalcBufferSize() As Long
        If CheckBox6.Checked = False Then Return CalcPartSize() + 1
        Select Case ComboBox2.SelectedIndex
            Case 0
                Return NumericUpDown3.Value
            Case 1
                Return NumericUpDown3.Value * 1024
            Case 2
                Return NumericUpDown3.Value * 1048576
        End Select
    End Function

    Private Function CalcNoParts() As Integer
        If File.Exists(TextBox1.Text) = False Then Exit Function
        Return Math.Ceiling(FileLen(TextBox1.Text) / CalcPartSize())
    End Function


    Public Sub CheckValidInfo()
        Select Case TabControl1.SelectedIndex
            Case 0
                'Tab #1
                Dim iTest As Single
                If Not File.Exists(TextBox1.Text) Then
                    Label1.ForeColor = Color.Red
                    Label2.ForeColor = Color.Red
                    If TextBox1.Text = vbNullString Then
                        Label2.Text = Timer_No_Selected
                    Else
                        Label2.Text = Timer_File_Not_Exists
                    End If
                Else
                    Label1.ForeColor = Color.Black
                    Label2.ForeColor = Color.Black
                    Dim _FileInfo As FileInfo
                    _FileInfo = New FileInfo(TextBox1.Text)
                    Label2.Text = Timer_File_Info.Replace("|filename|", _FileInfo.Name).Replace("|filesize|", FormatSize(_FileInfo.Length))
                    _FileInfo = Nothing
                    iTest += 1
                End If

                If Directory.Exists(TextBox2.Text) = False Then
                    Label5.ForeColor = Color.Red
                Else
                    Label5.ForeColor = Color.Black
                    iTest += 1
                End If

                If CalcNoParts() <= 1 Or CalcNoParts() > 999 Then
                    Label3.ForeColor = Color.Red
                    RadioButton1.ForeColor = Color.Red
                    RadioButton2.ForeColor = Color.Red
                    Label4.ForeColor = Color.Red
                Else
                    Label3.ForeColor = Color.Black
                    RadioButton1.ForeColor = Color.Black
                    RadioButton2.ForeColor = Color.Black
                    Label4.ForeColor = Color.Black
                    iTest += 1
                End If

                'Final check
                If iTest = 3 And IsSplitPrg = False Then
                    IsSplitValid = True
                Else
                    IsSplitValid = False
                End If
                Button1.Enabled = IsSplitValid
            Case 1
                'Tab #2
                Dim iTest2 As Single
                If ListView1.Items.Count <= 1 Then
                    Label9.ForeColor = Color.Red
                Else
                    Label9.ForeColor = Color.Black
                    iTest2 += 1
                End If

                If Not Directory.Exists(CutFileName(TextBox3.Text)) Or IsFileNameValid(CutPath(TextBox3.Text)) = False Then
                    Label13.ForeColor = Color.Red
                Else
                    Label13.ForeColor = Color.Black
                    iTest2 += 1
                End If

                If iTest2 = 2 And IsMergePrg = False Then
                    Button6.Enabled = True
                Else
                    Button6.Enabled = False
                End If

                If ListView1.Items.Count = 0 Then
                    Button2.Enabled = False
                    Button5.Enabled = False
                    Button8.Enabled = False
                Else
                    Button2.Enabled = True
                    Button5.Enabled = True
                    Button8.Enabled = True
                End If
            Case 2
                'Tab #3
                Dim iTest2 As Single = 0

                If Not File.Exists(TextBox4.Text) Then
                    Label19.ForeColor = Color.Red
                    Label20.ForeColor = Color.Red
                    If TextBox4.Text = vbNullString Then
                        Label20.Text = Timer_No_Selected
                    Else
                        Label20.Text = Timer_File_Not_Exists
                    End If
                Else
                    Label19.ForeColor = Color.Black
                    Label20.ForeColor = Color.Black

                    Dim _FileInfo2 As FileInfo = New FileInfo(TextBox4.Text)
                    Label20.Text = Timer_File_Info.Replace("|filename|", _FileInfo2.Name).Replace("|filesize|", FormatSize(_FileInfo2.Length))
                    _FileInfo2 = Nothing
                    iTest2 += 1
                End If

                If Not Directory.Exists(CutFileName(TextBox6.Text)) Or IsFileNameValid(CutPath(TextBox6.Text)) = False Then
                    Label22.ForeColor = Color.Red
                Else
                    Label22.ForeColor = Color.Black
                    iTest2 += 1
                End If

                If TextBox5.Text = vbNullString Then
                    Label21.ForeColor = Color.Red
                    Label23.ForeColor = Color.Red
                Else
                    Label21.ForeColor = Color.Black
                    If RadioButton3.Checked = True Then
                        If TextBox7.Text = TextBox5.Text Then
                            Label23.ForeColor = Color.Black
                            iTest2 += 1
                        Else
                            Label23.ForeColor = Color.Red
                        End If
                    Else
                        Label23.ForeColor = Color.Black
                        iTest2 += 1
                    End If
                End If

                'Final check
                If iTest2 = 3 Then
                    Button15.Enabled = True
                Else
                    Button15.Enabled = False
                End If
            Case Else

        End Select
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        fDlg1.FileName = vbNullString
        fDlg1.Multiselect = False
        fDlg1.Filter = Btn3_Dlg_Filter
        fDlg1.ShowDialog()
        If fDlg1.FileName <> vbNullString Then TextBox1.Text = fDlg1.FileName
        CheckValidInfo()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        fDlg2.SelectedPath = vbNullString
        fDlg2.ShowDialog()
        If fDlg2.SelectedPath <> vbNullString Then TextBox2.Text = fDlg2.SelectedPath
        CheckValidInfo()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        CheckOptions()
        If Button1.Enabled = False Then Exit Sub

        Dim IsError As Boolean
        Dim _Parts As Int32 = CalcNoParts()
        If _Parts >= 20 Then
            If MsgBox(Btn1_Too_Many_Parts.Replace("|parts|", CalcNoParts), MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, Btn1_Too_Many_Parts_Title) = MsgBoxResult.No Then Exit Sub
        End If

        If File.Exists(TextBox2.Text & Slash(TextBox2.Text) & CutPath(TextBox1.Text) & ".001") Then
            If MsgBox(Btn1_File_Exists.Replace("|filename|", CutPath(TextBox1.Text)), MsgBoxStyle.Question + MsgBoxStyle.YesNo, Btn1_File_Exists_Title) = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        'Disable all controls
        IsSplitPrg = True
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Label1.Enabled = False
        Label2.Enabled = False
        Label5.Enabled = False
        Label3.Enabled = False
        Label4.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        NumericUpDown1.Enabled = False
        NumericUpDown2.Enabled = False
        ComboBox1.Enabled = False
        Button1.Enabled = False
        Button17.Enabled = False
        Button10.Enabled = True

        'ProgressBar1.Maximum = CalcNoParts()
        ProgressBar1.Value = 0

        'Progress

        'IsError = Not SplitFile(TextBox1.Text, TextBox2.Text, CalcPartSize, CalcBufferSize, CheckBox1.Checked, CheckBox3.Checked)

        Dim objThreadClass As New clsSplitThread(0, Me, TextBox1.Text, TextBox2.Text, CalcPartSize, CalcBufferSize, CheckBox1.Checked, CheckBox3.Checked)
        Dim objNewThread As New Threading.Thread(AddressOf objThreadClass.StartThread)
        objNewThread.IsBackground = True
        SplitStatus = 1
        objNewThread.Start()

        Do While SplitStatus = 1
            If StopSlit = False Then
                Application.DoEvents()
            Else
                StopSlit = False
                If objNewThread.IsAlive Then objNewThread.Abort()
                IsError = True
                Exit Do
            End If
        Loop

        If SplitStatus = 3 Then IsError = True
        If IsError = False Then If objNewThread.IsAlive Then objNewThread.Abort()
        objNewThread = Nothing
        objThreadClass = Nothing

        If CheckBox2.Checked = True Then File.Delete(TextBox1.Text)

        TextBox1.Enabled = True
        TextBox2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Label1.Enabled = True
        Label2.Enabled = True
        Label5.Enabled = True
        Label3.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        NumericUpDown1.Enabled = True
        NumericUpDown2.Enabled = True
        ComboBox1.Enabled = True
        CheckOptions()
        Button1.Enabled = True
        Button17.Enabled = True
        Button10.Enabled = False

        Label7.Text = Btn1_Lb7_Complete
        Label8.Text = vbNullString

        IsSplitPrg = False

        SplitStatus = 0

        If ClosePending = True Then Me.Close()

        CheckValidInfo()
        If IsError = True Then Exit Sub

        If MsgBox(Btn1_Complete_Msg.Replace("|filename|", CutPath(TextBox1.Text)).Replace("|parts|", _Parts).Replace("|chunksize|", FormatSize(CalcPartSize)), MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Btn1_Complete_Title") = MsgBoxResult.Yes Then
            Process.Start(TextBox2.Text)
        End If
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        fDlg1.FileName = vbNullString
        If CheckBox5.Checked = False Then
            fDlg1.Multiselect = True
            fDlg1.Filter = Btn7_Dlg_Filter
            fDlg1.ShowDialog()
            For Each _FileName As String In fDlg1.FileNames
                AddItemToList(_FileName)
            Next
        Else
            fDlg1.Multiselect = False
            fDlg1.Filter = Btn7_Auto_Dlg_Filter
            fDlg1.ShowDialog()
            If Not File.Exists(fDlg1.FileName) Then Exit Sub
            Dim _FileNames As String() = getMergeFiles(GetMergeFileName(fDlg1.FileName))
            If IsNothing(_FileNames) Then Exit Sub
            For Each _FileName As String In _FileNames
                AddItemToList(_FileName)
            Next
        End If
    End Sub

    Public Sub AddItemToList(ByVal FileName As String)
        If Not File.Exists(FileName) Then Exit Sub
        Dim NewItem As ListViewItem
        NewItem = New ListViewItem
        With NewItem
            .Text = CutPath(FileName)
            .SubItems.Add(FileName)
            .SubItems.Add(FileSize(FileName))
            .ToolTipText = FileName
        End With
        ListView1.Items.Add(NewItem)
        NewItem = Nothing
        If TextBox3.Text = vbNullString And ListView1.Items.Count > 0 Then TextBox3.Text = GetMergeFileName(ListView1.Items(0).SubItems(1).Text)
        CheckValidInfo()
    End Sub

    Public Function getMergeFiles(ByVal MergeFileName As String) As String()
        Dim _tempFiles As String() = Directory.GetFiles(Directory.GetParent(MergeFileName).ToString)
        Dim _MergeFiles As String()
        Dim _Index As Integer
        Array.Sort(_tempFiles)
        Do
            _Index += 1
            If Array.BinarySearch(_tempFiles, MergeFileName & "." & Format(_Index, "000")) >= 0 Then
                ReDim Preserve _MergeFiles(_Index - 1)
                _MergeFiles(_Index - 1) = MergeFileName & "." & Format(_Index, "000")
            Else
                Exit Do
            End If
        Loop
        Return _MergeFiles
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MoveListViewItem(ListView1, True)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        MoveListViewItem(ListView1, False)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        For i As Integer = 0 To ListView1.SelectedItems.Count - 1
            ListView1.SelectedItems(i).Remove()
        Next
        CheckValidInfo()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        CheckOptions()
        If Button6.Enabled = False Then Exit Sub
        If ListView1.Items.Count = 0 Then Exit Sub
        Dim IsError As Boolean
        Dim i As Short
        Dim _MergeFiles As String()
        For Each _Item As ListViewItem In ListView1.Items
            i += 1
            If Not File.Exists(_Item.SubItems(1).Text) Then _
                MsgBox(Btn6_File_Not_Exists.Replace("|filename|", _Item.Text), _
                       MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, Btn6_File_Not_Exists_Title) _
                     : Exit Sub
            ReDim Preserve _MergeFiles(i - 1)
            _MergeFiles(i - 1) = _Item.SubItems(1).Text
        Next
        If File.Exists(TextBox3.Text) Then
            If MsgBox(Btn6_File_Exists.Replace("|filename|", TextBox3.Text) _
                      , MsgBoxStyle.Question + MsgBoxStyle.YesNo, Btn6_File_Exists_Title) _
                      = MsgBoxResult.No Then Exit Sub
        End If

        Dim _Parts As Integer = ListView1.Items.Count
        'Disable all controls
        IsMergePrg = True
        ProgressBar2.Value = 0
        'ProgressBar2.Maximum = _MergeFiles.Length
        Button7.Enabled = False
        Button8.Enabled = False
        Button2.Enabled = False
        Button5.Enabled = False
        Button9.Enabled = False
        CheckBox5.Enabled = False
        CheckBox4.Enabled = False
        TextBox3.Enabled = False
        Label9.Enabled = False
        Label13.Enabled = False
        ListView1.Enabled = False
        Button6.Enabled = False
        Button16.Enabled = False
        Button11.Enabled = True

        'IsError = Not MergeFile(_MergeFiles, TextBox3.Text, CalcBufferSize)
        'Progress
        Dim objThreadClass As New clsMergeThread(1, Me, _MergeFiles, TextBox3.Text, CalcBufferSize, CheckBox4.Checked)
        Dim objNewThread As New Threading.Thread(AddressOf objThreadClass.StartThread)
        objNewThread.IsBackground = True
        MergeStatus = 1
        objNewThread.Start()
        Do While MergeStatus = 1
            If StopMerge = False Then
                Application.DoEvents()
            Else
                StopMerge = False
                If objNewThread.IsAlive Then objNewThread.Abort()
                IsError = True
                Exit Do
            End If
        Loop
        If MergeStatus = 3 Then IsError = True
        If IsError = False And objNewThread.IsAlive Then objNewThread.Abort()
        objNewThread = Nothing
        objThreadClass = Nothing

        Button7.Enabled = True
        Button8.Enabled = True
        Button2.Enabled = True
        Button5.Enabled = True
        Button9.Enabled = True
        CheckBox5.Enabled = True
        CheckBox4.Enabled = True
        TextBox3.Enabled = True
        Label9.Enabled = True
        Label13.Enabled = True
        ListView1.Enabled = True
        Button6.Enabled = True
        Button16.Enabled = True
        If CheckBox4.Checked = True Then ListView1.Items.Clear()
        Label11.Text = Btn6_Lb11_Complete
        Label10.Text = vbNullString
        IsMergePrg = False
        Button11.Enabled = False

        MergeStatus = 0

        If ClosePending = True Then Me.Close()

        CheckValidInfo()
        If IsError = True Then Exit Sub
        If MsgBox(Btn6_Complete_Msg.Replace("|filename|", CutPath(TextBox3.Text)).Replace("|filesize|", FormatSize(FileLen(TextBox3.Text))).Replace("|parts|", _Parts), MsgBoxStyle.Information + MsgBoxStyle.YesNo, Btn6_Complete_Title) = MsgBoxResult.Yes Then
            System.Diagnostics.Process.Start(TextBox3.Text)
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If ListView1.Items.Count >= 1 Then fDlg3.FileName = GetMergeFileName(ListView1.Items(0).Text)
        fDlg3.ShowDialog()
        If Directory.Exists(CutFileName(fDlg3.FileName)) Then TextBox3.Text = fDlg3.FileName
        CheckValidInfo()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("mailto:luungoc2005@yahoo.com")
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        StopSlit = True
        Button10.Enabled = False
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        StopMerge = True
        Button11.Enabled = False
    End Sub

    Private Sub TextBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox1.DragDrop
        Dim s() As String = e.Data.GetData("FileDrop", False)
        TextBox1.Text = s(0)
    End Sub

    Private Sub TextBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        CheckValidInfo()
        If TextBox2.Text = vbNullString Then TextBox2.Text = CutFileName(TextBox1.Text)
    End Sub

    Private Sub ListView1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListView1.DragDrop
        Dim s() As String = e.Data.GetData("FileDrop", False)
        For Each a As String In s
            AddItemToList(a)
        Next

        If CheckBox5.Checked = True And s.Length = 1 Then
            Dim _FileNames As String() = getMergeFiles(GetMergeFileName(s(0)))
            If IsNothing(_FileNames) Then Exit Sub
            For Each _FileName As String In _FileNames
                If _FileName <> s(0) Then AddItemToList(_FileName)
            Next
        End If
    End Sub

    Private Sub ListView1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListView1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        CheckValidInfo()
    End Sub

    Private Sub TextBox3_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox3.DragDrop

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        CheckValidInfo()
    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        CheckValidInfo()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        CheckValidInfo()
    End Sub

    Private Sub NumericUpDown2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown2.ValueChanged
        CheckValidInfo()
    End Sub

#End Region

#Region "Encrypt/Decrypt"

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        CheckOptions()
        If Button15.Enabled = False Then Exit Sub

    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        CheckValidInfo()
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        CheckValidInfo()
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        CheckValidInfo()
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        CheckOptions()
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        CheckValidInfo()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        RadioButton3.Checked = True
        TextBox4.Text = vbNullString
        TextBox6.Text = vbNullString
        TextBox5.Text = vbNullString
        TextBox7.Text = vbNullString
        CheckOptions()
    End Sub
#End Region

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        ChangeLang(LangFiles(ComboBox3.SelectedIndex))
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        CheckOptions()
    End Sub


    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        CheckValidInfo()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        ListView1.Items.Clear()
        TextBox3.Clear()
        CheckValidInfo()
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        TextBox1.Text = vbNullString
        TextBox2.Text = vbNullString
        NumericUpDown1.Value = 1
        NumericUpDown2.Value = 1
        CheckOptions()
    End Sub

    Private Sub LnkAsso_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LnkAsso.LinkClicked
        AssociateMerge()
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            Dim _Directory = CutFileName(TextBox3.Text)
            If _Directory.Length > 1 Then
                If Not Directory.Exists(_Directory) Then
                    Directory.CreateDirectory(_Directory)
                End If

                Dim _TextFile As String = _Directory & Slash(_Directory) & "Merge.bat"

                Dim objText As StreamWriter
                Dim objTextStream As FileStream
                objTextStream = New FileStream(_TextFile, FileMode.Create)
                objText = New StreamWriter(objTextStream)

                objText.WriteLine("@echo off" & vbCrLf & _
                   "echo File Splitter v1.00 beta" & vbCrLf & _
    "echo Made by luungoc2005 (luungoc2005@yahoo.com)" & vbCrLf & _
    "pause")

                Dim strMerge As String, strDel As String
                strMerge = "copy /b "
                strDel = vbNullString

                For i As Integer = 0 To ListView1.Items.Count - 1
                    If i = 0 Then
                        strMerge = strMerge & ListView1.Items(i).Text
                        strDel = strDel & "del " & ListView1.Items(i).Text
                    Else
                        strDel = strDel & vbCrLf & "del " & ListView1.Items(i).Text
                        strMerge = strMerge & " + " & ListView1.Items(i).Text
                    End If
                Next

                objText.WriteLine(strMerge)
                objText.WriteLine(strDel)
                objText.Flush()
                objTextStream.Flush()
                objText.Close()
                objTextStream.Close()
                objText.Dispose()
                objTextStream.Dispose()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        fDlg4.FileName = vbNullString
        fDlg4.Multiselect = False
        fDlg4.Filter = Btn3_Dlg_Filter
        fDlg4.ShowDialog()
        If fDlg4.FileName <> vbNullString Then TextBox4.Text = fDlg4.FileName
        CheckValidInfo()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        fDlg5.FileName = vbNullString
        fDlg5.Filter = Btn3_Dlg_Filter
        fDlg5.ShowDialog()
        If fDlg5.FileName <> vbNullString Then TextBox6.Text = fDlg5.FileName
        CheckValidInfo()
    End Sub
End Class
