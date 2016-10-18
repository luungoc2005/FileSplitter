Imports System.IO
Public Class clsSplitThread
    Private _MainWindow As Form
    Private _ThreadIndex As Integer
    Private _Args() As Object

    Private Delegate Sub UpdateProgress(ByVal ThreadIndex As Integer, ByVal Text1 As String, ByVal Text2 As String)
    Private Delegate Sub UpdateBar(ByVal BytesProcessed As Long, ByVal BytesTotal As Long)
    Private Delegate Sub ReportComplete(ByVal ThreadIndex As Integer, ByVal IsError As Boolean)

    Private _UpdateProgress As UpdateProgress
    Private _UpdateBar As UpdateBar
    Private _ReportComplete As ReportComplete

    'Split:
    Dim InputFile As String, OutputPath As String, ChunkSize As Long, BufferSize As Long, CreateMerge As Boolean, DelFile As Boolean

    Private Function CalcFragments(ByVal FileName As String, ByVal ChunkSize As Long) As Integer
        If File.Exists(FileName) = False Then Exit Function
        Return Math.Floor(FileLen(FileName) / ChunkSize)
    End Function

    Private Function CalcParts(ByVal FileName As String, ByVal ChunkSize As Long) As Integer
        If File.Exists(FileName) = False Then Exit Function
        Return Math.Ceiling(FileLen(FileName) / ChunkSize)
    End Function

    Public Sub New(ByVal ThreadIndex As Integer, ByVal MainWindow As Form1, ByVal sInputFile As String, ByVal sOutputPath As String, ByVal sChunkSize As Long, ByVal sBufferSize As Long, ByVal sCreateMerge As Boolean, Optional ByVal sDelFile As Boolean = False)
        _MainWindow = MainWindow
        _ThreadIndex = ThreadIndex
        _UpdateProgress = AddressOf Form1.ReceiveSplitProgress
        _UpdateBar = AddressOf Form1.UpdateSplitPrgBar
        _ReportComplete = AddressOf Form1.SplitComplete

        InputFile = sInputFile
        OutputPath = sOutputPath
        ChunkSize = sChunkSize
        BufferSize = sBufferSize
        CreateMerge = sCreateMerge
        DelFile = sDelFile

        MsgBox(MainWindow.Text)
    End Sub

    Private Sub ChangeStatusSplit(ByVal Text1 As String, ByVal Text2 As String)
        ReDim _Args(2)
        _Args(0) = _ThreadIndex
        _Args(1) = Text1
        _Args(2) = Text2
        _MainWindow.Invoke(_UpdateProgress, _Args)
    End Sub

    Private Sub UpdateProgressSplit(ByVal BytesProcessed As Long, ByVal BytesTotal As Long)
        ReDim _Args(1)
        _Args(0) = BytesProcessed
        _Args(1) = BytesTotal
        _MainWindow.Invoke(_UpdateBar, _Args)
    End Sub

    Private Sub CompleteSplit(ByVal IsError As Boolean)
        ReDim _Args(1)
        _Args(0) = _ThreadIndex
        _Args(1) = IsError
        _MainWindow.Invoke(_ReportComplete, _Args)
    End Sub

    Public Sub StartThread()

        Dim IsError As Boolean

        Dim totalBytes As Long
        Dim processedBytes As Long

        Dim sFileName As String
        Dim sDir As String
        Dim InputStream As FileStream
        Dim OutputStream As FileStream
        Dim objRead As BinaryReader
        Dim objWrite As BinaryWriter
        Dim sOutputFile As String
        Dim sBaseName As String
        Dim iChunkSize As Long
        Dim iFragments As Integer
        Dim _Index As Integer
        Dim _Buffer As Byte() = New Byte() {}
        Dim _FinalBytes As Long
        Dim _BufferSize As Long
        Dim NoParts As Integer

        Dim Text1 As String, Text2 As String

        Dim strMerge As String
        Dim strDel As String
        strMerge = "Copy/b "

        Try
            _BufferSize = BufferSize
            sFileName = InputFile
            sDir = OutputPath
            iFragments = CalcFragments(sFileName, ChunkSize)
            iChunkSize = ChunkSize
            NoParts = CalcParts(sFileName, ChunkSize)
            sBaseName = CutPath(sFileName)
            InputStream = New FileStream(sFileName, FileMode.Open)
            objRead = New BinaryReader(InputStream)
            If Not Directory.Exists(sDir) Then Directory.CreateDirectory(sDir)
            objRead.BaseStream.Seek(0, SeekOrigin.Begin)
            totalBytes = FileLen(sFileName)
            _FinalBytes = totalBytes - (iFragments * iChunkSize)

            For _Index = 1 To iFragments
                sOutputFile = sDir & Slash(sDir) & sBaseName & "." & Format(_Index, "000")
                Text1 = Btn1_Lb7_Text_Working
                Text2 = Btn1_Lb8_Text_Working.Replace("|filename|", CutPath(sOutputFile)).Replace("|prg|", _Index & "/" & NoParts)
                ChangeStatusSplit(Text1, Text2)

                If File.Exists(sOutputFile) Then File.Delete(sOutputFile)
                OutputStream = New FileStream(sOutputFile, FileMode.Create)
                objWrite = New BinaryWriter(OutputStream)

                If iChunkSize <= _BufferSize Then
                    ReDim _Buffer(iChunkSize - 1)
                    objRead.Read(_Buffer, 0, iChunkSize)
                    'objRead.BaseStream.Seek(0, SeekOrigin.Current)
                    objWrite.Write(_Buffer)
                    OutputStream.Flush()
                    processedBytes = processedBytes + iChunkSize
                    UpdateProgressSplit(processedBytes, totalBytes)
                Else
                    Dim _bFragments As Long
                    Dim _bLastBytes As Long
                    Dim _index2 As Long
                    _bFragments = Math.Floor(iChunkSize / _BufferSize)
                    _bLastBytes = iChunkSize - (_bFragments * _BufferSize)
                    For _index2 = 1 To _bFragments
                        Text2 = Btn1_Lb8_Text_Working.Replace("|filename|", CutPath(sOutputFile)).Replace("|prg|", _Index & "/" & NoParts) & " - " & _
    Math.Round(_index2 / _bFragments * 100) & "%"
                        ChangeStatusSplit(Text1, Text2)
                        ReDim _Buffer(_BufferSize - 1)
                        objRead.Read(_Buffer, 0, _BufferSize)
                        'objRead.BaseStream.Seek(0, SeekOrigin.Current)
                        objWrite.Write(_Buffer)
                        OutputStream.Flush()
                        processedBytes = processedBytes + _BufferSize
                        UpdateProgressSplit(processedBytes, totalBytes)
                        'Application.DoEvents()
                    Next
                    If _bLastBytes > 0 Then
                        ReDim _Buffer(_bLastBytes - 1)
                        objRead.Read(_Buffer, 0, _Buffer.Length)
                        objWrite.Write(_Buffer)
                        OutputStream.Flush()
                        processedBytes = processedBytes + _bLastBytes
                        UpdateProgressSplit(processedBytes, totalBytes)
                        'Application.DoEvents()
                    End If
                End If

                objWrite.Close()
                OutputStream.Close()

                strDel = strDel & vbCrLf & "Del " & """" & CutPath(sOutputFile) & """"
                If _FinalBytes = 0 And _Index = iFragments Then
                    strMerge = strMerge & """" & CutPath(sOutputFile) & """" & " " & """" & sBaseName & """"
                    strDel = strDel & vbCrLf & "Del Merge.bat"
                Else
                    strMerge = strMerge & """" & CutPath(sOutputFile) & """" & " + "
                End If

                'Application.DoEvents()
                Threading.Thread.Sleep(1)
            Next
            If _FinalBytes > 0 Then
                sOutputFile = sDir & Slash(sDir) & sBaseName & "." & Format(_Index, "000")
                Text1 = Btn1_Lb7_Text_Working.Replace("|percent|", Math.Ceiling(iFragments + 1 / NoParts * 100))
                Text2 = Btn1_Lb8_Text_Working.Replace("|filename|", CutPath(sOutputFile)).Replace("|prg|", iFragments + 1 & "/" & NoParts)
                ChangeStatusSplit(Text1, Text2)

                If File.Exists(sOutputFile) Then File.Delete(sOutputFile)
                OutputStream = New FileStream(sOutputFile, FileMode.Create)
                objWrite = New BinaryWriter(OutputStream)

                If _FinalBytes <= _BufferSize Then
                    ReDim _Buffer(_FinalBytes - 1)
                    objRead.Read(_Buffer, 0, _FinalBytes)
                    objWrite.Write(_Buffer)
                    OutputStream.Flush()
                    objWrite.Close()
                    OutputStream.Close()
                    processedBytes = processedBytes + _FinalBytes
                    UpdateProgressSplit(processedBytes, totalBytes)
                Else
                    Dim _bFragments As Long
                    Dim _bLastBytes As Long
                    Dim _index2 As Long
                    _bFragments = Math.Floor(_FinalBytes / _BufferSize)
                    _bLastBytes = _FinalBytes - (_bFragments * _BufferSize)
                    For _index2 = 1 To _bFragments
                        Text2 = Btn1_Lb8_Text_Working.Replace("|filename|", CutPath(sOutputFile)).Replace("|prg|", _Index & "/" & NoParts) & " - " & _
                            Math.Round(_index2 / _bFragments * 100) & "%"
                        ChangeStatusSplit(Text1, Text2)

                        ReDim _Buffer(_BufferSize - 1)
                        objRead.Read(_Buffer, 0, _BufferSize)
                        objRead.BaseStream.Seek(0, SeekOrigin.Current)
                        objWrite.Write(_Buffer)
                        OutputStream.Flush()
                        processedBytes = processedBytes + _BufferSize
                        UpdateProgressSplit(processedBytes, totalBytes)
                        'Application.DoEvents()
                    Next
                    If _bLastBytes > 0 Then
                        ReDim _Buffer(_bLastBytes - 1)
                        objRead.Read(_Buffer, 0, _Buffer.Length)
                        objWrite.Write(_Buffer)
                        OutputStream.Flush()
                        processedBytes = processedBytes + _bLastBytes
                        UpdateProgressSplit(processedBytes, totalBytes)
                        'Application.DoEvents()
                    End If
                End If
                objWrite.Close()
                OutputStream.Close()

                strMerge = strMerge & """" & CutPath(sOutputFile) & """" & " " & """" & sBaseName & """"
                strDel = strDel & vbCrLf & "Del " & """" & CutPath(sOutputFile) & """"
                strDel = strDel & vbCrLf & "Del Merge.bat"

                'Application.DoEvents()
                Threading.Thread.Sleep(100)
            End If
            InputStream.Close()
            OutputStream.Close()
            objRead.Close()

        Catch t_ex As Threading.ThreadAbortException
            IsError = True
        Catch m_Ex As System.InsufficientMemoryException
            MsgBox(Error_Msg.Replace("|msg|", Error_InMemory), MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, Error_Title)
            IsError = True
        Catch ex As Exception
            MsgBox(Error_Msg.Replace("|msg|", ex.Message), MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, Error_Title)
            IsError = True
        End Try

        If CreateMerge = True Then
            Dim sMergeOutput As String
            Dim objWText As StreamWriter
            Dim objText As FileStream
            sMergeOutput = OutputPath & Slash(OutputPath) & "Merge.bat"
            If File.Exists(sMergeOutput) Then File.Delete(sMergeOutput)
            objText = New FileStream(sMergeOutput, FileMode.Create)
            objWText = New StreamWriter(objText)
            objWText.WriteLine("@echo off" & vbCrLf & _
                               "echo File Splitter v1.00 beta" & vbCrLf & _
            "echo Made by luungoc2005 (luungoc2005@yahoo.com)" & vbCrLf & _
            "pause")
            objWText.WriteLine(strMerge)
            objWText.WriteLine(IIf(DelFile = True, strDel, vbNullString))
            objText.Flush()
            objWText.Close()
            objText.Close()
            objText = Nothing
            objWText = Nothing
        End If
        CompleteSplit(IsError)
    End Sub
End Class
