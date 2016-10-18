Imports System.IO
Public Class clsMergeThread
    Private _MainWindow As Form
    Private _ThreadIndex As Integer
    Private _Args() As Object

    Private Delegate Sub UpdateProgress(ByVal ThreadIndex As Integer, ByVal Text1 As String, ByVal Text2 As String)
    Private Delegate Sub UpdateBar(ByVal BytesProcessed As Long, ByVal BytesTotal As Long)
    Private Delegate Sub ReportComplete(ByVal ThreadIndex As Integer, ByVal IsError As Boolean)

    Private _UpdateProgress As UpdateProgress
    Private _UpdateBar As UpdateBar
    Private _ReportComplete As ReportComplete

    'Merge:
    Dim _MergeFiles() As String, sResult As String, BufferSize As Long, _DelFile As Boolean

    Public Sub New(ByVal ThreadIndex As Integer, ByVal MainWindow As Form, ByVal MergeFiles() As String, ByVal Result As String, ByVal iBuffer As Long, ByVal DelFile As Boolean)
        _ThreadIndex = ThreadIndex
        _MainWindow = MainWindow
        _UpdateProgress = AddressOf Form1.ReceiveMergeProgress
        _UpdateBar = AddressOf Form1.UpdateMergePrgBar
        _ReportComplete = AddressOf Form1.MergeComplete
        _DelFile = DelFile

        _MergeFiles = MergeFiles.Clone
        sResult = Result
        BufferSize = iBuffer
    End Sub

    Public Sub ChangeStatusMerge(ByVal Text1 As String, ByVal Text2 As String)
        ReDim _Args(2)
        _Args(0) = _ThreadIndex
        _Args(1) = Text1
        _Args(2) = Text2
        _MainWindow.Invoke(_UpdateProgress, _Args)
    End Sub

    Public Sub UpdateProgressMerge(ByVal BytesProcessed As Long, ByVal BytesTotal As Long)
        ReDim _Args(1)
        _Args(0) = BytesProcessed
        _Args(1) = BytesTotal
        _MainWindow.Invoke(_UpdateBar, _Args)
    End Sub

    Public Sub MergeComplete(ByVal IsError As Boolean)
        ReDim _Args(1)
        _Args(0) = _ThreadIndex
        _Args(1) = IsError
        _MainWindow.Invoke(_ReportComplete, _Args)
    End Sub

    Public Sub StartThread()
        Dim IsError As Boolean

        Dim bytesTotal As Long
        Dim bytesProcessed As Long
        Dim InputStream As FileStream
        Dim OutputStream As FileStream
        Dim objRead As BinaryReader
        Dim objWrite As BinaryWriter
        Dim _Index As Short
        Dim buffer() As Byte = New Byte() {}
        Dim FileSize As Long
        Dim MergedName As String
        Dim _BufferSize As Long

        For Each _cfSize As String In _MergeFiles
            If IO.File.Exists(_cfSize) Then bytesTotal = bytesTotal + FileLen(_cfSize)
        Next

        Try
            _BufferSize = BufferSize
            MergedName = sResult
            If File.Exists(MergedName) Then File.Delete(MergedName)
            OutputStream = New FileStream(MergedName, FileMode.Create)
            objWrite = New BinaryWriter(OutputStream)
            For _Index = 0 To _MergeFiles.Length - 1
                If Not File.Exists(_MergeFiles(_Index)) Then
                    MsgBox(Btn6_File_Not_Exists_Title.Replace("|filename|", _MergeFiles(_Index)), MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, _
                           Btn6_File_Not_Exists_Title)
                    Exit For
                End If

                ChangeStatusMerge(Btn6_Lb11_Text_Working, Btn6_Lb10_Text_Working.Replace("|filename|", CutPath(_MergeFiles(_Index))).Replace("|prg|", _Index + 1 & "/" & _MergeFiles.Length))
                FileSize = FileLen(_MergeFiles(_Index))
                InputStream = New FileStream(_MergeFiles(_Index), FileMode.Open)
                objRead = New BinaryReader(InputStream)

                If FileSize <= _BufferSize Then
                    objRead.BaseStream.Seek(0, SeekOrigin.Begin)
                    ReDim buffer(FileSize - 1)
                    objRead.Read(buffer, 0, buffer.Length)
                    objWrite.Write(buffer)
                    OutputStream.Flush()
                    bytesProcessed = bytesProcessed + FileSize
                    UpdateProgressMerge(bytesProcessed, bytesTotal)
                Else
                    Dim _index2 As Long
                    Dim _bFragments As Long
                    Dim _LastBytes As Long
                    _bFragments = Math.Floor(FileSize / _BufferSize)
                    _LastBytes = FileSize - (_bFragments * _BufferSize)
                    objRead.BaseStream.Seek(0, SeekOrigin.Current)
                    For _index2 = 1 To _bFragments
                        ChangeStatusMerge(Btn6_Lb11_Text_Working, _
                                        Btn6_Lb10_Text_Working.Replace("|filename|", CutPath(_MergeFiles(_Index))).Replace("|prg|", _Index + 1 & "/" & _MergeFiles.Length) & " - " & _
    Math.Round(_index2 / _bFragments * 100) & "%")
                        ReDim buffer(_BufferSize - 1)
                        objRead.Read(buffer, 0, buffer.Length)
                        objWrite.Write(buffer)
                        OutputStream.Flush()
                        bytesProcessed = bytesProcessed + _BufferSize
                        UpdateProgressMerge(bytesProcessed, bytesTotal)
                    Next
                    If _LastBytes > 0 Then
                        ReDim buffer(_LastBytes - 1)
                        objRead.Read(buffer, 0, buffer.Length)
                        objWrite.Write(buffer)
                        OutputStream.Flush()
                        bytesProcessed = bytesProcessed + _LastBytes
                        UpdateProgressMerge(bytesProcessed, bytesTotal)
                    End If
                End If

                objRead.Close()
                InputStream.Close()

                'If _DelFile = True Then File.Delete(_MergeFiles(_Index))
            Next

            OutputStream.Close()
            objWrite.Close()

            If _DelFile = True Then
                For Each _FileToDel As String In _MergeFiles
                    If File.Exists(_FileToDel) Then File.Delete(_FileToDel)
                Next
            End If
        Catch t_ex As Threading.ThreadAbortException
            IsError = True
        Catch m_Ex As System.InsufficientMemoryException
            MsgBox(Error_Msg.Replace("|msg|", Error_InMemory), MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, Error_Title)
            IsError = True
        Catch ex As Exception
            MsgBox(Error_Msg.Replace("|msg|", ex.Message), MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, Error_Title)
            IsError = True
        End Try

        MergeComplete(IsError)
    End Sub
End Class
