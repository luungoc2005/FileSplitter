Imports System.Security.Cryptography
Imports System.IO
Public Class clsEncryptThread
    Private _MainWindow As Form
    Private _ThreadIndex As Integer
    Private _Args() As Object

    Private Delegate Sub UpdateProgress(ByVal ThreadIndex As Integer, ByVal Text1 As String, ByVal Text2 As String)
    Private Delegate Sub UpdateBar(ByVal BytesProcessed As Long, ByVal BytesTotal As Long)
    Private Delegate Sub UpdateComplete(ByVal ThreadIndex As Integer, ByVal IsError As Boolean)

    Private _UpdateProgress As UpdateProgress
    Private _UpdateBar As UpdateBar
    Private _UpdateComplete As UpdateComplete

    Dim _Input As String, _Output As String, _BufferSize As Long, _DelFile As Boolean, _Password As String, _Encrypt As Boolean

    Public Sub New(ByVal ThreadIndex As Integer, ByVal MainWindow As Form, ByVal InputFile As String, ByVal OutputFile As String, ByVal BufferSize As Long, ByVal EncKey As String, ByVal DelFile As Boolean, ByVal Encrypt As Boolean)
        _ThreadIndex = ThreadIndex
        _MainWindow = MainWindow

        _UpdateProgress = AddressOf Form1.ReceiveEncryptProgress
        _UpdateBar = AddressOf Form1.UpdateEncPrgBar
        _UpdateComplete = AddressOf Form1.EncComplete

        _Input = InputFile
        _Output = OutputFile
        _BufferSize = BufferSize
        _DelFile = DelFile
        _Password = EncKey
        _Encrypt = Encrypt

    End Sub

    Public Sub ChangeStatusEnc(ByVal Text1 As String, ByVal Text2 As String)
        ReDim _Args(2)
        _Args(0) = _ThreadIndex
        _Args(1) = Text1
        _Args(2) = Text2
        _MainWindow.Invoke(_UpdateProgress, _Args)
    End Sub

    Public Sub UpdateProgressEnc(ByVal bytesProcessed As Long, ByVal bytesTotal As Long)
        ReDim _Args(1)
        _Args(0) = bytesProcessed
        _Args(1) = bytesTotal
        _MainWindow.Invoke(_UpdateBar, _Args)
    End Sub

    Public Sub EncryptComplete(ByVal IsError As Boolean)
        ReDim _Args(1)
        _Args(0) = _ThreadIndex
        _Args(1) = IsError
        _MainWindow.Invoke(_UpdateComplete, _Args)
    End Sub

    Public Sub StartThread()
        Dim _EncryptObj As RijndaelManaged = New RijndaelManaged
        Dim _CryptoStream As CryptoStream
        Dim _InputStream As FileStream
        Dim _OutputStream As FileStream
        Dim _FileSize As Long
        Dim _CurrentBuffer As Long
        Dim _bytesProcessed As Long = 0
        Dim IsError As Boolean
        Dim Key() As Byte
        Dim IV() As Byte
        Dim buffer() As Byte = New Byte() {}

        Try
            If File.Exists(_Output) Then File.Delete(_Output)
            _FileSize = FileLen(_Input)
            _InputStream = New FileStream(_Input, FileMode.Open)
            _OutputStream = New FileStream(_Output, FileMode.Create)

            Key = HashKey(_Password)
            IV = GetIV()
            If _Encrypt = True Then
                _CryptoStream = New CryptoStream(_OutputStream, _EncryptObj.CreateEncryptor(Key, IV), CryptoStreamMode.Write)
            Else
                _CryptoStream = New CryptoStream(_OutputStream, _EncryptObj.CreateDecryptor(Key, IV), CryptoStreamMode.Write)
            End If

            ReDim buffer(_BufferSize - 1)
            Do While _bytesProcessed < _FileSize
                If _Encrypt = True Then
                    ChangeStatusEnc(Btn15_Enc_Text1_Working, Btn15_Enc_Text2_Working.Replace("|filename|", CutPath(_Output)).Replace("|prg|", Math.Round(_bytesProcessed / _FileSize * 100)))
                Else
                    ChangeStatusEnc(Btn15_Dec_Text1_Working, Btn15_Dec_Text2_Working.Replace("|filename|", CutPath(_Output)).Replace("|prg|", Math.Round(_bytesProcessed / _FileSize * 100)))
                End If
                _CurrentBuffer = _InputStream.Read(buffer, 0, buffer.Length)
                _CryptoStream.Write(buffer, 0, _CurrentBuffer)
                _bytesProcessed = _bytesProcessed + _CurrentBuffer
                _CryptoStream.Flush()
                UpdateProgressEnc(_bytesProcessed, _FileSize)
            Loop
        Catch ex As Exception
            IsError = True
        End Try
    End Sub

    Public Function GetIV() As Byte()
        Dim bResult(0 To 15) As Byte
        bResult(0) = 235
        bResult(1) = 160
        bResult(2) = 151
        bResult(3) = 247
        bResult(4) = 129
        bResult(5) = 183
        bResult(6) = 191
        bResult(7) = 141
        bResult(8) = 10
        bResult(9) = 63
        bResult(10) = 43
        bResult(11) = 107
        bResult(12) = 126
        bResult(13) = 160
        bResult(14) = 205
        bResult(15) = 67
        Return bResult
    End Function
End Class
