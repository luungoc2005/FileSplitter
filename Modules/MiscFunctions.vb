Imports System.Security.Cryptography

Module MiscFunctions
    Function CutPath(ByVal FilePath As String) As String
        CutPath = vbNullString
        For X = Len(FilePath$) To 3 Step -1
            If Mid$(FilePath$, X, 1) = "\" Then
                CutPath = Mid$(FilePath$, X + 1)
                Exit For
            End If
        Next X
    End Function

    Function CutFileName(ByVal FilePath As String) As String
        CutFileName = vbNullString
        For X = Len(FilePath) To 3 Step -1
            If Mid$(FilePath, X, 1) = "\" Then
                CutFileName = Mid$(FilePath, 1, X - 1)
                Exit For
            End If
        Next X
    End Function

    Function IsFileNameValid(ByVal FileName As String) As Boolean
        If FileName = vbNullString Then Return False
        IsFileNameValid = True
        Dim sTest As String = FileName
        Dim InvalidChars As Char() = New Char() {"\", "/", ":", "*", "?", """", "<", ">", "|"}
        For Each _char As Char In InvalidChars
            If sTest.Contains(_char) Then Return False
        Next
    End Function

    Public Function HashKey(ByVal KeyToHash As String) As Byte()
        Dim key() As Byte = System.Text.Encoding.Unicode.GetBytes(KeyToHash.ToCharArray)
        Dim SHAObj As SHA256Managed = New SHA256Managed
        Dim arrHash As Byte()
        arrHash = SHAObj.ComputeHash(key)
        SHAObj = Nothing
        ReDim Preserve arrHash(31)
        Return arrHash
    End Function

    Public Function HashKey512(ByVal KeyToHash As String) As Byte()
        Dim key() As Byte = System.Text.Encoding.Unicode.GetBytes(KeyToHash.ToCharArray)
        Dim SHAObj As SHA512Managed = New SHA512Managed
        Dim arrHash As Byte()
        arrHash = SHAObj.ComputeHash(key)
        SHAObj = Nothing
        Return arrHash
    End Function

    Function GetMergeFileName(ByVal FileName As String) As String
        If FileName = vbNullString Then Return vbNullString
        If FileName.Substring(FileName.Length - 4, 4) = ".001" Then
            Return FileName.Substring(0, FileName.Length - 4)
        Else
            Return FileName
        End If
    End Function

    Function FileSize(ByVal FilePath As String) As String
        Dim iSize As Long = FileLen(FilePath)
        Return FormatSize(iSize)
    End Function

    Function FormatSize(ByVal FileSize As Long) As String
        Dim iSize As Long = FileSize
        If iSize < 1024 Then
            Return iSize & " bytes"
        ElseIf iSize >= 1024 And iSize < 1048576 Then
            Return Math.Round(iSize / 1024, 2) & " KB"
        ElseIf iSize >= 1048576 And iSize < 1073741824 Then
            Return Math.Round(iSize / 1048576, 2) & " MB"
        Else
            Return Math.Round(iSize / 1073741824, 2) & " GB"
        End If
    End Function

    Function Slash(ByVal Path As String) As String
        Return IIf(Right(Path, 1) = "\", "", "\")
    End Function

    Function NumberToString(ByVal NumberToConvert As Integer) As String
        If NumberToConvert > 18278 Then Return vbNullString
        Dim c1 As Int32, c2 As Int32, c3 As Int32 : c1 = 0 : c2 = 0 : c3 = 0
        Dim intTemp As Int32 : intTemp = NumberToConvert
        c1 = intTemp Mod 676
        intTemp = intTemp - (676 * Math.Floor(intTemp / 676))
        c2 = intTemp Mod 26
        intTemp = intTemp - (26 * Math.Floor(intTemp / 26))
        c3 = intTemp
        Return Chr(c1 + 96) & Chr(c2 + 96) & Chr(c3 + 96)
    End Function

    Function StringToNumber(ByVal StringToConvert As String) As Integer
        If StringToConvert = vbNullString Or StringToConvert.Length > 3 Then Return 0
        Dim sTemp As String : sTemp = StringToConvert
        Do Until sTemp.Length = 3
            sTemp = "_" & sTemp
        Loop
        Dim c1 As Int32, c2 As Int32, c3 As Int32 : c1 = 0 : c2 = 0 : c3 = 0
        Dim s1 As String, s2 As String, s3 As String
        s1 = Left(sTemp, 1)
        s2 = sTemp.Substring(1, 1)
        s3 = Right(sTemp, 1)
        If s1 <> "_" Then c1 = Asc(s1) - 96
        If s2 <> "_" Then c2 = Asc(s2) - 96
        If s3 <> "_" Then c3 = Asc(s3) - 96
        Return (c1 * 26 * 26) + (c2 * 26) + c3
    End Function
End Module
