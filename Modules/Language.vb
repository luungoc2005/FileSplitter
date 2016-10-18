Imports System.Text
Module Language

    Public Error_Msg As String
    Public Error_Title As String
    Public Error_InMemory As String

    Public Timer_No_Selected As String
    Public Timer_File_Not_Exists As String
    Public Timer_File_Info As String

    Public Btn3_Dlg_Filter As String

    Public Btn1_Too_Many_Parts As String
    Public Btn1_Too_Many_Parts_Title As String
    Public Btn1_File_Exists As String
    Public Btn1_File_Exists_Title As String
    Public Btn1_Lb7_Text_Working As String
    Public Btn1_Lb8_Text_Working As String
    Public Btn1_Lb7_Complete As String
    Public Btn1_Complete_Msg As String
    Public Btn1_Complete_Title As String

    Public Btn7_Dlg_Filter As String
    Public Btn7_Auto_Dlg_Filter As String

    Public Btn6_File_Exists As String
    Public Btn6_File_Exists_Title As String
    Public Btn6_File_Not_Exists As String
    Public Btn6_File_Not_Exists_Title As String
    Public Btn6_Lb11_Text_Working As String
    Public Btn6_Lb10_Text_Working As String
    Public Btn6_Lb11_Complete As String
    Public Btn6_Complete_Msg As String
    Public Btn6_Complete_Title As String

    Public Btn15_File_Exists As String
    Public Btn15_File_Exists_Title As String
    Public Btn15_Enc_Text1_Working As String
    Public Btn15_Dec_Text1_Working As String
    Public Btn15_Enc_Text2_Working As String
    Public Btn15_Dec_Text2_Working As String
    Public Btn15_Complete As String
    Public Btn15_Enc_Complete_Msg As String
    Public Btn15_Enc_Complete_Title As String
    Public Btn15_Dec_Complete_Msg As String
    Public Btn15_Dec_Invalid_Pass As String
    Public Btn15_Enc_Dlg_Filter As String
    Public Btn15_Dec_Dlg_Filter As String

    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Int32, _
    ByVal lpFileName As String) As Int32

    Private Function GetLangFile(ByVal FileName As String) As String
        Return Application.StartupPath & Slash(Application.StartupPath) & FileName
    End Function

    Private Function GetLangString(ByVal FileName As String, ByVal KeyName As String, ByVal sDefault As String) As String
        Dim n As Int32
        Dim sData As String
        sData = Space$(1024)
        n = GetPrivateProfileString("LANGUAGE", KeyName, sDefault, sData, sData.Length, FileName)
        If n > 0 Then
            Return sData.Substring(0, n).Replace("|newline|", vbCrLf)
        Else
            Return vbNullString
        End If
    End Function

    Public Function GetInfoString(ByVal FileName As String, ByVal KeyName As String, ByVal sDefault As String) As String
        Dim n As Int32
        Dim sData As String
        sData = Space$(1024)
        n = GetPrivateProfileString("FILEINFO", KeyName, sDefault, sData, sData.Length, FileName)
        If n > 0 Then
            Return sData.Substring(0, n)
        Else
            Return vbNullString
        End If
    End Function

    Public Sub ChangeLang(ByVal FileName As String)
        Dim fPath As String = GetLangFile(FileName)
        If Not IO.File.Exists(fPath) Then
            If Not IO.File.Exists(FileName) Then
                Exit Sub
            Else
                fPath = FileName
            End If
        End If
        'Messages

        Error_Msg = GetLangString(fPath, "LANGUAGE", "Lỗi: |msg||newline|Xin bạn hãy thông báo lỗi này cho tác giả")
        Error_Title = GetLangString(fPath, "ERROR_TITLE", "Có lỗi trong quá trình")
        Error_InMemory = GetLangString(fPath, "ERROR_IMEM", "Không đủ bộ nhớ để thực hiện quá trình|newline|Bạn hãy chỉnh Thiết lập-Giới hạn dung lượng bộ nhớ đệm|newline|và thử lại sau")

        Timer_No_Selected = GetLangString(fPath, "SPLIT_NO_SELECTED", "Bạn chưa chọn file nào")
        Timer_File_Not_Exists = GetLangString(fPath, "SPLIT_FILE_NOT_EXISTS", "File không tồn tại")
        Timer_File_Info = GetLangString(fPath, "SPLIT_FILE_INFO", "File: |filename|  Kích thước: |filesize|")

        Btn3_Dlg_Filter = GetLangString(fPath, "SPLIT_DLG_FILTER", "All Files") & "|*.*"

        Btn1_Too_Many_Parts = GetLangString(fPath, "SPLIT_TOO_MANY_PARTS", "File sẽ được cắt thành |parts| mảnh|newline|Bạn có muốn tiếp tục?")
        Btn1_Too_Many_Parts_Title = GetLangString(fPath, "SPLIT_TOO_MANY_PARTS_TITLE", "File được cắt thành quá nhiều mảnh")
        Btn1_File_Exists = GetLangString(fPath, "SPLIT_FILE_EXISTS", "File |filename|.001 tồn tại trong thư mục xuất|newline|Bạn có muốn tiếp tục quá trình và ghi đè file này?")
        Btn1_File_Exists_Title = GetLangString(fPath, "SPLIT_FILE_EXISTS_TITLE", "Chú ý")
        Btn1_Lb7_Text_Working = GetLangString(fPath, "SPLIT_TEXT1_WORKING", "Đang cắt file (|percent|%)")
        Btn1_Lb8_Text_Working = GetLangString(fPath, "SPLIT_TEXT2_WORKING", "Đang ghi file |filename| (|prg|)")
        Btn1_Lb7_Complete = GetLangString(fPath, "SPLIT_TEXT1_COMPLETE", "Hoàn tất")
        Btn1_Complete_Msg = GetLangString(fPath, "SPLIT_COMPLETE_MSG", "File |filename| đã được cắt xong|newline|File đã cắt thành |parts| mảnh|newline|Kích thước mỗi mảnh: |chunksize||newline|Bạn có muốn mở thử mục xuất?")
        Btn1_Complete_Title = GetLangString(fPath, "SPLIT_COMPLETE_TITLE", "Thông báo")

        Btn7_Dlg_Filter = GetLangString(fPath, "MERGE_DLG_FILTER", "All Files") & "|*.*"
        Btn7_Auto_Dlg_Filter = GetLangString(fPath, "MERGE_AUTO_DLG_FILTER", "First part") & "|*.001"

        Btn6_File_Exists = GetLangString(fPath, "MERGE_FILE_EXISTS", "File |filename| tồn tại trong thư mục xuất|newline|Bạn có muốn tiếp tục quá trình và ghi đè file này?")
        Btn6_File_Exists_Title = GetLangString(fPath, "MERGE_FILE_EXISTS_TITLE", "Chú ý")
        Btn6_File_Not_Exists = GetLangString(fPath, "MERGE_FILE_NOT_EXISTS", "File |filename| không tồn tại|newline|Quá trình nối file không thể tiếp tục")
        Btn6_File_Not_Exists_Title = GetLangString(fPath, "MERGE_FILE_NOT_EXISTS_TITLE", "Có lỗi trong quá trình")
        Btn6_Lb11_Text_Working = GetLangString(fPath, "MERGE_TEXT1_WORKING", "Đang nối file (|percent|%)")
        Btn6_Lb10_Text_Working = GetLangString(fPath, "MERGE_TEXT2_WORKING", "Đang ghi file |filename| (|prg|)")
        Btn6_Lb11_Complete = GetLangString(fPath, "MERGE_TEXT1_COMPLETE", "Hoàn tất")
        Btn6_Complete_Msg = GetLangString(fPath, "MERGE_COMPLETE_MSG", "File |filename| đã được nối xong|newline|Dung lượng: |filesize||newline|File đã nối từ |parts| mảnh|newline|Bạn có muốn mở file này?")
        Btn6_Complete_Title = GetLangString(fPath, "MERGE_COMPLETE_TITLE", "Thông báo")

        Btn15_File_Exists = GetLangString(fPath, "ENC_FILE_EXISTS", "File |filename| tồn tại trong thư mục xuất|newline|Bạn có muốn tiếp tục quá trình và ghi đè file này?")
        Btn15_File_Exists_Title = GetLangString(fPath, "ENC_FILE_EXISTS_TITLE", "Chú ý")
        Btn15_Enc_Text1_Working = GetLangString(fPath, "ENC_TEXT1_WORKING", "Đang mã hoá...")
        Btn15_Enc_Text2_Working = GetLangString(fPath, "ENC_TEXT2_WORKING", "Đang ghi file |filename| (|prg|)")
        Btn15_Dec_Text1_Working = GetLangString(fPath, "DEC_TEXT1_WORKING", "Đang giải mã...")
        Btn15_Dec_Text2_Working = GetLangString(fPath, "DEC_TEXT2_WORKING", "Đang ghi file |filename| (|prg|)")
        Btn15_Complete = GetLangString(fPath, "ENC_TEXT1_COMPLETE", "Hoàn tất")
        Btn15_Enc_Complete_Msg = GetLangString(fPath, "ENC_COMPLETE_MSG", "File |filename| đã được mã hoá|newline|Dung lượng: |filesize|")
        Btn15_Dec_Complete_Msg = GetLangString(fPath, "DEC_COMPLETE_MSG", "File |filename| đã được giải mã|newline|Dung lượng: |filesize||newline|Bạn có muốn mở file này?")
        Btn15_Enc_Complete_Title = GetLangString(fPath, "ENC_COMPLETE_MSG_TITLE", "Thông báo")
        Btn15_Dec_Invalid_Pass = GetLangString(fPath, "DEC_INVALID_PASS", "File |filename||newline|CRC check thất bại|newline|File đã bị hỏng hoặc mật khẩu của bạn sai.")

        'Controls
        With Form1
            .TabControl1.TabPages(0).Text = GetLangString(fPath, "TAB_SPLIT", "Cắt File")
            .TabControl1.TabPages(1).Text = GetLangString(fPath, "TAB_MERGE", "Nối file")
            .TabControl1.TabPages(2).Text = GetLangString(fPath, "TAB_ENC", "Mã hoá")
            .TabControl1.TabPages(3).Text = GetLangString(fPath, "TAB_SETTINGS", "Thiết lập")
            .TabControl1.TabPages(4).Text = GetLangString(fPath, "TAB_ABOUT", "Tác giả")
            .Label6.Text = GetLangString(fPath, "LB_PROGRESS", "Quá trình:")
            .Label12.Text = GetLangString(fPath, "LB_PROGRESS", "Quá trình:")
            .Label26.Text = GetLangString(fPath, "LB_PROGRESS", "Quá trình:")
            .Button6.Text = GetLangString(fPath, "BTN_START", "Bắt đầu")
            .Button1.Text = GetLangString(fPath, "BTN_START", "Bắt đầu")
            .Button15.Text = GetLangString(fPath, "BTN_START", "Bắt đầu")
            .Button10.Text = GetLangString(fPath, "BTN_STOP", "Dừng")
            .Button11.Text = GetLangString(fPath, "BTN_STOP", "Dừng")
            .Button14.Text = GetLangString(fPath, "BTN_STOP", "Dừng")
            .Button16.Text = GetLangString(fPath, "BTN_RESET", "Nhập lại")
            .Button17.Text = GetLangString(fPath, "BTN_RESET", "Nhập lại")
            .Button18.Text = GetLangString(fPath, "BTN_RESET", "Nhập lại")

            .Label1.Text = GetLangString(fPath, "SPLIT_LB1", "File:")
            .Label5.Text = GetLangString(fPath, "SPLIT_LB2", "Xuất ra:")
            .Label3.Text = GetLangString(fPath, "SPLIT_LB3", "Số mảnh:")
            .Label4.Text = GetLangString(fPath, "SPLIT_LB4", "mảnh")
            .RadioButton1.Text = GetLangString(fPath, "SPLIT_OPT1", "Kích thước mỗi mảnh:")
            .RadioButton2.Text = GetLangString(fPath, "SPLIT_OPT2", "Cắt thành:")
            .CheckBox1.Text = GetLangString(fPath, "SPLIT_CHK1", "Tạo file Merge.bat để nối các file")
            .CheckBox3.Text = GetLangString(fPath, "SPLIT_CHK2", "Xóa các mảnh sau khi nối")
            .CheckBox2.Text = GetLangString(fPath, "SPLIT_CHK3", "Xóa file gốc sau khi cắt")

            .Label9.Text = GetLangString(fPath, "MERGE_LB1", "Các file cần nối:")
            .Label13.Text = GetLangString(fPath, "MERGE_LB2", "Xuất ra:")
            .ListView1.Columns(0).Text = GetLangString(fPath, "MERGE_LST_COL1", "File")
            .ListView1.Columns(1).Text = GetLangString(fPath, "MERGE_LST_COL2", "Đường dẫn")
            .ListView1.Columns(2).Text = GetLangString(fPath, "MERGE_LST_COL3", "Kích thước")
            .Button7.Text = GetLangString(fPath, "MERGE_BTN1", "Thêm")
            .Button8.Text = GetLangString(fPath, "MERGE_BTN2", "Xóa")
            .CheckBox5.Text = GetLangString(fPath, "MERGE_CHK1", "Tự động thêm file")
            .CheckBox4.Text = GetLangString(fPath, "MERGE_CHK2", "Xóa các mảnh sau khi nối")
            .LinkLabel2.Text = GetLangString(fPath, "MERGE_LB3", "Tạo file Merge.bat để nối các file")

            .CheckBox6.Text = GetLangString(fPath, "SETTI_CHK1", "Giới hạn dung lượng bộ nhớ đệm")
            .Label16.Text = GetLangString(fPath, "SETTI_LB1", "Dung lượng bộ nhớ đệm:")
            .Label17.Text = GetLangString(fPath, "SETTI_LB2", "Ngôn ngữ:")
            .LnkAsso.Text = GetLangString(fPath, "SETTI_LB3", "Làm chương trình mặc định cho file *.001")

            .Label15.Text = GetLangString(fPath, "ABOUT_VERSION", "Phiên bản |version| (build |build|)|newline||newline|Tác giả: |author|").Replace( _
            "|version|", VERSION).Replace("|build|", BUILD).Replace("|author|", AUTHOR)
            .CheckValidInfo()
        End With
    End Sub
End Module
