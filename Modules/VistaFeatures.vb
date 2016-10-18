Imports System.Runtime.InteropServices

Module VistaFeatures
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MARGINS
        Public cxLeftWidth As Integer
        Public cxRightWidth As Integer
        Public cyTopHeight As Integer
        Public cyBottomheight As Integer
    End Structure

    <DllImport("dwmapi.dll")> _
    Public Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As MARGINS) As Integer
    End Function

    <DllImport("dwmapi.dll")> _
    Public Function DwmIsCompositionEnabled(<MarshalAs(UnmanagedType.Bool)> ByRef pfEnabled As Boolean) As Integer
    End Function

    Public Function IsGlassEnabled() As Boolean
        If Environment.OSVersion.Version.Major < 6 Then
            Return False 'Not Vista.
        End If

        Dim isGlassSupported As Boolean = False
        DwmIsCompositionEnabled(isGlassSupported)
        Return isGlassSupported
    End Function
End Module
