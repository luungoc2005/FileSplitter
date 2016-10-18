Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Drawing.Drawing2D
Imports System.Drawing.Text

Public Class Splitter_Button
    Private Btn_Rectangle As Rectangle
    Private Btn_Shape As GraphicsPath
    Private Btn_Graphic As Graphics

    Private _InitialColor As Color

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _InitialColor = Color.Red
        Btn_Rectangle = New Rectangle
        Btn_Shape = New GraphicsPath
        Btn_Graphic = Graphics.FromHwnd(Me.Handle)
        CreateShape()
        DrawStroke()
    End Sub

#Region "Paint button"
    Private Function CreateRoundRectangle(ByVal rectangle As Rectangle, ByVal radius As Integer) As GraphicsPath
        Dim path As New GraphicsPath()
        Dim l As Integer = rectangle.Left
        Dim t As Integer = rectangle.Top
        Dim w As Integer = rectangle.Width
        Dim h As Integer = rectangle.Height
        Dim d As Integer = radius << 1
        path.AddArc(l, t, d, d, 180, 90) ' topleft
        path.AddLine(l + radius, t, l + w - radius, t) ' top
        path.AddArc(l + w - d, t, d, d, 270, 90) ' topright
        path.AddLine(l + w, t + radius, l + w, t + h - radius) ' right
        path.AddArc(l + w - d, t + h - d, d, d, 0, 90) ' bottomright
        path.AddLine(l + w - radius, t + h, l + radius, t + h) ' bottom
        path.AddArc(l, t + h - d, d, d, 90, 90) ' bottomleft
        path.AddLine(l, t + h - radius, l, t + radius) ' left
        path.CloseFigure()
        Return path
    End Function

    Private Sub CreateShape()
        Btn_Rectangle = New Rectangle(0, 0, Me.Width, Me.Height)
        Btn_Shape = CreateRoundRectangle(Btn_Rectangle, 5)
    End Sub

    Private Sub DrawStroke()
        Dim _Pen As New Pen(_InitialColor, 1)
        Btn_Graphic.DrawPath(_Pen, Btn_Shape)
        _Pen.Dispose()
    End Sub
#End Region

    Private Sub Splitter_Button_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        '_InitialColor = Color.Red
        'CreateShape()
        DrawStroke()
    End Sub

    Private Sub Splitter_Button_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        '_InitialColor = Color.Red
        'CreateShape()
        DrawStroke()
    End Sub
End Class
