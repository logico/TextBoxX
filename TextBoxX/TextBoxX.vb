Option Explicit On

Imports System
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Text.RegularExpressions

Public Class TextBoxX
    Inherits System.Windows.Forms.TextBox

    Private Const ECM_FIRST As Long = &H1500
    Private Const EM_SETCUEBANNER As Long = (ECM_FIRST + 1)
    Private Declare Unicode Function SendMessageW Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As String) As Int32

    Public Declare Function GetWindowDC Lib "user32" Alias "GetWindowDC" (ByVal hwnd As Integer) As Integer

    '--Border color
    Private _borderColor As Color = Color.Black
    Private Const WM_PAINT As Integer = 15

    '--Regular expression string
    Private _regexString As String = "*"

    '--Cue banner text
    Private _cuebanner As String = "Watermark"

    '--Only Numbers
    Private _onlyNumbers As Boolean = False

    Public Property OnlyNumbers() As Boolean
        Get
            Return _onlyNumbers
        End Get
        Set(value As Boolean)
            _onlyNumbers = value
        End Set
    End Property

    Public Property CueBanner() As String
        Get
            Return _cuebanner
        End Get
        Set(value As String)
            _cuebanner = value
            SetCueBanner()
        End Set
    End Property

    '--Set the cue banner string
    Private Sub SetCueBanner()
        Dim sCue As String = Me._cuebanner
        Call SendMessageW(Me.Handle, EM_SETCUEBANNER, 0&, sCue)
    End Sub

    ''' <summary>
    ''' Flag indicating if the field is requiered
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IsRequired() As Boolean = False

    'Set and get the regular expression string to validate with
    Public Property ValidationString() As String
        Get
            Return _regexString
        End Get
        Set(value As String)
            _regexString = value
        End Set
    End Property

    'Check if the text match with the pattern of the regular expression
    Public Function IsValid() As Boolean
        If (_IsRequired = True) And (Me.Text.Length = 0) Then
            Return False
        End If
        If (Me._regexString.Length <> 0) And (Me.Text.Length <> 0) Then     'If the regex string and the content of the text box are not empty

            Dim regexObj As Regex = New Regex(_regexString)
            Dim match As Match = regexObj.Match(Me.Text)

            Return match.Success                                        'Return if match
        Else
            Return False
        End If
    End Function

    Public Property BorderColor() As Color
        Get
            Return _borderColor
        End Get
        Set(ByVal Value As Color)
            Me._borderColor = Value
            Me.Refresh()
        End Set
    End Property

    Protected Overloads Overrides Sub WndProc(ByRef m As Message)
        Select Case m.Msg
            Case WM_PAINT
                MyBase.WndProc(m)
                OnPaint()
                ' break
            Case Else
                MyBase.WndProc(m)
                ' break
        End Select
    End Sub

 

    Protected Overloads Sub OnPaint()
        Dim rcItem As Rectangle = New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height)
        Dim hDC As IntPtr = GetWindowDC(Me.Handle)
        Dim gfx As Graphics = Graphics.FromHdc(hDC)
        DrawBorder(gfx, rcItem, _borderColor)
        gfx.Dispose()
    End Sub


    Private Sub DrawBorder(ByVal arGfx As Graphics, ByVal arRC As Rectangle, ByVal arcColor As Color)

        Dim lpPen As Pen = New Pen(arcColor, 1)
        Dim hDC As IntPtr = GetWindowDC(Me.Handle)
        Dim gfx As Graphics = Graphics.FromHdc(hDC)

        gfx.DrawLine(lpPen, arRC.X, arRC.Y + arRC.Height - 1, arRC.X, arRC.Y)
        gfx.DrawLine(lpPen, arRC.X, arRC.Y, arRC.X + arRC.Width - 1, arRC.Y)
        If Not (arRC.Width = 0) Then
            gfx.DrawLine(lpPen, arRC.X + arRC.Width - 1, arRC.Y, arRC.X + arRC.Width - 1, arRC.Top + arRC.Height - 1)
            gfx.DrawLine(lpPen, arRC.X + arRC.Width - 1, arRC.Top + arRC.Height - 1, arRC.X, arRC.Top + arRC.Height - 1)
            lpPen.Dispose()
        End If
    End Sub


    Private Sub TextBoxX_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        If (_onlyNumbers = True) And ((Asc(e.KeyChar)) <> 8) And (IsNumeric(e.KeyChar) = False) Then
            e.Handled = True
        End If
    End Sub
End Class
