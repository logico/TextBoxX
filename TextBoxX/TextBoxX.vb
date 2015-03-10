' Money field code by user shavinder on CodeProject
' http://www.codeproject.com/Tips/311959/Format-a-textbox-for-currency-input-VB-NET

Option Explicit On

Imports System
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Text.RegularExpressions

Public Class TextBoxX
    Inherits System.Windows.Forms.TextBox

    Dim strCurrency As String = ""
    Dim acceptableKey As Boolean = False

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

    Private _MoneyField As Boolean = False

    ' -- Currency field
    Public Property MoneyField() As Boolean
        Get
            Return _MoneyField
        End Get
        Set(value As Boolean)
            If value = True Then
                Me.TextAlign = HorizontalAlignment.Right
                OnlyNumbers = False
            End If
            _MoneyField = value
        End Set
    End Property


    Public Property OnlyNumbers() As Boolean
        Get
            Return _onlyNumbers
        End Get
        Set(value As Boolean)
            If value = True Then
                MoneyField = False
            End If
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

        ' Devuelve true porque se analiza un campo que no es necesario
        If _IsRequired = False Then Return True

        ' Si es requerido Y no tiene texto = fallo
        If (_IsRequired = True) And (Me.Text.Length = 0) Then Return False

        ' If required and acept any string and the text is not empty = aprobed
        If (_IsRequired) And (Me._regexString = "*") And (Me.Text.Length <> 0) Then Return True


        If (Me.IsRequired) And (Me._regexString.Length <> 0) And (Me.Text.Length <> 0) Then     'If the regex string and the content of the text box are not empty

            Dim regexObj As Regex = New Regex(_regexString)
            Dim match As Match = regexObj.Match(Me.Text)

            Debug.Print(Name)
            Return match.Success                                        'Return if match

        Else
            Debug.Print(Name)
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

    ''' <summary>
    ''' Devuelve la cadena para ser convertida en decimal cuando se ingresa el campo de moneda
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetMoneyValue() As String
        If _MoneyField And Not Text.Length = 0 Then
            Return Text.Replace("$ ", vbNullString)
        Else
            Return Text
        End If
    End Function

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



    Private Sub TextBoxX_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        If (e.KeyCode >= Keys.D0 And e.KeyCode <= Keys.D9) OrElse (e.KeyCode >= Keys.NumPad0 And e.KeyCode <= Keys.NumPad9) OrElse e.KeyCode = Keys.Back OrElse e.KeyCode = Keys.Decimal OrElse e.KeyValue = 46 OrElse e.KeyCode = Keys.Oemcomma Then

            acceptableKey = True
            If e.KeyCode = 46 And _MoneyField Then
                strCurrency = ""
                Me.Clear()
            End If

            If e.KeyCode = Keys.Decimal And _MoneyField Then
                e.SuppressKeyPress = True
                SendKeys.Send(",")
            End If

        Else
            acceptableKey = False
        End If
    End Sub

    Private Sub TextBoxX_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress

        If _onlyNumbers Then
            If ((Asc(e.KeyChar)) <> 8) And (IsNumeric(e.KeyChar) = False) Then
                e.Handled = True
            End If
        End If

        If _MoneyField Then
            If acceptableKey = False Then
                ' Stop the character from being entered into the control since it is non-numerical.
                e.Handled = True
                Return
            Else

                If e.KeyChar = Convert.ToChar(Keys.Back) Then
                    If strCurrency.Length > 0 Then
                        strCurrency = strCurrency.Substring(0, strCurrency.Length - 1)
                    End If
                Else
                    strCurrency = strCurrency & e.KeyChar
                End If

                If strCurrency.Length = 0 Then
                    Me.Text = ""
                ElseIf strCurrency.Length = 1 Then
                    Me.Text = "0,0" & strCurrency
                ElseIf strCurrency.Length = 2 Then
                    Me.Text = "0," & strCurrency
                ElseIf strCurrency.Length > 2 Then
                    Me.Text = strCurrency.Substring(0, strCurrency.Length - 2) & "," & strCurrency.Substring(strCurrency.Length - 2)
                End If
                Me.Select(Me.Text.Length, 0)
            End If
            e.Handled = True
        End If

    End Sub

    Private Sub TextBoxX_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        If _MoneyField Then
            If Text.StartsWith("$ ") Then
                Text = Text.Replace("$ ", "")
                Me.SelectAll()
            End If
        End If
    End Sub

    Private Sub TextBoxX_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
        If _MoneyField And Not Text.Length = 0 Then
            Text = "$ " & Text
        End If
    End Sub
End Class
