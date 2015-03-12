﻿Option Explicit On

Imports System
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Text.RegularExpressions

Public Class TextBoxX
    Inherits System.Windows.Forms.TextBox

    Private strCurrency As String = ""

    Private strCurrencyInt As String = "0"
    Private strCurrencyDec As String = "00"
    Private boolCurrencyDecFirstPos = True

    Private acceptableKey As Boolean = False
    Private commaEntered As Boolean = False

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
            Else
                Me.TextAlign = HorizontalAlignment.Left
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
            Dim tmp As String = Text.Replace("$ ", vbNullString)
            Return tmp.Replace(".", vbNullString)
        Else
            Return Text
        End If
    End Function

    Protected Overloads Overrides Sub WndProc(ByRef m As Message)
        Select Case m.Msg
            Case WM_PAINT
                MyBase.WndProc(m)
                OnPaint()
            Case Else
                MyBase.WndProc(m)
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

        If (e.KeyCode >= Keys.D0 And e.KeyCode <= Keys.D9) _
            OrElse (e.KeyCode >= Keys.NumPad0 And e.KeyCode <= Keys.NumPad9) _
            OrElse e.KeyCode = Keys.Decimal _
            OrElse e.KeyValue = 46 _
            OrElse e.KeyCode = Keys.Oemcomma Then

            acceptableKey = True
            If e.KeyCode = 46 And _MoneyField Then
                Me.Clear()
                commaEntered = False
                strCurrencyInt = "0"
                strCurrencyDec = "00"
                boolCurrencyDecFirstPos = True
                Text = "00,00"
            End If

            If e.KeyCode = Keys.Decimal And _MoneyField Then
                commaEntered = True
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

                ' If a number is entered
                If IsNumeric(e.KeyChar) Then
                    ' and the comma is not entered
                    If commaEntered = False Then
                        ' and it's the first digit
                        If strCurrencyInt.Length = 1 And strCurrencyInt = "0" Then
                            ' Add only one digit
                            strCurrencyInt = e.KeyChar
                        Else
                            ' else, put the char at the end of the string
                            strCurrencyInt = strCurrencyInt & e.KeyChar
                        End If
                    Else
                        ' The comma is entered
                        If boolCurrencyDecFirstPos Then
                            strCurrencyDec = e.KeyChar & strCurrencyDec.Substring(1, 1)
                            boolCurrencyDecFirstPos = False
                        Else
                            strCurrencyDec = strCurrencyDec.Substring(0, 1) & e.KeyChar
                            boolCurrencyDecFirstPos = True
                        End If

                    End If
                Else

                    ' The only key accepted here is the comma
                    If e.KeyChar = "," Then
                        commaEntered = True
                    End If
                End If
                e.Handled = True
                Me.Text = strCurrencyInt & "," & strCurrencyDec
                ' FIN DE LA RAMA aceptableKey
            End If
        End If
    End Sub

    Private Sub TextBoxX_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        If _MoneyField Then
            If Text.StartsWith("$ ") Then
                Text = Text.Replace("$ ", "")
            End If
        End If
    End Sub

    Private Sub TextBoxX_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
        If _MoneyField And Not Text.Length = 0 Then
            Text = String.Format("{0:C2}", CDec(Text))
        End If
    End Sub

    Private Sub TextBoxX_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        SelectAll()
    End Sub
End Class
