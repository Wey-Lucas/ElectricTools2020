Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Namespace Engine2

    ''' <summary>
    ''' Detecta a pressão sobre teclas durante o comando
    ''' </summary>
    ''' <remarks></remarks>
    Public Class KeyPressMonitor

        '''' <summary>
        '''' Determina se a tecla escape esta pressionada
        '''' </summary>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Shared Function PressEscape() As Boolean
        '    Return (System.Windows.Forms.Control.ModifierKeys And System.Windows.Forms.Keys.Escape) > 0
        'End Function

        '''' <summary>
        '''' Determina se a tecla Shift esta pressionada
        '''' </summary>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Shared Function PressShift() As Boolean
        '    Return (System.Windows.Forms.Control.ModifierKeys And System.Windows.Forms.Keys.Shift) > 0
        'End Function

        '''' <summary>
        '''' Determina se a tecla Control esta pressionada
        '''' </summary>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Shared Function PressControl() As Boolean
        '    Return (System.Windows.Forms.Control.ModifierKeys And System.Windows.Forms.Keys.Control) > 0
        'End Function

        ''' <summary>
        ''' Estilos de texto
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eInputType
            _Integer = 0
            _Text = 1
            _Real = 2
            _TextOnly = 3
        End Enum

        ''' <summary>
        ''' Determina se a tecla escape esta pressionada
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PressEscape() As Boolean
            Return (System.Windows.Forms.Control.ModifierKeys And System.Windows.Forms.Keys.Escape) > 0
        End Function

        ''' <summary>
        ''' Determina se a tecla Shift esta pressionada
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PressShift() As Boolean
            Return (System.Windows.Forms.Control.ModifierKeys And System.Windows.Forms.Keys.Shift) > 0
        End Function

        ''' <summary>
        ''' Determina se a tecla Control esta pressionada
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PressControl() As Boolean
            Return (System.Windows.Forms.Control.ModifierKeys And System.Windows.Forms.Keys.Control) > 0
        End Function

        ''' <summary>
        ''' Avalia a entrada de caracteres a partir do método KeyPress do controle
        ''' </summary>
        ''' <param name="InputType">Tipo de entrada</param>
        ''' <param name="sender">Controle</param>
        ''' <param name="e">Argumentos do evento</param>
        ''' <param name="Decimals">Número de casas de precisão para numeros reais, se zero ou menos for informado a precisão será equivalente a variável LUPREC do AutoCAD</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function KeyPress(InputType As eInputType, sender As Object, e As Windows.Forms.KeyPressEventArgs, Optional Decimals As Integer = 0) As Boolean
            Dim TextBox As TextBox
            Dim Previous As String = ""
            TextBox = sender
            Select Case InputType
                Case eInputType._Integer
                    If (Asc(e.KeyChar) > 47 And Asc(e.KeyChar) < 58) Or
                        Asc(e.KeyChar) = 8 Or
                        Asc(e.KeyChar) = 45 Then
                        Return False
                    Else
                        Return True
                    End If
                Case eInputType._Real
                    If Decimals <= 0 Then
                        Decimals = Application.DocumentManager.MdiActiveDocument.Database.Luprec
                    End If
                    If TextBox.Text.Contains(".") = False Then
                        If Asc(e.KeyChar) = 8 Then
                            Return False
                        Else
                            If (Asc(e.KeyChar) > 47 And Asc(e.KeyChar) < 58) Or
                                Asc(e.KeyChar) = 46 Or
                                Asc(e.KeyChar) = 45 Then
                                If Asc(e.KeyChar) = 46 Then
                                    Previous = TextBox.Text.Substring(0, TextBox.SelectionStart) & e.KeyChar & TextBox.Text.Substring(TextBox.SelectionStart)
                                    If Previous.Substring(Previous.IndexOf(".") + 1).Length <= Decimals Then
                                        Return False
                                    Else
                                        MsgBox("Apenas " & Decimals.ToString & " casa(s) de precisão pode(m) ser utilizada(s).", MsgBoxStyle.Exclamation)
                                        Return True
                                    End If
                                Else
                                    Return False
                                End If
                            Else
                                Return True
                            End If
                        End If
                    Else
                        If Asc(e.KeyChar) = 8 Then
                            Return False
                        Else
                            If (Asc(e.KeyChar) > 47 And Asc(e.KeyChar) < 58) Or
                                Asc(e.KeyChar) = 45 Then
                                If TextBox.SelectionLength >= 1 Then
                                    Return False
                                Else
                                    If TextBox.Text.ToString.Substring(TextBox.Text.ToString.IndexOf(".") + 1).Length < Decimals Then
                                        Return False
                                    Else
                                        Return True
                                    End If
                                End If
                            Else
                                Return True
                            End If
                        End If
                    End If
                Case eInputType._TextOnly
                    If ((Asc(e.KeyChar) > 64 And Asc(e.KeyChar) < 91) Or (Asc(e.KeyChar) > 96 And Asc(e.KeyChar) < 123)) Or
                         Asc(e.KeyChar) = 8 Or
                         Asc(e.KeyChar) = 45 Then
                        Return False
                    Else
                        Return True
                    End If
                Case eInputType._Text
                    Return False
            End Select
            Return False
        End Function

    End Class

End Namespace

