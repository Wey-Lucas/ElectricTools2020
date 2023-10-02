'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System.Windows.Forms

Namespace Engine2

    ''' <summary>
    ''' Cria uma caixa de mensagem top most
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class TopMostMessageBox

        Public Shared Function Show(Message As String, Title As String, Buttons As MessageBoxButtons, Icon As MessageBoxIcon) As DialogResult
            Dim Form As New Form()
            With Form
                .Size = New System.Drawing.Size(1, 1)
                .StartPosition = FormStartPosition.Manual
                Dim rect As System.Drawing.Rectangle = SystemInformation.VirtualScreen
                .Location = New System.Drawing.Point(rect.Bottom + 10, rect.Right + 10)
                .Show()
                .Focus()
                .BringToFront()
                .TopMost = True
                Dim result As DialogResult = MessageBox.Show(Form, Message, Title, Buttons, Icon)
                .Dispose()
                Return result
            End With
        End Function

    End Class

End Namespace
