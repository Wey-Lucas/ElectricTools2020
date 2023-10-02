Imports System.Drawing

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='
Namespace Engine2

    ''' <summary>
    ''' Classe para medição de strings
    ''' </summary>
    ''' <remarks></remarks>
    Public Class MeasureString

        ''' <summary>
        ''' Mede as dimensões de uma string
        ''' </summary>
        ''' <param name="Font">Fonte</param>
        ''' <param name="Text">Texto a ser medido</param>
        ''' <param name="Graphics">Grafico</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetMeasure(Font As System.Drawing.Font, Text As String, Optional Graphics As Graphics = Nothing) As SizeF
            Dim Form As New System.Windows.Forms.Form
            If IsNothing(Graphics) = True Then
                Graphics = Form.CreateGraphics
            End If
            Return Graphics.MeasureString(Text, Font)
        End Function

    End Class

End Namespace


