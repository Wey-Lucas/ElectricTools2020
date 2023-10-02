'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System.Drawing

Namespace Engine2

    ''' <summary>
    ''' Permite a conversão de cores
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ConvertionColors

        ''' <summary>
        ''' Color em Brush
        ''' </summary>
        ''' <param name="Color">Color</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ColorToBrush(Color As Color) As Brush
            Return New SolidBrush(Color)
        End Function

        ''' <summary>
        ''' AcadColor em Brush
        ''' </summary>
        ''' <param name="Color">Acad Color</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AcadColorToBrush(Color As Autodesk.AutoCAD.Colors.Color) As Brush
            Return New SolidBrush(Color.ColorValue)
        End Function

        ''' <summary>
        ''' AcadColor em Color
        ''' </summary>
        ''' <param name="Color">Acad Color</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AcadColorToColor(Color As Autodesk.AutoCAD.Colors.Color) As Color
            Return Color.ColorValue
        End Function

        ''' <summary>
        ''' AcadIndexColor em AcadColor
        ''' </summary>
        ''' <param name="ColorIndex">Color</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AcadColorIndexToAcadColor(ColorIndex As Integer) As Autodesk.AutoCAD.Colors.Color
            Select Case ColorIndex
                Case -1 'ByLayer
                    Return Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 256)
                Case Else
                    Return Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, ColorIndex)
            End Select
        End Function

        ''' <summary>
        ''' Color em AcadColor
        ''' </summary>
        ''' <param name="Color">Color</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ColorToAcadColor(Color As Color) As Autodesk.AutoCAD.Colors.Color
            Return Autodesk.AutoCAD.Colors.Color.FromColor(Color)
        End Function

        ''' <summary>
        ''' Color em Brush
        ''' </summary>
        ''' <param name="Brush">Brush</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BrushToColor(Brush As Brush) As Color
            Return New Pen(Brush).Color
        End Function

        ''' <summary>
        ''' RGBString em Color
        ''' </summary>
        ''' <param name="RGBString">Texto RGB</param>
        ''' <param name="CharDivisor">Caracter de divisão</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function RGBStringToColor(RGBString As String, CharDivisor As String) As Color
            Dim RGBItems() As String = RGBString.Split(CharDivisor)
            Return System.Drawing.Color.FromArgb(CInt(RGBItems(0)), CInt(RGBItems(1)), CInt(RGBItems(2)))
        End Function

        ''' <summary>
        ''' RGBString em AcadColor
        ''' </summary>
        ''' <param name="RGBString">Texto RGB</param>
        ''' <param name="CharDivisor">Caracter de divisão</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function RGBStringToAcadColor(RGBString As String, CharDivisor As String) As Autodesk.AutoCAD.Colors.Color
            Dim RGBItems() As String = RGBString.Split(CharDivisor)
            Return Autodesk.AutoCAD.Colors.Color.FromRgb(CInt(RGBItems(0)), CInt(RGBItems(1)), CInt(RGBItems(2)))
        End Function

        ''' <summary>
        ''' RGBString em Brush
        ''' </summary>
        ''' <param name="RGBString">Texto RGB</param>
        ''' <param name="CharDivisor">Caracter de divisão</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function RGBStringToBrush(RGBString As String, CharDivisor As String) As Brush
            Dim RGBItems() As String = RGBString.Split(CharDivisor)
            Return New SolidBrush(System.Drawing.Color.FromArgb(CInt(RGBItems(0)), CInt(RGBItems(1)), CInt(RGBItems(2))))
        End Function

        ''' <summary>
        ''' Color em RGBString
        ''' </summary>
        ''' <param name="Color">Color</param>
        ''' <param name="CharDivisor">Caracter de divisão</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ColorToRGBString(Color As Color, CharDivisor As String) As String
            Return Color.R & CharDivisor & Color.G & CharDivisor & Color.B
        End Function

        ''' <summary>
        ''' Color em RGBString
        ''' </summary>
        ''' <param name="Color">Acad Color</param>
        ''' <param name="CharDivisor">Caracter de divisão</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AcadColorToRGBString(Color As Autodesk.AutoCAD.Colors.Color, CharDivisor As String) As String
            Return Color.ColorValue.R & CharDivisor & Color.ColorValue.G & CharDivisor & Color.ColorValue.B
        End Function

        ''' <summary>
        ''' Brush em RGBString
        ''' </summary>
        ''' <param name="Brush">Brush</param>
        ''' <param name="CharDivisor">Caracter de divisão</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ColorToRGBString(Brush As Brush, CharDivisor As String) As String
            Dim Color As Color = New Pen(Brush).Color
            Return Color.R & CharDivisor & Color.G & CharDivisor & Color.B
        End Function

    End Class

End Namespace