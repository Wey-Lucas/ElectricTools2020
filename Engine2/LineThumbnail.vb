'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Drawing.Drawing2D
Imports System.Drawing
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.IO
Imports System.Windows.Forms

Namespace Engine2

    ''' <summary>
    ''' Gera o preview a partir do código da linha
    ''' </summary>
    ''' <remarks></remarks>
    Public Class LineThumbnail

        'Variáveis
        Private _WidthSegment As Single
        Private _TextParts As TextParts
        Private _LineCode As String
        Private _PenWidth As Single
        Private _PenColor As Brush
        Private _BackColor As Brush
        Private _Items() As String
        Private _Width As Single
        Private _Height As Single
        Private _Scale As Single
        Private _Segments As Integer

        ''' <summary>
        ''' Obtem a imagem da linha
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ToBitmap() As Bitmap
            If Me._Scale <> (Me._Width * -1) Then
                Try
                    Dim Bitmap As Bitmap
                    Dim X As Single
                    Dim Y As Single
                    Dim Width As Integer
                    Dim Height As Integer
                    Dim Value As Object
                    Dim PixelFormat As System.Drawing.Imaging.PixelFormat = Imaging.PixelFormat.Format32bppArgb
                    Dim StringFormat As New StringFormat
                    With StringFormat
                        .Alignment = StringAlignment.Near
                        .LineAlignment = StringAlignment.Far
                        .FormatFlags = StringFormatFlags.NoWrap
                        .Trimming = StringTrimming.None
                    End With
                    Width = Engine2.ConvertionUnits.MmToPx(Me._Width, ConvertionUnits.DpiPos.X)
                    Height = Engine2.ConvertionUnits.MmToPx(Me._Height, ConvertionUnits.DpiPos.Y)
                    Bitmap = New Bitmap(Width, Height, PixelFormat)
                    X = 0
                    Y = (Me._Height / 2)
                    Using Graphics As Graphics = Graphics.FromImage(Bitmap)
                        With Graphics
                            .PageUnit = GraphicsUnit.Millimeter
                            .CompositingMode = CompositingMode.SourceOver
                            .CompositingQuality = CompositingQuality.HighQuality
                            .InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                            .TextRenderingHint = Text.TextRenderingHint.AntiAlias
                            .FillRegion(Me._BackColor, New System.Drawing.Region(New RectangleF(0, 0, Width, Height)))
                            If Me._LineCode.Trim <> "" Then
                                While X < Me._Width
                                    For Each Item As String In Me._Items
                                        If Item.ToUpper <> "A" Then
                                            If Item.Contains(Chr(11)) = False Then
                                                Value = (CSng(Item) * Me._Scale)
                                                If Value = 0 Then
                                                    .FillEllipse(Me._PenColor, New RectangleF(X, (Y - (Me._PenWidth / 2)), Me._PenWidth, Me._PenWidth))
                                                ElseIf Value > 0 Then
                                                    .DrawLine(New Pen(Me._PenColor, Me._PenWidth), New PointF(X, Y), New PointF((X + Value), Y))
                                                    X = (X + Value)
                                                ElseIf Value < 0 Then
                                                    Value = Value
                                                    X = (X + (Value * -1))
                                                End If
                                            Else
                                                Dim Text As TextPart = Me._TextParts.Item(Item.Substring(0, 1))
                                                Using Matrix = New Matrix
                                                    Matrix.RotateAt((Text.Angle * -1), New PointF(X + Text.Point.X, Y + Text.Point.Y), MatrixOrder.Append)
                                                    .ResetTransform()
                                                    .Transform = Matrix
                                                    .DrawString(Text.Text, New System.Drawing.Font(Text.Font.FontFamily, Text.Font.Size, FontStyle.Regular, GraphicsUnit.Millimeter, 150, 150), Me._PenColor, X + Text.Point.X, (Y + (.MeasureString(Text.Text, Text.Font, Width, StringFormat).ToSize.Height - Text.Font.Size)) + Text.Point.Y, StringFormat)
                                                    .ResetTransform()
                                                End Using
                                            End If
                                        End If
                                    Next
                                End While
                            Else
                                .DrawLine(New Pen(Me._PenColor, Me._PenWidth), New PointF(X, Y), New PointF(Me._Width, Y))
                            End If
                            .Save()
                        End With
                    End Using
                    Return Bitmap
                Catch
                    Return Nothing
                End Try
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Obtem a escala
        ''' </summary>
        ''' <param name="LineCode">Cpodigo da linha</param>
        ''' <param name="Width">Comprimento da imagem</param>
        ''' <returns>Single</returns>
        ''' <remarks></remarks>
        Private Function GetScale(ByVal LineCode As String, Width As Single) As Single
            Try
                If LineCode.Trim <> "" Then
                    Dim Texts As New TextParts(LineCode, 1)
                    For Index As Integer = 0 To Texts.Count - 1
                        LineCode = LineCode.Replace(Texts.Item(Index).Code, Index & Chr(11))
                    Next
                    Dim Items() As String = LineCode.Split(",")
                    Dim Sum As Single = 0
                    For Each Item As Object In Items
                        If Item.ToString.ToUpper <> "A" And Item <> "" Then
                            If Item.ToString.Contains(Chr(11)) = False Then
                                If CSng(Item) < 0 Then
                                    Item = (Item * -1)
                                    Sum = Sum + CSng(Item)
                                ElseIf CSng(Item) > 0 Then
                                    Sum = Sum + CSng(Item)
                                End If
                            End If
                        End If
                    Next
                    If Width < Sum Then
                        Return (Width / Sum)
                    Else
                        If Sum <> 0 Then
                            Return ((Width / Sum) / Me._Segments)
                        Else
                            Return (Width * -1)
                        End If
                    End If
                Else
                    Return 1
                End If
            Catch
                Return (Width * -1)
            End Try
        End Function

        ''' <summary>
        ''' Obtem as partes que são texto
        ''' </summary>
        ''' <param name="LineCode">Código da linha</param>
        ''' <returns>ArrayList</returns>
        ''' <remarks></remarks>
        Private Function GetTextParts(LineCode As String) As ArrayList
            Try
                GetTextParts = New ArrayList
                Dim Regex As RegularExpressions.Regex = New RegularExpressions.Regex("\[([^\]]*)\]")
                If Regex.IsMatch(LineCode) = True Then
                    For Each Match As Match In Regex.Matches(LineCode)
                        GetTextParts.Add(Match.Value)
                    Next
                End If
                Return GetTextParts
            Catch ex As System.Exception
                Throw New System.Exception("Erro em LineThumbnail.GetTextPart, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="LineCode">Código da linha</param>
        ''' <param name="PenWidth">Espessura da linha</param>
        ''' <param name="PenColor">Cor da linha</param>
        ''' <param name="BackColor">Cor de fundo</param>
        ''' <param name="Width">Comprimento da imagem</param>
        ''' <param name="Height">Altura da imagem</param>
        ''' <param name="Segments">Número de segmentos a serem exibidos</param>
        ''' <remarks></remarks>
        Public Sub New(LineCode As String, PenWidth As Single, PenColor As Brush, BackColor As Brush, Width As Single, Height As Single, Optional Segments As Integer = 6)
            If LineCode.Trim <> "" Then
                Me._Segments = Segments
                Me._Scale = Me.GetScale(LineCode, Width)
                Me._Width = Width
                Me._Height = Height
                Me._LineCode = LineCode
                Me._PenWidth = PenWidth
                Me._PenColor = PenColor
                Me._BackColor = BackColor
                If Me._LineCode.Substring(0, 1).ToUpper <> "A" Then
                    Throw New System.Exception("O código gerador da linha não é válido." & vbCrLf & "O marcador de início 'A' não foi informado.")
                End If
                Me._TextParts = New TextParts(Me._LineCode, Me._Scale, True)
                For Index As Integer = 0 To Me._TextParts.Count - 1
                    Me._LineCode = Me._LineCode.Replace(Me._TextParts.Item(Index).Code, Index & Chr(11))
                Next
                Me._Items = Me._LineCode.Split(",")
            Else
                Me._Segments = Segments
                Me._Scale = Me.GetScale(LineCode, Width)
                Me._Width = Width
                Me._Height = Height
                Me._LineCode = LineCode
                Me._PenWidth = PenWidth
                Me._PenColor = PenColor
                Me._BackColor = BackColor
            End If
        End Sub

    End Class

    ''' <summary>
    ''' Gerencia os textos da linha
    ''' </summary>
    ''' <remarks></remarks>
    Friend NotInheritable Class TextParts : Inherits List(Of TextPart)

        ''' <summary>
        ''' Obtem as partes do código referente a texto
        ''' </summary>
        ''' <param name="LineCode">Código da linha</param>
        ''' <returns>ArrayList</returns>
        ''' <remarks></remarks>
        Private Function GetTextParts(LineCode As String) As ArrayList
            Try
                GetTextParts = New ArrayList
                Dim Regex As RegularExpressions.Regex = New RegularExpressions.Regex("\[([^\]]*)\]")
                If Regex.IsMatch(LineCode) = True Then
                    For Each Match As Match In Regex.Matches(LineCode)
                        GetTextParts.Add(Match.Value)
                    Next
                End If
                Return GetTextParts
            Catch ex As System.Exception
                Throw New System.Exception("Erro em TextParts.GetTextPart, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="LineCode">Código da linha</param>
        ''' <param name="Scale">Fator de escala</param>
        ''' <param name="ReverseWCS">Determina se os valores negativos devem ter sinal inverso em função da WCS</param>
        ''' <remarks></remarks>
        Public Sub New(LineCode As String, Scale As Single, Optional ReverseWCS As Boolean = True)
            Dim TextCodes As ArrayList = Me.GetTextParts(LineCode)
            For Each TextCode As String In TextCodes
                MyBase.Add(New TextPart(TextCode, Scale, ReverseWCS))
            Next
        End Sub

    End Class

    ''' <summary>
    ''' Texto da linha item da classe TextParts
    ''' </summary>
    ''' <remarks></remarks>
    Friend NotInheritable Class TextPart

        'Variáveis
        Private _Items() As String
        Private _Angle As Single
        Private _Font As System.Drawing.Font
        Private _Point As PointF
        Private _Code As String
        Private _Text As String
        Private _FontName As String
        Private _FontSize As Single
        Private _X As Single
        Private _Y As Single
        Private _Fonts As ArrayList
        Private _TextStyleTableRecord As TextStyleTableRecord

        ''' <summary>
        ''' Estilo de texto no AutoCAD
        ''' </summary>
        ''' <value></value>
        ''' <returns>TextStyleTableRecord</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property TextStyleTableRecord As TextStyleTableRecord
            Get
                Return Me._TextStyleTableRecord
            End Get
        End Property

        ''' <summary>
        ''' Código da linha
        ''' </summary>
        ''' <value></value>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Code As String
            Get
                Return Me._Code
            End Get
        End Property

        ''' <summary>
        ''' Fonte
        ''' </summary>
        ''' <value></value>
        ''' <returns>Font</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Font As System.Drawing.Font
            Get
                Return Me._Font
            End Get
        End Property

        ''' <summary>
        ''' Ponto
        ''' </summary>
        ''' <value></value>
        ''' <returns>PointF</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Point As PointF
            Get
                Return Me._Point
            End Get
        End Property

        ''' <summary>
        ''' Texto
        ''' </summary>
        ''' <value></value>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Text As String
            Get
                Return Me._Text
            End Get
        End Property

        ''' <summary>
        ''' Ângulo de rotação do texto
        ''' </summary>
        ''' <value></value>
        ''' <returns>Single</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Angle As Single
            Get
                Return Me._Angle
            End Get
        End Property

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="TextCode">Código do texto</param>
        ''' <param name="Scale">Fator de escala</param>
        ''' <param name="ReverseWCS">Determina se os valores negativos devem ter sinal inverso em função da WCS</param>
        ''' <remarks></remarks>
        Public Sub New(TextCode As String, Scale As Single, Optional ReverseWCS As Boolean = True)
            Me._Code = TextCode
            Dim Code As String = TextCode.Substring(1, TextCode.Length - 2)
            Me._Items = Code.Split(",")
            For Each Item As Object In Me._Items
                If Item.Contains(Chr(34)) = True Then
                    Me._Text = Item.Replace("""", "")
                ElseIf Item.Contains("=") = False Then
                    Me._TextStyleTableRecord = Engine2.TextStyle.GetTextStyle(Item)
                    If IsNothing(Me._TextStyleTableRecord) = False Then
                        Me._FontName = Me._TextStyleTableRecord.Font.TypeFace
                    Else
                        If Engine2.TextStyle.Create(Item, Item, 0) = True Then
                            Me._FontName = Item
                        Else
                            Throw New System.Exception("A fonte '" & Item & "' não pode ser encontrada")
                        End If
                    End If
                ElseIf Item.Contains("S=") = True Or Item.Contains("s=") = True Then
                    Item = Item.Replace("S=", "")
                    Item = Item.Replace("s=", "")
                    Me._FontSize = (CSng(Item) * Scale)
                ElseIf Item.Contains("R=") = True Or Item.Contains("r=") = True Then
                    Item = Item.Replace("R=", "")
                    Item = Item.Replace("r=", "")
                    Me._Angle = CSng(Item)
                ElseIf Item.Contains("X=") = True Or Item.Contains("x=") = True Then
                    Item = Item.Replace("X=", "")
                    Item = Item.Replace("x=", "")
                    Me._X = (CSng(Item) * Scale)
                ElseIf Item.Contains("Y=") = True Or Item.Contains("y=") = True Then
                    Item = Item.Replace("Y=", "")
                    Item = Item.Replace("y=", "")
                    If ReverseWCS = True Then
                        Me._Y = ((CSng(Item) * -1) * Scale)
                    Else
                        Me._Y = (CSng(Item) * Scale)
                    End If
                End If
            Next
            If Me._Text.Trim = "" Then
                Throw New System.Exception("O texto não foi informado")
            End If
            Me._Font = New System.Drawing.Font(Me._FontName, Me._FontSize, FontStyle.Regular, GraphicsUnit.Millimeter, 150, 150)
            Me._Point = New PointF(Me._X, Me._Y)
        End Sub

    End Class

End Namespace


