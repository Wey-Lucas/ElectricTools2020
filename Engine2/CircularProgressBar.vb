Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Imaging

Namespace Engine2

    ''' <summary>
    ''' Gera a barra de progresso circular
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CircularProgressBar

        ''' <summary>
        ''' Evento que informa a posição atual da barra de progresso
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Event CircularProgressBarValue(sender As Object, e As Integer)

        ''' <summary>
        ''' Controle base
        ''' </summary>
        ''' <remarks></remarks>
        Private _PictureBox As PictureBox

        ''' <summary>
        ''' Valor 
        ''' </summary>
        ''' <remarks></remarks>
        Private _Value As Integer

        ''' <summary>
        ''' Estilo
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum eStyle

            ''' <summary>
            ''' Nenhum
            ''' </summary>
            ''' <remarks></remarks>
            ValueOnly = 0

            ''' <summary>
            ''' Percentual %
            ''' </summary>
            ''' <remarks></remarks>
            Percentual = 1

            ''' <summary>
            ''' Graus °
            ''' </summary>
            ''' <remarks></remarks>
            Degree = 2

        End Enum

        ''' <summary>
        ''' Estilo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Style As eStyle

        ''' <summary>
        ''' Valor
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Value As Integer
            Get
                Return Me._Value
            End Get
            Set(value As Integer)
                RaiseEvent CircularProgressBarValue(Me, value)
                Me._Value = value
                Me.SetValue(Me._Value)
            End Set
        End Property

        ''' <summary>
        ''' Espessura da pena
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property PenWidth As Single

        ''' <summary>
        ''' Cor do texto
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property TextColor As Color

        ''' <summary>
        ''' Cor da barra de progresso
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property BarColor As Color

        ''' <summary>
        ''' Cor do fundo da barra de progresso
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property BarBackGroundColor As Color

        ''' <summary>
        ''' Tamanho da imagem do controle
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Property Size As Integer

        ''' <summary>
        ''' Atualiza o valor
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Private Sub SetValue(Value As Integer)
            Dim PixelFormat As PixelFormat = Imaging.PixelFormat.Format32bppArgb
            Dim Bitmap As Bitmap = New Bitmap(Me.Size, Me.Size, PixelFormat)
            Using Graphics As Graphics = Graphics.FromImage(Bitmap)
                With Graphics
                    .PageUnit = GraphicsUnit.Pixel
                    .CompositingMode = Drawing2D.CompositingMode.SourceOver
                    .CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                    .InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    .SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias
                    .FillRegion(New SolidBrush(Me._PictureBox.BackColor), New System.Drawing.Region(New RectangleF(0, 0, Me.Size, Me.Size)))
                    Dim Rect As Rectangle = New Rectangle(New Point((PenWidth / 2) + 1, (PenWidth / 2) + 1), New Size(Me.Size - (PenWidth + 1), Me.Size - (PenWidth + 1)))
                    Dim Angle As Integer
                    Dim Text As String = ""
                    Dim TextSize As SizeF
                    Dim TextPoint As Point
                    Using ProgressPen As New Pen(Me.BarColor, PenWidth), RemainderPen As New Pen(Me.BarBackGroundColor, PenWidth)
                        Select Case Me.Style
                            Case eStyle.ValueOnly, eStyle.Percentual
                                Angle = CInt(360 / 100 * Value)
                                .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                                .DrawEllipse(RemainderPen, Rect)
                                .DrawArc(ProgressPen, Rect, -90, Angle)
                            Case eStyle.Degree
                                Value = (Value + Math.Ceiling(-Value / 360) * 360)
                                Angle = Value * -1
                                .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                                .DrawEllipse(RemainderPen, Rect)
                                .DrawArc(ProgressPen, Rect, 0, Angle)
                        End Select
                    End Using
                    Using Font As New Font(Me._PictureBox.Font.FontFamily, 13)
                        Select Case Me.Style
                            Case eStyle.ValueOnly
                                Text = Value.ToString
                            Case eStyle.Percentual
                                Text = Value.ToString & "%"
                            Case eStyle.Degree
                                Text = Value.ToString & "°"
                        End Select
                        TextSize = .MeasureString(Text, Font)
                        TextPoint = New Point(CInt(Rect.Left + (Rect.Width / 2) - (TextSize.Width / 2)), CInt(Rect.Top + (Rect.Height / 2) - (TextSize.Height / 2)))
                        .DrawString(Text, Font, New SolidBrush(Me.TextColor), TextPoint)
                    End Using
                    .Save()
                End With
            End Using
            Me._PictureBox.Image = Bitmap
            Application.DoEvents()
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New(PictureBox As PictureBox)
            Me._PictureBox = PictureBox
            Me.Size = 50
            Me.Style = eStyle.Degree
            Me.PenWidth = 4
            Me.TextColor = Color.Black
            Me.BarColor = Color.Blue
            Me.BarBackGroundColor = Color.LightGray
            Me._Value = 0
        End Sub

    End Class

End Namespace
