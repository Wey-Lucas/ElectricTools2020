'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System.Windows.Forms
Imports System.Drawing
Imports Autodesk.AutoCAD.DatabaseServices
Imports ElectricTools2020.Engine2

Namespace Engine2

    ''' <summary>
    ''' Cria o ComboBox com as espessuras de linhas
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DropDownLineWeight

        'Variáveis
        'Private _ImageList As ImageList
        Private _ddlLineWeight As DropDownImage
        Private _Width As Integer
        Private _Height As Integer

        ''' <summary>
        ''' Coleção de itens
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ImageItems As ImageItems
            Get
                Return Me._ddlLineWeight.ImageItems
            End Get
        End Property

        ''' <summary>
        ''' Item selecionado
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Item As Object
            Get
                Return Me._ddlLineWeight.Item
            End Get
            Set(value As Object)
                Me._ddlLineWeight.Item = value
            End Set
        End Property

        ''' <summary>
        ''' Valor selecionado
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Value As Object
            Get
                Return Me._ddlLineWeight.Value
            End Get
            Set(value As Object)
                Me._ddlLineWeight.Value = value
            End Set
        End Property

        ''' <summary>
        ''' Posição corrente
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property SelectedIndex As Integer
            Get
                Return Me._ddlLineWeight.SelectedIndex
            End Get
            Set(value As Integer)
                Me._ddlLineWeight.SelectedIndex = value
            End Set
        End Property

        ''' <summary>
        ''' Verifica se um valor pertence a coleção
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Contains(Value As Integer) As Boolean
            Return Me._ddlLineWeight.ImageItems.Contains(Value)
        End Function

        ''' <summary>
        ''' Verifica se um item pertence a coleção
        ''' </summary>
        ''' <param name="Item"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Contains(Item As String) As Boolean
            Return Me._ddlLineWeight.ImageItems.Contains(Item)
        End Function

        ''' <summary>
        ''' Verifica se um valor pertence a coleção
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetItem(Value As Integer) As ImageItem
            Return Me._ddlLineWeight.ImageItems.GetItem(Value)
        End Function

        ''' <summary>
        ''' Verifica se um item pertence a coleção
        ''' </summary>
        ''' <param name="Item"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetItem(Item As String) As ImageItem
            Return Me._ddlLineWeight.ImageItems.GetItem(Item)
        End Function

        ''' <summary>
        ''' Obtem a imagem da linha
        ''' </summary>
        ''' <param name="PenWidth">Espessura da pena</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ToBitmap(PenWidth As Single) As Bitmap
            Try
                If PenWidth < 0 Then
                    PenWidth = 0
                End If
                Dim Bitmap As Bitmap
                Dim X As Single
                Dim Y As Single
                Dim PixelFormat As System.Drawing.Imaging.PixelFormat = Imaging.PixelFormat.Format32bppArgb
                Bitmap = New Bitmap(Me._Width, Me._Height, PixelFormat)
                X = 0
                Y = (Me._Height / 2)
                Using Graphics As Graphics = Graphics.FromImage(Bitmap)
                    With Graphics
                        .PageUnit = GraphicsUnit.Pixel
                        .CompositingMode = Drawing2D.CompositingMode.SourceOver
                        .CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                        .InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                        .TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
                        .FillRegion(Brushes.White, New System.Drawing.Region(New RectangleF(0, 0, Me._Width, Me._Height)))
                        .DrawLine(New Pen(Brushes.Black, PenWidth), New PointF(X, Y), New PointF((X + Me._Width), Y))
                        .Save()
                    End With
                End Using
                Return Bitmap
            Catch ex As System.Exception
                Throw New System.Exception("Erro em DropDownLineWeight.ToBitmap, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="ComboBox">Controle ComboBox</param>
        ''' <param name="IncludeByLayer">Determina se o item ByLayer será incluído</param>
        ''' <param name="IncludeByBlock">Determina se o item ByBlock será incluído</param>
        ''' <param name="IncludeDefault">Determina se o item Default será incluído</param>
        ''' <remarks></remarks>
        Public Sub New(ComboBox As ComboBox, Optional IncludeByLayer As Boolean = True, Optional IncludeByBlock As Boolean = True, Optional IncludeDefault As Boolean = True)
            Dim Item As String
            Dim Value As Integer
            Dim lstPens() As String
            Dim Bitmap As Bitmap
            Dim Pen As Single
            Dim Enumerator As Type
            Me._ddlLineWeight = New DropDownImage(ComboBox, DrawMode.OwnerDrawVariable, New Size(80, 15), True, 2)
            Me._Width = 80
            Me._Height = 16
            If IncludeDefault = True Then
                Me._ddlLineWeight.ImageItems.Add("Default", -3, Me.ToBitmap(0))
            End If
            If IncludeByBlock = True Then
                Me._ddlLineWeight.ImageItems.Add("ByBlock", -2, Me.ToBitmap(0))
            End If
            If IncludeByLayer = True Then
                Me._ddlLineWeight.ImageItems.Add("ByLayer", -1, Me.ToBitmap(0))
            End If
            Enumerator = GetType(Autodesk.AutoCAD.DatabaseServices.LineWeight)
            lstPens = System.Enum.GetNames(Enumerator)
            For ini As Integer = 1 To lstPens.Length
                Item = lstPens(ini - 1)
                Value = System.Enum.GetValues(Enumerator)(ini - 1)
                If Value <> -1 And Value <> -2 And Value <> -3 Then
                    Pen = (Value / 15)
                    Bitmap = Me.ToBitmap(Pen)
                    Me._ddlLineWeight.ImageItems.Add((Value / 100).ToString("0.00"), Value, Bitmap)
                End If
            Next
            Me._ddlLineWeight.DataBind()
            Me._ddlLineWeight.SelectedIndex = -1
        End Sub

    End Class

End Namespace
