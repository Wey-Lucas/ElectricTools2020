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
Imports adoNetExtension
Imports ElectricTools2020.Engine2

Namespace Engine2

    ''' <summary>
    ''' Cria o ComboBox com as cores
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DropDownColor

        'Variáveis
        Private _ImageList As ImageList
        Private _ddlColor As DropDownImage
        Private _Width As Integer
        Private _Height As Integer
        Private _UseAcadReverseColor As Boolean

        ''' <summary>
        ''' Coleção de itens
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ImageItems As ImageItems
            Get
                Return Me._ddlColor.ImageItems
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
                Return Me._ddlColor.Item
            End Get
            Set(value As Object)
                Me._ddlColor.Item = value
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
                Return Me._ddlColor.Value
            End Get
            Set(value As Object)
                Me._ddlColor.Value = value
            End Set
        End Property


        ''' <summary>
        ''' Retorna se a cor existe
        ''' </summary>
        ''' <param name="ColorIndex"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function ContainsColorIndex(ColorIndex As Integer) As Boolean
            If IsNothing(Me._ddlColor.ImageItems.Find(Function(s As ImageItem) s.Value = ColorIndex)) = True Then
                Return False
            Else
                Return True
            End If
        End Function

        ''' <summary>
        ''' Retorna se a cor existe
        ''' </summary>
        ''' <param name="Color"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Contains(Color As Color) As Boolean
            For Each Item As ImageItem In Me._ddlColor.ImageItems
                Dim Bitmap As Bitmap = Item.Image
                If Bitmap.GetPixel(5, 5).Equals(Color) = True Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Retorna se a cor existe
        ''' </summary>
        ''' <param name="Color"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Contains(Color As Autodesk.AutoCAD.Colors.Color) As Boolean
            For Each Item As ImageItem In Me._ddlColor.ImageItems
                Dim Bitmap As Bitmap = Item.Image
                If Bitmap.GetPixel(5, 5).Equals(Color.ColorValue) = True Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Verifica se um valor pertence a coleção
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Contains(Value As Integer) As Boolean
            Return Me._ddlColor.ImageItems.Contains(Value)
        End Function

        ''' <summary>
        ''' Verifica se um item pertence a coleção
        ''' </summary>
        ''' <param name="Item"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Contains(Item As String) As Boolean
            Return Me._ddlColor.ImageItems.Contains(Item)
        End Function

        ''' <summary>
        ''' Verifica se um valor pertence a coleção
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetItem(Value As Integer) As ImageItem
            Return Me._ddlColor.ImageItems.GetItem(Value)
        End Function

        ''' <summary>
        ''' Verifica se um item pertence a coleção
        ''' </summary>
        ''' <param name="Item"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetItem(Item As String) As ImageItem
            Return Me._ddlColor.ImageItems.GetItem(Item)
        End Function

        ''' <summary>
        ''' Cor
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Color As Color
            Get
                Dim Bitmap As Bitmap = Me._ddlColor.Image
                Return Bitmap.GetPixel(5, 5)
            End Get
            Set(value As Color)
                For Each Item As ImageItem In Me._ddlColor.ImageItems
                    Dim Bitmap As Bitmap = Item.Image
                    If Bitmap.GetPixel(5, 5).Equals(value) = True Then
                        Me._ddlColor.Value = Item.Value
                        Exit Property
                    End If
                Next
                Throw New System.Exception("A cor não pode ser encontrada")
            End Set
        End Property

        ''' <summary>
        ''' Cor (Brush)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Brush As Brush
            Get
                Return New SolidBrush(Me.Color)
            End Get
            Set(value As Brush)
                Dim Color As Color = New Pen(value).Color
                For Each Item As ImageItem In Me._ddlColor.ImageItems
                    Dim Bitmap As Bitmap = Item.Image
                    If Bitmap.GetPixel(5, 5).Equals(Color) = True Then
                        Me._ddlColor.Value = Item.Value
                        Exit Property
                    End If
                Next
                Throw New System.Exception("A cor não pode ser encontrada")
            End Set
        End Property

        ''' <summary>
        ''' Número da cor
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ColorIndex As Integer
            Get
                Return Me._ddlColor.Value
            End Get
            Set(value As Integer)
                Me._ddlColor.Value = value
            End Set
        End Property

        ''' <summary>
        ''' Cor RGB
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RGBColor As String
            Get
                Dim Bitmap As Bitmap = Me._ddlColor.Image
                Return Engine2.ConvertionColors.ColorToRGBString(Bitmap.GetPixel(5, 5), ";")
            End Get
            Set(value As String)
                Dim RGBItems() As String = value.Split(";")
                Dim Color As Color = System.Drawing.Color.FromArgb(CInt(RGBItems(0)), CInt(RGBItems(1)), CInt(RGBItems(2)))
                For Each Item As ImageItem In Me._ddlColor.ImageItems
                    Dim Bitmap As Bitmap = Item.Image
                    If Bitmap.GetPixel(5, 5).Equals(Color) = True Then
                        Me._ddlColor.Value = Item.Value
                        Exit Property
                    End If
                Next
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
                Return Me._ddlColor.SelectedIndex
            End Get
            Set(value As Integer)
                Me._ddlColor.SelectedIndex = value
            End Set
        End Property

        ''' <summary>
        ''' Adiciona uma nova cor
        ''' </summary>
        ''' <param name="RGBColor"></param>
        ''' <remarks></remarks>
        Public Sub AddColor(RGBColor As String)
            Dim Item As String
            Dim Value As Integer
            Dim Bitmap As Bitmap
            Item = RGBColor
            Value = (Me.ImageItems.Count + 1)
            Bitmap = Me.ToBitmap(RGBColor)
            Me._ddlColor.ImageItems.Add(Item, Value, Bitmap)
            Me._ddlColor.DataBind()
            Me._ddlColor.Value = Value
        End Sub

        ''' <summary>
        ''' Obtem a imagem da linha
        ''' </summary>
        ''' <param name="ColorIndexOrRGB">RGB ou Index da cor</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Overloads Function ToBitmap(ColorIndexOrRGB As Object) As Bitmap
            Try
                Dim Brush As Brush
                Dim Bitmap As Bitmap
                Dim PixelFormat As System.Drawing.Imaging.PixelFormat
                Dim CadColor As Autodesk.AutoCAD.Colors.Color
                If IsNumeric(ColorIndexOrRGB) = True Then
                    If CInt(ColorIndexOrRGB) >= -1 And CInt(ColorIndexOrRGB) <= 256 Then
                        If ColorIndexOrRGB > -1 Then
                            If Me._UseAcadReverseColor = True Then
                                CadColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, ColorIndexOrRGB)
                                Brush = New SolidBrush(System.Drawing.Color.FromArgb(CadColor.ColorValue.R, CadColor.ColorValue.G, CadColor.ColorValue.B))
                            Else
                                If ColorIndexOrRGB = 7 Then
                                    Brush = New SolidBrush(System.Drawing.Color.FromArgb(255, 255, 255))
                                Else
                                    CadColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByColor, ColorIndexOrRGB)
                                    Brush = New SolidBrush(System.Drawing.Color.FromArgb(CadColor.ColorValue.R, CadColor.ColorValue.G, CadColor.ColorValue.B))
                                End If
                            End If
                        Else
                            Brush = New SolidBrush(System.Drawing.Color.FromArgb(255, 255, 255))
                        End If
                    Else
                        Throw New System.Exception("O número da cor não pode ser encontrado")
                    End If
                Else
                    If ColorIndexOrRGB.ToString.Split(";").Length = 3 Then
                        Brush = Engine2.ConvertionColors.RGBStringToBrush(ColorIndexOrRGB, ";")
                    Else
                        Throw New System.Exception("O RGB da cor não é válido")
                    End If
                End If
                PixelFormat = Imaging.PixelFormat.Format32bppArgb
                Bitmap = New Bitmap(Me._Width, Me._Height, PixelFormat)
                Using Graphics As Graphics = Graphics.FromImage(Bitmap)
                    With Graphics
                        .PageUnit = GraphicsUnit.Pixel
                        .CompositingMode = Drawing2D.CompositingMode.SourceOver
                        .CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                        .InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                        .FillRegion(Brush, New System.Drawing.Region(New RectangleF(0, 0, Me._Width, Me._Height)))
                        .DrawRectangle(New Pen(Brushes.Black, 1), New Rectangle(0, 0, Me._Width - 1, Me._Height - 1))
                        .Save()
                    End With
                End Using
                Return Bitmap
            Catch ex As System.Exception
                Throw New System.Exception("Erro em DropDownColor.ToBitmap, motivo: " & ex.Message)
            End Try
        End Function

        ''ARRUMOU AS CORES INDEX NA TABELA DO SISTEMA
        'Public Sub UpdateDataBase()
        '    Dim DT As System.Data.DataTable
        '    Using DB As New AdoNetConnect.AdoNet
        '        With DB
        '            .OpenConnection(Commands.MyPlugin.DbConnect)
        '            With DB.SqlSelect
        '                .Clear()
        '                .AppendLine("SELECT")
        '                .AppendLine("LLC.llcID")
        '                .AppendLine(", LLC.llcRGBBackColor")
        '                .AppendLine("FROM")
        '                .AppendLine("exlnLinLibraryConfig as LLC")
        '                DT = .CreateConsult()
        '            End With
        '            For Index As Integer = 0 To DT.Rows.Count - 1
        '                If IsDBNull(DT.Rows.Item(Index).Item("llcRGBBackColor")) = False Then
        '                    Dim Bitmap As Bitmap = Me.ToBitmap(DT.Rows.Item(Index).Item("llcRGBBackColor"))
        '                    Dim Color As Color = Bitmap.GetPixel(3, 3)
        '                    Dim CadColor As Autodesk.AutoCAD.Colors.Color = Autodesk.AutoCAD.Colors.Color.FromColor(Color)
        '                    With DB.Fields
        '                        .ClearAll()
        '                        .Add("llcBackColorIndex", CadColor.ColorIndex)
        '                        With .Conditional
        '                            .Add("llcID", "=", DT.Rows.Item(Index).Item("llcID"))
        '                        End With
        '                        .UpdateRegistry("exlnLinLibraryConfig", True)
        '                    End With
        '                Else
        '                    With DB.Fields
        '                        .ClearAll()
        '                        .Add("llcBackColorIndex", -1)
        '                        With .Conditional
        '                            .Add("llcID", "=", DT.Rows.Item(Index).Item("llcID"))
        '                        End With
        '                        .UpdateRegistry("exlnLinLibraryConfig", True)
        '                    End With
        '                End If
        '            Next
        '            .CloseConnection()
        '        End With
        '    End Using
        'End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="ComboBox">Controle ComboBox</param>
        ''' <param name="IndexColors">Coleção de cores a serem exibidas</param>
        ''' <param name="UseAcadReverseColor">Define se o controle irá modificar a cor 7 de acordo com a cor de fundo do AutoCAD</param>
        ''' <remarks></remarks>
        Public Sub New(ComboBox As ComboBox, Optional IndexColors As ArrayList = Nothing, Optional UseAcadReverseColor As Boolean = False)
            Me._UseAcadReverseColor = UseAcadReverseColor
            Dim Item As String
            Dim Value As Integer
            Dim Bitmap As Bitmap
            Me._Width = 15
            Me._Height = 15
            Me._ddlColor = New DropDownImage(ComboBox, DrawMode.OwnerDrawVariable, New Size(15, 15), True, 2)
            If IsNothing(IndexColors) = True Then
                IndexColors = New ArrayList
                For ini As Integer = -1 To 255
                    IndexColors.Add(ini)
                Next
            End If
            For Each Ini As Short In IndexColors
                Select Case Ini
                    Case -1
                        Item = "ByLayer"
                    Case 0
                        Item = "ByBlock"
                    Case 1
                        Item = "Red"
                    Case 2
                        Item = "Yellow"
                    Case 3
                        Item = "Green"
                    Case 4
                        Item = "Cyan"
                    Case 5
                        Item = "Blue"
                    Case 6
                        Item = "Magenta"
                    Case 7
                        If Me._UseAcadReverseColor = True Then
                            Item = "White\Black"
                        Else
                            Item = "White"
                        End If
                    Case 250
                        Item = "Black"
                    Case Else
                        Item = CStr(Ini)
                End Select
                Value = Ini
                Bitmap = Me.ToBitmap(Ini)
                Me._ddlColor.ImageItems.Add(Item, Value, Bitmap)
            Next
            Me._ddlColor.DataBind()
            Me._ddlColor.SelectedIndex = -1
        End Sub

    End Class

End Namespace
