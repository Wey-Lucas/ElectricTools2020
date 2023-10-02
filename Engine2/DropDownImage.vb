'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System.Windows.Forms
Imports System.Drawing

Namespace Engine2

    ''' <summary>
    ''' Cria um combo com imagens
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DropDownImage

        'Declarações
        Private WithEvents _ComboBox As ComboBox
        Private _ImageItems As ImageItems
        Private _ImageSize As Size
        Private _ImageVerticalSpace As Integer

        ''' <summary>
        ''' Coleção de itens
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ImageItems As ImageItems
            Get
                Return Me._ImageItems
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
                Return Me._ComboBox.SelectedItem.Item
            End Get
            Set(value As Object)
                Me._ComboBox.SelectedIndex = -1
                For Index As Integer = 0 To Me._ComboBox.Items.Count - 1
                    Dim Item As ImageItem = Me._ComboBox.Items.Item(Index)
                    If Item.Item = value Then
                        Me._ComboBox.SelectedItem = Item
                        Exit For
                    End If
                Next
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
                Return Me._ComboBox.SelectedItem.Value
            End Get
            Set(value As Object)
                Me._ComboBox.SelectedIndex = -1
                For Index As Integer = 0 To Me._ComboBox.Items.Count - 1
                    Dim Item As ImageItem = Me._ComboBox.Items.Item(Index)
                    If Item.Value = value Then
                        Me._ComboBox.SelectedItem = Item
                        Exit For
                    End If
                Next
            End Set
        End Property

        ''' <summary>
        ''' Valor adicional para usos diversos
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property OtherValue As Object
            Get
                Return Me._ComboBox.SelectedItem.OtherValue
            End Get
            Set(value As Object)
                Me._ComboBox.SelectedIndex = -1
                For Index As Integer = 0 To Me._ComboBox.Items.Count - 1
                    Dim Item As ImageItem = Me._ComboBox.Items.Item(Index)
                    If Item.OtherValue.Equals(value) = True Then
                        Me._ComboBox.SelectedItem = Item
                        Exit For
                    End If
                Next
            End Set
        End Property

        ''' <summary>
        ''' Item selecionado
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property SelectedItem As Object
            Get
                Return Me._ComboBox.SelectedItem
            End Get
        End Property

        ''' <summary>
        ''' Imagem selecionada
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Image As Image
            Get
                Return Me._ComboBox.SelectedItem.Image
            End Get
        End Property

        ''' <summary>
        ''' Posição corrente
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property SelectedIndex As Integer
            Get
                Return Me._ComboBox.SelectedIndex
            End Get
            Set(value As Integer)
                Me._ComboBox.SelectedIndex = value
            End Set
        End Property


        ' ''' <summary>
        ' ''' Desenha imagem e texto
        ' ''' </summary>
        ' ''' <param name="sender"></param>
        ' ''' <param name="e"></param>
        ' ''' <remarks></remarks>
        'Private Sub _ComboBox_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles _ComboBox.DrawItem
        '    If e.Index <= Me._ComboBox.Items.Count - 1 And e.Index > -1 Then
        '        Dim ImageItem As ImageItem = Me._ComboBox.Items(e.Index)
        '        e.DrawBackground()
        '        e.Graphics.FillRectangle(New SolidBrush(Color.LightBlue), e.Bounds)


        '        e.Graphics.DrawImage(ImageItem.Image, New Point(e.Bounds.X, e.Bounds.Y))


        '        e.Graphics.DrawString(ImageItem.Item, Me._ComboBox.Font, Brushes.Black, New Point(ImageItem.Image.Width + 3, e.Bounds.Y))
        '        If (e.State And DrawItemState.Focus) = 0 Then
        '            e.Graphics.FillRectangle(New SolidBrush(Color.White), e.Bounds)
        '            e.Graphics.DrawImage(ImageItem.Image, New Point(e.Bounds.X, e.Bounds.Y))
        '            e.Graphics.DrawString(ImageItem.Item, Me._ComboBox.Font, Brushes.Black, New Point(ImageItem.Image.Width + 3, e.Bounds.Y))
        '        End If
        '        e.DrawFocusRectangle()
        '    End If
        'End Sub

        ' ''' <summary>
        ' ''' Ajusta o tamanho
        ' ''' </summary>
        ' ''' <param name="sender"></param>
        ' ''' <param name="e"></param>
        ' ''' <remarks></remarks>
        'Private Sub _ComboBox_MeasureItem(ByVal sender As Object, ByVal e As System.Windows.Forms.MeasureItemEventArgs) Handles _ComboBox.MeasureItem
        '    Dim ImageItem As ImageItem = Me._ComboBox.Items(e.Index)
        '    e.ItemHeight = ImageItem.Image.Height + 2
        '    e.ItemWidth = ImageItem.Image.Width
        'End Sub

        ''' <summary>
        ''' Desenha imagem e texto
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _ComboBox_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles _ComboBox.DrawItem
            Dim TextSize As SizeF
            If e.Index <= Me._ComboBox.Items.Count - 1 And e.Index > -1 Then
                Dim Item As ImageItem = Me._ComboBox.Items(e.Index)
                e.DrawBackground()
                e.Graphics.FillRectangle(New SolidBrush(Color.LightBlue), e.Bounds)
                e.Graphics.DrawImage(Item.Image, New Point(e.Bounds.X, e.Bounds.Y + (Me._ImageVerticalSpace / 2)))
                TextSize = Engine2.MeasureString.GetMeasure(Me._ComboBox.Font, Item.Item, Me._ComboBox.CreateGraphics)
                If IsNothing(Me._ImageSize) = False Then
                    e.Graphics.DrawString(Item.Item, Me._ComboBox.Font, Brushes.Black, New Point(Me._ImageSize.Width + 3, ((e.Bounds.Y + (Me._ImageSize.Height / 2)) - (TextSize.Height / 2))))
                Else
                    e.Graphics.DrawString(Item.Item, Me._ComboBox.Font, Brushes.Black, New Point(Item.Image.Width + 3, e.Bounds.Y))
                End If
                If (e.State And DrawItemState.Focus) = 0 Then
                    e.Graphics.FillRectangle(New SolidBrush(Color.White), e.Bounds)
                    If IsNothing(Me._ImageSize) = False Then
                        e.Graphics.DrawImage(Item.Image, New Point(e.Bounds.X, e.Bounds.Y + (Me._ImageVerticalSpace / 2)))
                        TextSize = Engine2.MeasureString.GetMeasure(Me._ComboBox.Font, Item.Item, Me._ComboBox.CreateGraphics)
                        e.Graphics.DrawString(Item.Item, Me._ComboBox.Font, Brushes.Black, New Point(Me._ImageSize.Width + 3, ((e.Bounds.Y + (Me._ImageSize.Height / 2)) - (TextSize.Height / 2))))
                    Else
                        e.Graphics.DrawString(Item.Item, Me._ComboBox.Font, Brushes.Black, New Point(e.Bounds.X, e.Bounds.Y))
                    End If
                End If
                e.DrawFocusRectangle()
            End If
        End Sub

        ''' <summary>
        ''' Ajusta o tamanho
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _ComboBox_MeasureItem(ByVal sender As Object, ByVal e As System.Windows.Forms.MeasureItemEventArgs) Handles _ComboBox.MeasureItem
            If IsNothing(Me._ImageSize) = False Then
                e.ItemHeight = (Me._ImageSize.Height + Me._ImageVerticalSpace)
                e.ItemWidth = Me._ImageSize.Width
            End If
        End Sub

        ''' <summary>
        ''' Carrega o controle
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub DataBind()
            With Me._ComboBox
                .DataSource = Nothing
                .Items.Clear()
                .DataSource = Me._ImageItems
                .DisplayMember = "Item"
                .ValueMember = "Value"
            End With
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="ComboBox">Controle ComboBox</param>
        ''' <param name="DrawMode">Modo de desenho</param>
        ''' <param name="ImageSize">Tamanho da imagem</param>
        ''' <param name="UseDropDownList">Define se será utilizadoo tipo DropDownList</param>
        ''' <param name="ImageVerticalSpace">Espaço vertical entre as imagens</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal ComboBox As ComboBox, Optional DrawMode As System.Windows.Forms.DrawMode = Windows.Forms.DrawMode.Normal, Optional ImageSize As Size = Nothing, Optional UseDropDownList As Boolean = True, Optional ImageVerticalSpace As Integer = 1)
            Me._ComboBox = ComboBox
            Me._ImageSize = ImageSize
            Me._ImageVerticalSpace = ImageVerticalSpace
            If UseDropDownList = True Then
                Me._ComboBox.DropDownStyle = ComboBoxStyle.DropDownList
            End If
            If IsNothing(Me._ImageSize) = False Then
                Me._ComboBox.DrawMode = DrawMode
                Me._ComboBox.ItemHeight = (Me._ImageSize.Height + Me._ImageVerticalSpace)
            End If
            Me._ImageItems = New ImageItems(Me._ComboBox)
        End Sub

    End Class

    ''' <summary>
    ''' Coleção de Item
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ImageItems : Inherits List(Of ImageItem)

        'Variáveis
        Private _ComboBox As ComboBox

        ''' <summary>
        ''' Adiciona item
        ''' </summary>
        ''' <param name="Item">Item a ser exibido</param>
        ''' <param name="Value">Valor a ser exibido</param>
        ''' <param name="Image">Imagem a ser exibida</param>
        ''' <param name="OtherValue">Valor adicional para usos diversos</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Add(ByVal Item As String, ByVal Value As Integer, Image As Image, Optional OtherValue As Object = Nothing) As ImageItem
            Dim ListItem2 As New ImageItem(Item, Value, Image, OtherValue)
            MyBase.Add(ListItem2)
            Return ListItem2
        End Function

        ''' <summary>
        ''' Verifica se o valor pertence a coleção
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Contains(Value As Integer) As Boolean
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).Value = Value Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Verifica se o item pertence a coleção
        ''' </summary>
        ''' <param name="Item"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Contains(Item As String) As Boolean
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).Item.ToUpper.Trim = Item.ToUpper.Trim Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' Verifica se o item pertence a coleção
        ''' </summary>
        ''' <param name="OtherValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function ContainsInOtherValue(OtherValue As Object) As Boolean
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).OtherValue.Equals(OtherValue) = True Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        '''  Retorna um item da coleção
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function GetItem(Value As Integer) As ImageItem
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).Value = Value Then
                    Return MyBase.Item(Index)
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        '''  Retorna um item da coleção
        ''' </summary>
        ''' <param name="OtherValue"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function GetItemInOtherValue(OtherValue As Object) As ImageItem
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).OtherValue.Equals(OtherValue) = True Then
                    Return MyBase.Item(Index)
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Retorna um item da coleção
        ''' </summary>
        ''' <param name="Item"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function GetItem(Item As String) As ImageItem
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).Item.ToUpper.Trim = Item.ToUpper.Trim Then
                    Return MyBase.Item(Index)
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Limpa a coleção
        ''' </summary>
        ''' <remarks></remarks>
        Public Overloads Sub Clear()
            Me._ComboBox.DataSource = Nothing
            MyBase.Clear()
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="ComboBox">Controle ComboBox</param>
        ''' <remarks></remarks>
        Public Sub New(ComboBox As ComboBox)
            Me._ComboBox = ComboBox
            Me._ComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        End Sub

    End Class

    ''' <summary>
    ''' Classe para criação de itens associados a valores
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ImageItem

        ''' <summary>
        ''' Armazena o item
        ''' </summary>
        ''' <remarks></remarks>
        Private _Item As String

        ''' <summary>
        ''' Armazena o valor
        ''' </summary>
        ''' <remarks></remarks>
        Private _Value As Integer

        ''' <summary>
        ''' Armazena a imagem
        ''' </summary>
        ''' <remarks></remarks>
        Private _Image As Image

        ''' <summary>
        ''' Valor adicional para usos diversos
        ''' </summary>
        ''' <remarks></remarks>
        Private _OtherValue As Object

        ''' <summary>
        ''' Retorna o item
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Item As String
            Get
                Return _Item
            End Get
        End Property

        ''' <summary>
        ''' Retorna o valor
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Value As Integer
            Get
                Return _Value
            End Get
        End Property

        ''' <summary>
        ''' Imagem
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Image As Image
            Get
                Return _Image
            End Get
        End Property

        ''' <summary>
        ''' Valor adicional para usos diversos
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property OtherValue As Object
            Get
                Return _OtherValue
            End Get
        End Property

        ''' <summary>
        ''' Construtor 
        ''' </summary>
        ''' <param name="Item">Item a ser exibido</param>
        ''' <param name="Value">Número associado ao item</param>
        ''' <param name="Image">Bitmap da imagem</param>
        ''' <param name="OtherValue">Valor adicional para usos diversos</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Item As String, ByVal Value As Integer, Image As Image, Optional OtherValue As Object = Nothing)
            Me._Item = Item
            Me._Value = Value
            Me._Image = Image
            Me._OtherValue = OtherValue
        End Sub

    End Class

End Namespace

