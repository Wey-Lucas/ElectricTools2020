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
    ''' Converte um ComboBox em DropDownList Compatível com imagens
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DropDownList

        'Declarações
        Private WithEvents _ComboBox As ComboBox
        Private _Items As Items
        Private _ImageList As ImageList
        Private _ImageVerticalSpace As Integer

        ''' <summary>
        ''' Coleção de itens
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Items As Items
            Get
                Return Me._Items
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
                If IsNothing(Me._ComboBox.SelectedItem) = False Then
                    Return Me._ComboBox.SelectedItem.Item
                Else
                    Return Nothing
                End If
            End Get
            Set(value As Object)
                Me._ComboBox.SelectedIndex = -1
                For Index As Integer = 0 To Me._ComboBox.Items.Count - 1
                    Dim Item As Item = Me._ComboBox.Items.Item(Index)
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
                If IsNothing(Me._ComboBox.SelectedItem) = False Then
                    Return Me._ComboBox.SelectedItem.Value
                Else
                    Return Nothing
                End If
            End Get
            Set(value As Object)
                Me._ComboBox.SelectedIndex = -1
                For Index As Integer = 0 To Me._ComboBox.Items.Count - 1
                    Dim Item As Item = Me._ComboBox.Items.Item(Index)
                    If Item.Value = value Then
                        Me._ComboBox.SelectedItem = Item
                        Exit For
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
                Return Me._ComboBox.SelectedIndex
            End Get
            Set(value As Integer)
                Me._ComboBox.SelectedIndex = value
            End Set
        End Property

        ''' <summary>
        ''' Retorna a imagem selecionada
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Image As Image
            Get
                If IsNothing(Me._ImageList) = False Then
                    Return Me._ImageList.Images(Me._ComboBox.SelectedIndex)
                Else
                    Return Nothing
                End If
            End Get
        End Property

        ''' <summary>
        ''' Desenha imagem e texto
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _ComboBox_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles _ComboBox.DrawItem
            Dim TextSize As SizeF
            If e.Index <> -2 Then
                If e.Index <= Me._ComboBox.Items.Count - 1 And e.Index > -1 Then
                    Dim Item As Item = Me._ComboBox.Items(e.Index)
                    e.DrawBackground()
                    If Item.ImageIndex <> -2 Then
                        e.Graphics.FillRectangle(New SolidBrush(Color.LightBlue), e.Bounds)
                        If IsNothing(Me._ImageList) = False Then
                            If Me._ImageList.Images.ContainsKey(CStr(Item.ImageIndex)) = True Then
                                e.Graphics.DrawImage(Me._ImageList.Images(CStr(Item.ImageIndex)), New Point(e.Bounds.X, e.Bounds.Y + (Me._ImageVerticalSpace / 2)))
                            Else
                                e.Graphics.DrawImage(Me._ImageList.Images(Item.ImageIndex), New Point(e.Bounds.X, e.Bounds.Y + (Me._ImageVerticalSpace / 2)))
                            End If
                            TextSize = Engine2.MeasureString.GetMeasure(Me._ComboBox.Font, Item.Item, Me._ComboBox.CreateGraphics)
                            e.Graphics.DrawString(Item.Item, Me._ComboBox.Font, Brushes.Black, New Point(Me._ImageList.ImageSize.Width + 3, ((e.Bounds.Y + (Me._ImageList.ImageSize.Height / 2)) - (TextSize.Height / 2))))
                        Else
                            e.Graphics.DrawString(Item.Item, Me._ComboBox.Font, Brushes.Black, New Point(e.Bounds.X, e.Bounds.Y))
                        End If
                    End If
                    If (e.State And DrawItemState.Focus) = 0 Then
                        e.Graphics.FillRectangle(New SolidBrush(Color.White), e.Bounds)
                        If IsNothing(Me._ImageList) = False Then
                            If Me._ImageList.Images.ContainsKey(CStr(Item.ImageIndex)) = True Then
                                e.Graphics.DrawImage(Me._ImageList.Images(CStr(Item.ImageIndex)), New Point(e.Bounds.X, e.Bounds.Y + (Me._ImageVerticalSpace / 2)))
                            Else
                                e.Graphics.DrawImage(Me._ImageList.Images(Item.ImageIndex), New Point(e.Bounds.X, e.Bounds.Y + (Me._ImageVerticalSpace / 2)))
                            End If
                            TextSize = Engine2.MeasureString.GetMeasure(Me._ComboBox.Font, Item.Item, Me._ComboBox.CreateGraphics)
                            e.Graphics.DrawString(Item.Item, Me._ComboBox.Font, Brushes.Black, New Point(Me._ImageList.ImageSize.Width + 3, ((e.Bounds.Y + (Me._ImageList.ImageSize.Height / 2)) - (TextSize.Height / 2))))
                        Else
                            e.Graphics.DrawString(Item.Item, Me._ComboBox.Font, Brushes.Black, New Point(e.Bounds.X, e.Bounds.Y))
                        End If
                    End If
                    e.DrawFocusRectangle()
                End If
            End If
        End Sub

        ''' <summary>
        ''' Ajusta o tamanho
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _ComboBox_MeasureItem(ByVal sender As Object, ByVal e As System.Windows.Forms.MeasureItemEventArgs) Handles _ComboBox.MeasureItem
            If IsNothing(Me._ImageList) = False Then
                e.ItemHeight = (Me._ImageList.ImageSize.Height + Me._ImageVerticalSpace)
                e.ItemWidth = Me._ImageList.ImageSize.Width
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
                .DataSource = Me._Items
                .DisplayMember = "Item"
                .ValueMember = "Value"
            End With
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="ComboBox">Controle ComboBox</param>
        ''' <param name="DrawMode">Modo de desenho</param>
        ''' <param name="ImageList">Coleção de imagem</param>
        ''' <param name="UseDropDownList">Define se será utilizadoo tipo DropDownList</param>
        ''' <param name="ImageVerticalSpace">Espaço vertical entre as imagens</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal ComboBox As ComboBox, Optional DrawMode As System.Windows.Forms.DrawMode = Windows.Forms.DrawMode.Normal, Optional ImageList As ImageList = Nothing, Optional UseDropDownList As Boolean = True, Optional ImageVerticalSpace As Integer = 1)
            Me._ImageList = ImageList
            Me._ComboBox = ComboBox
            Me._ImageVerticalSpace = ImageVerticalSpace
            If UseDropDownList = True Then
                Me._ComboBox.DropDownStyle = ComboBoxStyle.DropDownList
            End If
            If IsNothing(Me._ImageList) = False Then
                Me._ComboBox.DrawMode = DrawMode
                Me._ComboBox.ItemHeight = (Me._ImageList.ImageSize.Height + Me._ImageVerticalSpace)
            End If
            Me._Items = New Items(Me._ComboBox)
        End Sub

    End Class

    ''' <summary>
    ''' Coleção de Item
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Items : Inherits List(Of Item)

        'Variáveis
        Private _ComboBox As ComboBox

        ''' <summary>
        ''' Adiciona item
        ''' </summary>
        ''' <param name="Item">Item a ser exibido</param>
        ''' <param name="Value">Valor a ser exibido</param>
        ''' <param name="ImageIndex">Imagem a ser exibida</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Add(ByVal Item As String, ByVal Value As Integer, Optional ImageIndex As Integer = -1) As Item
            Dim ListItem2 As New Item(Item, Value, ImageIndex)
            MyBase.Add(ListItem2)
            Return ListItem2
        End Function

        ''' <summary>
        ''' Adiciona um enumerador
        ''' </summary>
        ''' <param name="EnumType">Enumerador, Exemplo: GetType(Nome do enumerador))</param>
        ''' <param name="RemoveValues">Coleção de valores a não serem considerados no carregamento</param>
        ''' <remarks></remarks>
        Public Overloads Sub Add(ByVal EnumType As System.Type, Optional RemoveValues As ArrayList = Nothing)
            Dim lstText() As String = System.Enum.GetNames(EnumType)
            For ini As Integer = 0 To lstText.Length - 1
                If IsNothing(RemoveValues) = False Then
                    If RemoveValues.Contains(CInt(System.Enum.GetValues(EnumType)(ini))) = False Then
                        MyBase.Add(New Item(lstText(ini), CInt(System.Enum.GetValues(EnumType)(ini))))
                    End If
                Else
                    MyBase.Add(New Item(lstText(ini), CInt(System.Enum.GetValues(EnumType)(ini))))
                End If
            Next
        End Sub

        ''' <summary>
        ''' Adiciona um DataTable
        ''' </summary>
        ''' <param name="DataTable">DataTable</param>
        ''' <param name="ItemField">Nome do campo que armazena o Item</param>
        ''' <param name="ValueField">Nome do campo que armazena o valor associado ao item</param>
        ''' <param name="ImageIndexField">Nome do campo que armazena a posição da imagem</param>
        ''' <remarks></remarks>
        Public Overloads Sub Add(ByVal DataTable As System.Data.DataTable, ItemField As String, ValueField As String, Optional ImageIndexField As Integer = -1)
            For Index As Integer = 0 To DataTable.Rows.Count - 1
                MyBase.Add(New Item(DataTable.Rows.Item(Index).Item(ItemField), DataTable.Rows.Item(Index).Item(ValueField), If(ImageIndexField <> -1, DataTable.Rows.Item(Index).Item(ImageIndexField), -1)))
            Next
        End Sub

        ''' <summary>
        ''' Verifica se o ImageIndex pertence a coleção
        ''' </summary>
        ''' <param name="ImageIndex"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function ContainsImageIndex(ImageIndex As Integer) As Boolean
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).ImageIndex = ImageIndex Then
                    Return True
                End If
            Next
            Return False
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
        '''  Retorna um item da coleção
        ''' </summary>
        ''' <param name="ImageIndex"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function GetItemFromImageIndex(ImageIndex As Integer) As Item
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).ImageIndex = ImageIndex Then
                    Return MyBase.Item(Index)
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        '''  Retorna um item da coleção
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function GetItem(Value As Integer) As Item
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).Value = Value Then
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
        Public Overloads Function GetItem(Item As String) As Item
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).Item.ToUpper.Trim = Item.ToUpper.Trim Then
                    Return MyBase.Item(Index)
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Remove item da coleção
        ''' </summary>
        ''' <param name="Item">Item</param>
        ''' <remarks></remarks>
        Public Overloads Sub RemoveAT(Item As String)
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).Item = Item Then
                    Me._ComboBox.Items.RemoveAt(Index)
                    Exit For
                End If
            Next
        End Sub

        ''' <summary>
        ''' Remove item da coleção
        ''' </summary>
        ''' <param name="Value">Valor associado ao item</param>
        ''' <remarks></remarks>
        Public Overloads Sub RemoveAT(Value As Integer)
            For Index As Integer = 0 To MyBase.Count - 1
                If MyBase.Item(Index).Value = Value Then
                    Me._ComboBox.Items.RemoveAt(Index)
                    Exit For
                End If
            Next
        End Sub

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
        End Sub

    End Class

    ''' <summary>
    ''' Classe para criação de itens associados a valores
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Item

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
        ''' Armazena a posição da imagem
        ''' </summary>
        ''' <remarks></remarks>
        Private _ImagemIndex As Integer

        ''' <summary>
        ''' Armazena a chave da imagem
        ''' </summary>
        ''' <remarks></remarks>
        Private _ImageKey As String

        ''' <summary>
        ''' Armazena a imagem
        ''' </summary>
        ''' <remarks></remarks>
        Private _Image As Image

        ''' <summary>
        ''' Retorna a imagem
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
        ''' Retorna a posição da imagem
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ImageIndex As Integer
            Get
                Return _ImagemIndex
            End Get
        End Property

        ''' <summary>
        ''' Construtor 
        ''' </summary>
        ''' <param name="Item">Item a ser exibido</param>
        ''' <param name="Value">Número associado ao item</param>
        ''' <param name="ImageIndex">Posição da imagem</param>
        ''' <param name="ImageKey">Chave da imagem</param>
        ''' <param name="Image">Bitmap da imagem</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Item As String, ByVal Value As Integer, Optional ImageIndex As Integer = -2, Optional ImageKey As String = "", Optional Image As Image = Nothing)
            _Item = Item
            _Value = Value
            _ImagemIndex = ImageIndex
            _ImageKey = ImageKey
            _Image = Image
        End Sub

    End Class

End Namespace

