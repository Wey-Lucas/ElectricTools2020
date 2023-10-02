Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows

Namespace Engine2

    ''' <summary>
    ''' Gerencia menus de contexto no AutoCad
    ''' </summary>
    ''' <remarks></remarks>
    Public Class AcadContextMenuCollection

        ''' <summary>
        ''' ContextMenuExtension
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents _ContextMenuExtension As ContextMenuExtension

        ''' <summary>
        ''' ContextMenuExtension
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ContextMenuExtension As ContextMenuExtension
            Get
                Return Me._ContextMenuExtension
            End Get
            Set(value As ContextMenuExtension)
                Me._ContextMenuExtension = value
            End Set
        End Property

        ''' <summary>
        ''' Evento Click
        ''' </summary>
        ''' <param name="MenuItem">MenuItem</param>
        ''' <param name="e">EventArgs</param>
        ''' <remarks></remarks>
        Public Event Click(MenuItem As MenuItem, e As EventArgs)

        ''' <summary>
        ''' Evento Popup
        ''' </summary>
        ''' <param name="AcadContextMenuCollection">AcadContextMenuCollection</param>
        ''' <param name="e">EventArgs</param>
        ''' <remarks></remarks>
        Public Event Popup(AcadContextMenuCollection As AcadContextMenuCollection, e As EventArgs)

        ''' <summary>
        ''' Dispara o evento Popup
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub InvokePopup(sender As Object, e As EventArgs)
            RaiseEvent Popup(Me, e)
        End Sub

        ''' <summary>
        ''' Dispara o evento Click
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub InvokeClick(sender As Object, e As EventArgs)
            RaiseEvent Click(sender, e)
        End Sub

        ''' <summary>
        ''' Adiciona o menu
        ''' </summary>
        ''' <param name="MenuItem">MenuItem</param>
        ''' <param name="Position">Posição</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Add(MenuItem As MenuItem, Optional Position As Integer = 0) As MenuItem
            AddHandler Me._ContextMenuExtension.Popup, AddressOf InvokePopup
            AddHandler MenuItem.Click, AddressOf InvokeClick
            Me._ContextMenuExtension.MenuItems.Insert(Position, MenuItem)
            Return MenuItem
        End Function

        ''' <summary>
        ''' Adiciona o menu
        ''' </summary>
        ''' <param name="Text">Texto do menu</param>
        ''' <param name="Icon">Ícone</param>
        ''' <param name="Position">Posição</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Add(Text As String, Optional Icon As System.Drawing.Icon = Nothing, Optional Position As Integer = 0) As MenuItem
            Dim MenuItem As MenuItem
            If IsNothing(Icon) = True Then
                MenuItem = New MenuItem(Text)
            Else
                MenuItem = New MenuItem(Text, Icon)
            End If
            AddHandler Me._ContextMenuExtension.Popup, AddressOf InvokePopup
            AddHandler MenuItem.Click, AddressOf InvokeClick
            Me._ContextMenuExtension.MenuItems.Insert(Position, MenuItem)
            Return MenuItem
        End Function

        ''' <summary>
        ''' Remove o menu
        ''' </summary>
        ''' <param name="MenuItem"></param>
        ''' <remarks></remarks>
        Public Overloads Sub Remove(MenuItem As MenuItem)
            RemoveHandler MenuItem.Click, AddressOf InvokeClick
            Me._ContextMenuExtension.MenuItems.Remove(MenuItem)
        End Sub

        ''' <summary>
        ''' Remove o menu
        ''' </summary>
        ''' <param name="Text">Texto do menu</param>
        ''' <remarks></remarks>
        Public Overloads Sub Remove(Text As String)
            Dim MenuItem As MenuItem = Me.GetItem(Text)
            RemoveHandler MenuItem.Click, AddressOf InvokeClick
            Me._ContextMenuExtension.MenuItems.Remove(MenuItem)
        End Sub

        ''' <summary>
        ''' Retorna o menu
        ''' </summary>
        ''' <param name="Text">Texto do menu</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetItem(Text As String) As MenuItem
            For Each MenuItem As MenuItem In Me._ContextMenuExtension.MenuItems
                If MenuItem.Text = Text Then
                    Return MenuItem
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="Type">Tipo de entidade para disparo do evento (GetType(...))</param>
        ''' <param name="Title">Tíyulo</param>
        ''' <remarks></remarks>
        Public Sub New([Type] As Type, Title As String)
            Me._ContextMenuExtension = New ContextMenuExtension
            Me._ContextMenuExtension.Title = Title
            Application.AddObjectContextMenuExtension(Entity.GetClass([Type]), Me._ContextMenuExtension)
        End Sub

    End Class

End Namespace
