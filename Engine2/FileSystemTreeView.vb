'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System.IO
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Collections
Imports System.Drawing

Namespace Engine2

    ''' <summary>
    ''' Classe de carregamento para TreeView
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FileSystemTreeView : Inherits TreeView

        'Declarações
        Private _Filter As String
        Private _Path As String
        Private _DirectoryImageIndex As Integer
        Private _FileImageIndex As Integer

        ''' <summary>
        ''' Imagem para o diretório
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DirectoryImageIndex As Integer
            Get
                Return Me._DirectoryImageIndex
            End Get
            Set(value As Integer)
                Me._DirectoryImageIndex = value
                Me.LoadTreeView()
            End Set
        End Property

        ''' <summary>
        ''' Imagem para arquivo
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FileImageIndex As Integer
            Get
                Return Me._FileImageIndex
            End Get
            Set(value As Integer)
                Me._FileImageIndex = value
                Me.LoadTreeView()
            End Set
        End Property

        ''' <summary>
        ''' Caminho
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Path As String
            Get
                Return Me._Path
            End Get
            Set(value As String)
                Me._Path = value
                Me.LoadTreeView()
            End Set
        End Property

        ''' <summary>
        ''' Filtro
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Filter As String
            Get
                Return Me._Filter
            End Get
            Set(value As String)
                Me._Filter = value
                Me.LoadTreeView()
            End Set
        End Property

        ''' <summary>
        ''' Detecta a seleção e expande o nó
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub FileSystemTreeView_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles Me.BeforeExpand
            If Not TypeOf e.Node Is FileNode Then
                Dim node As DirectoryNode = e.Node
                If Not node.Loaded Then
                    node.Nodes(0).Remove()
                    node.LoadDirectory()
                    node.LoadFiles()
                End If
            End If
        End Sub

        ''' <summary>
        ''' Detecta a seleção do nó e dispara o carregamento
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub FileSystemTreeView_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
            Dim node As TreeNode = Me.GetNodeAt(e.X, e.Y)
            If IsNothing(node) = False Then
                Me.SelectedNode = node
            End If
        End Sub

        ''' <summary>
        ''' Carrega a TreeView
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub LoadTreeView()
            If Directory.Exists(Me._Path) = False Then
                Throw New System.Exception("O diretório '" & Me._Path & "' não pode ser encontrado")
            End If
            If Me._Filter.Trim = "" Then
                Throw New System.Exception("O filtro não foi informado")
            End If
            Nodes.Clear()
            Dim node As New DirectoryNode(Me, New DirectoryInfo(Me._Path))
            node.Expand()
        End Sub

        ''' <summary>
        ''' Contrutor
        ''' </summary>
        ''' <param name="Path">Caminho</param>
        ''' <param name="Filter">Filtro</param>
        ''' <param name="ImageList">ImageList</param>
        ''' <param name="DirectoryImageIndex">Imagem para diretório</param>
        ''' <param name="FileImageIndex">Imagem para arquivo</param>
        ''' <remarks></remarks>
        Public Sub New(Path As String, Filter As String, ImageList As ImageList, DirectoryImageIndex As Integer, FileImageIndex As Integer)
            Me._Path = Path
            Me._Filter = Filter
            Me.ImageList = ImageList
            Me._DirectoryImageIndex = DirectoryImageIndex
            Me._FileImageIndex = FileImageIndex
        End Sub

        ''' <summary>
        ''' Contrutor
        ''' </summary>
        ''' <param name="Path">Caminho</param>
        ''' <param name="Filter">Filtro</param>
        ''' <param name="DirectoryImageIndex">Imagem para diretório</param>
        ''' <param name="FileImageIndex">Imagem para arquivo</param>
        ''' <remarks></remarks>
        Public Sub New(Path As String, Filter As String, DirectoryImageIndex As Integer, FileImageIndex As Integer)
            Me._Path = Path
            Me._Filter = Filter
            Me.ImageList = Me.ImageList
            Me._DirectoryImageIndex = DirectoryImageIndex
            Me._FileImageIndex = FileImageIndex
        End Sub

    End Class

    ''' <summary>
    ''' Complemento de FileSystemTreeView (Para carregamento de diretórios)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DirectoryNode : Inherits TreeNode

        'Variáveis
        Private _FileSystemTreeView As FileSystemTreeView
        Private _DirectoryInfo As DirectoryInfo

        ''' <summary>
        ''' Evento que permite a edição de valores na criação dos diretórios
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <remarks></remarks>
        Public Event DirectoryChangeText(ByRef sender As Object)

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="FileSystemTreeView">FileSystemTreeView</param>
        ''' <param name="DirectoryNode">DirectoryNode</param>
        ''' <param name="DirectoryInfo">DirectoryInfo</param>
        ''' <remarks></remarks>
        Public Sub New(FileSystemTreeView As FileSystemTreeView, DirectoryNode As DirectoryNode, DirectoryInfo As DirectoryInfo)
            MyBase.New(DirectoryInfo.Name)
            Me._FileSystemTreeView = FileSystemTreeView
            Me._DirectoryInfo = DirectoryInfo
            Me.ImageIndex = Me._FileSystemTreeView.DirectoryImageIndex
            Me.SelectedImageIndex = Me.ImageIndex
            DirectoryNode.Nodes.Add(Me)
            Me.Tag = Me._DirectoryInfo.FullName
            RaiseEvent DirectoryChangeText(Me)
            Me.Virtualize()
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="FileSystemTreeView">FileSystemTreeView</param>
        ''' <param name="DirectoryInfo">DirectoryInfo</param>
        ''' <remarks></remarks>
        Public Sub New(FileSystemTreeView As FileSystemTreeView, DirectoryInfo As DirectoryInfo)
            MyBase.New(DirectoryInfo.Name)
            Me._FileSystemTreeView = FileSystemTreeView
            Me._DirectoryInfo = DirectoryInfo
            Me.ImageIndex = Me._FileSystemTreeView.DirectoryImageIndex
            Me.SelectedImageIndex = Me.ImageIndex
            FileSystemTreeView.Nodes.Add(Me)
            Me.Tag = Me._DirectoryInfo.FullName
            RaiseEvent DirectoryChangeText(Me)
            Me.Virtualize()
        End Sub

        ''' <summary>
        ''' Virtualiza o processo de carregamento
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Virtualize()
            Dim Count As Integer = 0
            Try
                Count = Me._DirectoryInfo.GetFiles().Length
                If (Count + Me._DirectoryInfo.GetDirectories().Length) > 0 Then
                    Dim FakeChildNode = New FakeChildNode(Me)
                End If
            Catch
                Exit Try
            End Try
        End Sub

        ''' <summary>
        ''' Carrega os diretórios
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub LoadDirectory()
            For Each DirectoryInfo As DirectoryInfo In _DirectoryInfo.GetDirectories()
                Dim DirectoryNode As New DirectoryNode(Me._FileSystemTreeView, Me, DirectoryInfo)
            Next
        End Sub

        ''' <summary>
        ''' Carrega os arquivos
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub LoadFiles()
            For Each File As FileInfo In _DirectoryInfo.GetFiles(Me._FileSystemTreeView.Filter)
                Dim FileNode As New FileNode(Me._FileSystemTreeView, Me, File)
            Next
        End Sub

        ''' <summary>
        ''' Determina se o processo foi carregado
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Loaded() As Boolean
            Get
                If Me.Nodes.Count <> 0 Then
                    If TypeOf Me.Nodes(0) Is FakeChildNode Then
                        Return False
                    End If
                End If
                Return True
            End Get
        End Property

    End Class

    ''' <summary>
    ''' Complemento de FileSystemTreeView (Para carregamento de arquivos)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FileNode : Inherits TreeNode

        'Variáveis
        Private _FileInfo As FileInfo
        Private _DirectoryNode As DirectoryNode

        ''' <summary>
        ''' Evento que permite a edição de valores na criação de arquivos
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <remarks></remarks>
        Public Event FileChangeText(ByRef sender As Object)

        ''' <summary>
        ''' Contrutor
        ''' </summary>
        ''' <param name="FileSystemTreeView">FileSystemTreeView</param>
        ''' <param name="DirectoryNode">DirectoryNode</param>
        ''' <param name="FileInfo">FileInfo</param>
        ''' <remarks></remarks>
        Public Sub New(FileSystemTreeView As FileSystemTreeView, DirectoryNode As DirectoryNode, FileInfo As FileInfo)
            MyBase.New(FileInfo.Name)
            Me._DirectoryNode = DirectoryNode
            Me._FileInfo = FileInfo
            Me.ImageIndex = FileSystemTreeView.FileImageIndex
            Me.SelectedImageIndex = Me.ImageIndex
            Me._DirectoryNode.Nodes.Add(Me)
            Me.Tag = FileInfo.FullName
            RaiseEvent FileChangeText(Me)
        End Sub

    End Class

    ''' <summary>
    ''' Complemento de DirectoryNode
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FakeChildNode : Inherits TreeNode

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="TreeNode">TreeNode</param>
        ''' <remarks></remarks>
        Public Sub New(TreeNode As TreeNode)
            MyBase.New()
            TreeNode.Nodes.Add(Me)
        End Sub

    End Class

End Namespace

