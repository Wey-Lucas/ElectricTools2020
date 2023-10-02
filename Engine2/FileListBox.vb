'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System.Windows.Forms
Imports System.IO
Imports System.Windows.Forms.ListBox

''' <summary>
''' Cria a caixa de seleção de arquivos a partir de um caminho e filtro
''' </summary>
''' <remarks></remarks>
Public Class FileListBox

    ''' <summary>
    ''' Caminho
    ''' </summary>
    ''' <remarks></remarks>
    Private _Path As String

    ''' <summary>
    ''' Filtro
    ''' </summary>
    ''' <remarks></remarks>
    Private _Pattern As String

    ''' <summary>
    ''' ListBox
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents _ListBox As ListBox

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
            Me.Update()
        End Set
    End Property

    ''' <summary>
    ''' Filtro
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Pattern As String
        Get
            Return Me._Pattern
        End Get
        Set(value As String)
            Me._Pattern = value
            Me.Update()
        End Set
    End Property

    ''' <summary>
    ''' Arquivo selecionado
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FileName As String
        Get
            If Me._ListBox.SelectedIndex <> -1 Then
                Return Me._Path & "\" & Me._ListBox.Text
            Else
                Return Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' Coleção
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Items As ObjectCollection
        Get
            Return Me._ListBox.Items
        End Get
    End Property

    ''' <summary>
    ''' Item selecionado
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SelectedIndex As Integer
        Get
            Return Me._ListBox.SelectedIndex
        End Get
        Set(value As Integer)
            Me._ListBox.SelectedIndex = value
        End Set
    End Property

    ''' <summary>
    ''' Atualiza a lista de arquivos
    ''' </summary>
    ''' <remarks></remarks>
    Public Shadows Sub Update()
        If System.IO.Directory.Exists(Me._Path) = True Then
            Dim DirDiretorio As DirectoryInfo = New DirectoryInfo(Me._Path)
            Dim FileInfo() As FileInfo = DirDiretorio.GetFiles(Me._Pattern)
            Me._ListBox.Items.Clear()
            For Each File As FileInfo In FileInfo
                Me._ListBox.Items.Add(File)
            Next
            If Me._ListBox.Items.Count > 0 Then
                Me._ListBox.Sorted = True
                Me._ListBox.SelectedIndex = 0
            End If
        Else
            MsgBox("O caminho para os arquivos de ajuda '" & Me._Path & "' não pode ser localizado.", MsgBoxStyle.Exclamation)
        End If
    End Sub

    ''' <summary>
    ''' Construtor
    ''' </summary>
    ''' <param name="ListBox">ListBox</param>
    ''' <param name="Path">Caminho</param>
    ''' <param name="Pattern">Filtro</param>
    ''' <remarks></remarks>
    Public Sub New(ListBox As ListBox, Optional Path As String = "C:\", Optional Pattern As String = "*.*")
        Me._ListBox = ListBox
        Me._ListBox.SelectionMode = Windows.Forms.SelectionMode.One
        Me._Path = Path
        Me._Pattern = Pattern
        Me.Update()
    End Sub

End Class
