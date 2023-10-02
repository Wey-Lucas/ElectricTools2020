'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows
Imports System.Windows.Forms

Namespace Engine2

    ''' <summary>
    ''' Cria a PaletteSet com o evento PaletteSetClose
    ''' </summary>
    ''' <remarks></remarks>
    Public Class MyPaletteSet : Inherits PaletteSet

        'Variáveis
        Private _idleHandled As Boolean
        Private _CurrentPalette As Palette

        ''' <summary>
        ''' Evento de fechamento da PaletteSet
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <remarks></remarks>
        Public Event PaletteSetClose(sender As Object)

        ''' <summary>
        ''' Retorna o paletteset corrente
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub MyPaletteSet_PaletteActivated(sender As Object, e As PaletteActivatedEventArgs) Handles Me.PaletteActivated
            Me._CurrentPalette = e.Activated
        End Sub

        ''' <summary>
        ''' Paleta corrente
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property CurrentPalette As Palette
            Get
                Return Me._CurrentPalette
            End Get
        End Property

        ''' <summary>
        ''' Detecta o status da paleta
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub MyPaletteSet_StateChanged(sender As Object, e As PaletteSetStateEventArgs) Handles Me.StateChanged
            Select Case e.NewState
                Case StateEventIndex.Hide
                    If Me._idleHandled = False Then
                        Me._idleHandled = True
                        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Me.onIdle
                    End If
                Case StateEventIndex.Show
                    If Me._idleHandled = True Then
                        Me._idleHandled = False
                        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Me.onIdle
                    End If
            End Select
        End Sub

        ''' <summary>
        ''' Detecta o fechamento da paleta
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub onIdle(sender As Object, e As EventArgs)
            RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Me.onIdle
            _idleHandled = False
            If Me.Visible = False Then
                Me._idleHandled = False
                RaiseEvent PaletteSetClose(Me)
            End If
        End Sub

        ''' <summary>
        ''' Contrutor
        ''' </summary>
        ''' <param name="Name">Nome da paleta</param>
        ''' <remarks></remarks>
        Public Sub New(Name As String)
            MyBase.New(Name)
            Me._idleHandled = False
        End Sub

        ''' <summary>
        ''' Construtor 
        ''' </summary>
        ''' <param name="Name">Nome da paleta</param>
        ''' <param name="ToolID">Guid</param>
        ''' <remarks></remarks>
        Public Sub New(Name As String, ToolID As System.Guid)
            MyBase.New(Name, ToolID)
            Me._idleHandled = False
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="Name">Nome da paleta</param>
        ''' <param name="Cmd">Nome do comando</param>
        ''' <param name="ToolID">Guid</param>
        ''' <remarks></remarks>
        Public Sub New(Name As String, Cmd As String, ToolID As System.Guid)
            MyBase.New(Name, Cmd, ToolID)
            Me._idleHandled = False
        End Sub

    End Class

End Namespace
