'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='
' Adicionar referencias de acWindows.dll e adWindows.dll
'=========================================================================================================='

Imports System.Windows.Forms
Imports System.Drawing
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.AcInfoCenterConn
Imports Autodesk.Internal.InfoCenter
Imports Autodesk.AutoCAD.Windows

Namespace Engine2

    ''' <summary>
    ''' Exibe mensagem de notificação no AutoCAD
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Notify

        'Variáveis
        Private _InfoCenterManager As InfoCenterManager
        Private _ResultItem As ResultItem

        ''' <summary>
        ''' Exibe a mensagem
        ''' </summary>
        ''' <param name="Title">Título</param>
        ''' <param name="Message">Mensagem</param>
        ''' <param name="Period">Período</param>
        ''' <remarks></remarks>
        Public Sub Show(Title As String, Message As String, Optional Period As Integer = 1)
            With Me._InfoCenterManager
                .PaletteManager.Configuration.BalloonDisplayPeriod = Period
            End With
            With Me._ResultItem
                .Category = Title
                .Title = Message
                .IsFavorite = True
                .IsNew = True
            End With
            Me._InfoCenterManager.PaletteManager.ShowBalloon(Me._ResultItem)
        End Sub

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me._InfoCenterManager = New InfoCenterManager
            Me._ResultItem = New ResultItem
        End Sub

    End Class

End Namespace


