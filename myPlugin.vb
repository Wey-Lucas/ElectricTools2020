Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports ElectricTools2020.Engine2

'=========================================================================================================='
'DESENVOLVIDO POR: LUCAS LEÔNCIO WEY DA SILVA (FORMANDO EM ANALISTA DE SISTEMAS)
'EM:2023
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: wey.lucas1@gmail.com \ (011) 99940-2202
'=========================================================================================================='

' This line is not mandatory, but improves loading performances
<Assembly: ExtensionApplication(GetType(ElectricTools.MyPlugin))>

Namespace ElectricTools

    ' This class is instantiated by AutoCAD once and kept alive for the 
    ' duration of the session. If you don't do any one time initialization 
    ' then you should remove this class.
    Public Class MyPlugin : Implements IExtensionApplication

        'Declarações (PlugIn)
        Private WithEvents acadApplication As Autodesk.AutoCAD.Interop.AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication
        Private WithEvents acadDocument As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Private WithEvents AcadDocumentCollection As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Private WithEvents acadEditor As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        'Declarações
        Private _SPHESysUpdate As SPHEUpdate47.SPHESysUpdate.AutoUpdate
        Private Shared _Guid As String
        Private Shared _SystemMemory As SystemMemory
        'Friend Shared ReadOnly ClearenceBlockFilterCollection As FilterCollection

        ''' <summary>
        ''' Inicialização
        ''' </summary>
        Public Sub Initialize() Implements IExtensionApplication.Initialize
            Engine2.NetLoad.RegisterMyApp()

            With acadEditor
                .WriteMessage(vbLf)
                .WriteMessage(vbLf & "ElectricTools carregado com sucesso...")
                .WriteMessage(vbLf)
            End With
        End Sub

        ''' <summary>
        ''' Término
        ''' </summary>
        Public Sub Terminate() Implements IExtensionApplication.Terminate
            With acadEditor
                .WriteMessage(vbLf)
                .WriteMessage(vbLf & "ElectricTools finalizado com sucesso...")
                .WriteMessage(vbLf)
            End With
        End Sub

        ''' <summary>
        ''' Retorna a conexão
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property StringConnection As String
            Get
                Return "Persist Security Info=False;User ID=sphe_cad;Initial Catalog=SPHECAD;Data Source=CETUS\SQLEXPRESS;Password=sphe"
            End Get
        End Property

        ''' <summary>
        ''' Atualiza o documento corrente e fecha o controle
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub AcadDocumentCollection_DocumentActivated(sender As Object, e As DocumentCollectionEventArgs) Handles AcadDocumentCollection.DocumentActivated
            Me.acadDocument = e.Document
        End Sub

        ''' <summary>
        ''' Fecha o controle ao abrir ou criar um documento
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub AcadDocumentCollection_DocumentCreateStarted(sender As Object, e As DocumentCollectionEventArgs) Handles AcadDocumentCollection.DocumentCreateStarted
            Me.acadDocument = e.Document
        End Sub

        ''' <summary>
        ''' Fecha o controle ao fechar o documento
        ''' </summary>
        ''' <param name="sender">Objeto de chamada</param>
        ''' <param name="e">Argumentos do evento</param>
        ''' <remarks></remarks>
        Private Sub AcadDocumentCollection_DocumentDestroyed(sender As Object, e As Autodesk.AutoCAD.ApplicationServices.DocumentDestroyedEventArgs) Handles AcadDocumentCollection.DocumentDestroyed
            Me.acadDocument = Nothing
        End Sub

        ''' <summary>
        ''' Lista de Nomes de bloco
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property BlockNameFilterCollection As BlockPropertiesTable

        ''' <summary>
        ''' Retorna o Guide de sistema
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property Guid As String
            Get
                Return _Guid
            End Get
        End Property

        ''' <summary>
        ''' Retorna o sistema de memória
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property SystemMemory As SystemMemory
            Get
                Return _SystemMemory
            End Get
        End Property

        ''' <summary>
        ''' Teste boolean
        ''' </summary>
        ''' <returns></returns>
        Public Shared Property Teste As Boolean
            Get
                Return If(_SystemMemory.GetValue("Teste") = "0", False, True)
            End Get
            Set(value As Boolean)
                _SystemMemory.SetValue("Teste", If(value = True, "1", "0"))
            End Set
        End Property

        Public Sub New()

            'SPHESysUpdate
            Me._SPHESysUpdate = New SPHEUpdate47.SPHESysUpdate.AutoUpdate("ElectricTools2020", acadApplication.HWND, SPHEUpdate47.SPHESysUpdate.AutoUpdate.eUpdateMode.InForeClosure)
            Me._SPHESysUpdate.Initialize()

            'Inicia a memória
            _SystemMemory = New SystemMemory(Regedit.RegistryRoots.HKEY_CURRENT_USER, "Software\DTSIMP")

            'Grava as memórias
            With _SystemMemory

                'Teste
                .AddValue("Teste", "0", Microsoft.Win32.RegistryValueKind.String)

            End With

            'Guid
            _Guid = adoNetExtension.AdoNetTools.Regedit.GetValue(adoNetExtension.AdoNetTools.Regedit.RegistryRoots.HKEY_CURRENT_USER, "Software\Microsoft\Windows\Guid", "Guid5000")

        End Sub

    End Class

End Namespace