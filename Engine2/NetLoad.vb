'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports Microsoft.Win32
Imports System.Reflection
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports System.Text
Imports System.Net
Imports System.Collections.Generic
Imports System.Drawing.Imaging
Imports System.Drawing
Imports System.IO
Imports System

Namespace Engine2

    ''' <summary>
    ''' Classe para carregamento automático da DLL no sistema
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class NetLoad

        ''' <summary>
        ''' Registra a aplicação
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub RegisterMyApp()
            Dim sAppName As String = My.Application.Info.AssemblyName
            Dim sProdKey As String = HostApplicationServices.Current.MachineRegistryProductRootKey
            Dim regAcadProdKey As Autodesk.AutoCAD.Runtime.RegistryKey = Autodesk.AutoCAD.Runtime.Registry.CurrentUser.OpenSubKey(sProdKey)
            Dim regAcadAppKey As Autodesk.AutoCAD.Runtime.RegistryKey = regAcadProdKey.OpenSubKey("Applications", True)
            If IsNothing(regAcadAppKey) = True Then
                regAcadProdKey = Autodesk.AutoCAD.Runtime.Registry.CurrentUser.OpenSubKey(sProdKey, True)
                regAcadProdKey.CreateSubKey("Applications")
                regAcadProdKey.Close()
                regAcadProdKey = Autodesk.AutoCAD.Runtime.Registry.CurrentUser.OpenSubKey(sProdKey)
                regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", True)
            End If
            Dim subKeys() As String = regAcadAppKey.GetSubKeyNames()
            For Each sSubKey As String In subKeys
                If (sSubKey.Equals(sAppName)) Then
                    regAcadAppKey.Close()
                    Exit Sub
                End If
            Next
            Dim sAssemblyPath As String = Assembly.GetExecutingAssembly().Location
            Dim regAppAddInKey As Autodesk.AutoCAD.Runtime.RegistryKey = regAcadAppKey.CreateSubKey(sAppName)
            regAppAddInKey.SetValue("DESCRIPTION", sAppName, RegistryValueKind.String)
            regAppAddInKey.SetValue("LOADCTRLS", 14, RegistryValueKind.DWord)
            regAppAddInKey.SetValue("LOADER", sAssemblyPath, RegistryValueKind.String)
            regAppAddInKey.SetValue("MANAGED", 1, RegistryValueKind.DWord)
            regAcadAppKey.Close()
        End Sub

        ''' <summary>
        ''' Desfaz o registro da aplicação
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub UnregisterMyApp()
            Dim sAppName As String = My.Application.Info.AssemblyName
            Dim sProdKey As String = HostApplicationServices.Current.MachineRegistryProductRootKey
            Dim regAcadProdKey As Autodesk.AutoCAD.Runtime.RegistryKey = Autodesk.AutoCAD.Runtime.Registry.CurrentUser.OpenSubKey(sProdKey)
            Dim regAcadAppKey As Autodesk.AutoCAD.Runtime.RegistryKey = regAcadProdKey.OpenSubKey("Applications", True)
            regAcadAppKey.DeleteSubKeyTree(sAppName)
            regAcadAppKey.Close()
        End Sub

    End Class

End Namespace

