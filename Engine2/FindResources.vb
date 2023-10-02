'Imports System
'Imports System.Windows.Forms
'Imports acApp = Autodesk.AutoCAD.ApplicationServices.Application
'Imports Autodesk.AutoCAD.ApplicationServices

'Namespace CsMgdAcad1
'    Partial Public Class Form1
'        Inherits Form

'        Public Sub New()
'            InitializeComponent()
'        End Sub

'        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
'            Dim esc As String = ""
'            Dim cmds As String = CStr(acApp.GetSystemVariable("CMDNAMES"))

'            If cmds.Length > 0 Then
'                Dim cmdNum As Integer = cmds.Split(New Char() {"'"c}).Length

'                For i As Integer = 0 To cmdNum - 1
'                    esc += ChrW(3)
'                Next
'            End If

'            Dim doc As Document = acApp.DocumentManager.MdiActiveDocument
'            doc.SendStringToExecute(esc & "_.LINE ", True, False, True)
'        End Sub
'    End Class
'End Namespace

Imports System
Imports System.ComponentModel
Imports System.Runtime.InteropServices

Namespace Engine2

    Public Class EntryPoint

        <DllImport("kernel32")>
        Private Shared Function LoadLibraryEx(ByVal dllFilePath As String, ByVal hFile As IntPtr, ByVal dwFlags As UInteger) As IntPtr
        End Function

        <DllImport("kernel32", CharSet:=CharSet.Ansi)>
        Public Shared Function GetProcAddress(ByVal dllPointer As IntPtr, ByVal functionName As String) As IntPtr
        End Function

        Private Delegate Function GetApiVersionAddress() As String
        Private Delegate Function GetEntryPointAddress(ByVal msg As Integer, ByVal appId As IntPtr) As Integer

        Private Shared Sub GetEntryPoint(AssemblyFile As String)
            Try
                ' Dim assemblyFile As String = "C:\Program Files\Autodesk\AutoCAD 2020\acblock.arx"
                Dim dllPointer As IntPtr = LoadLibraryEx(AssemblyFile, IntPtr.Zero, &H1)
                If dllPointer = IntPtr.Zero Then Throw New Win32Exception()
                Dim apiVersionAddress As IntPtr = GetProcAddress(dllPointer, "acrxGetApiVersion")
                If apiVersionAddress = IntPtr.Zero Then Throw New Win32Exception()
                Dim entryPointAddress As IntPtr = GetProcAddress(dllPointer, "acrxEntryPoint")
                If entryPointAddress = IntPtr.Zero Then Throw New Win32Exception()

                Try
                    Dim apiVersionAddressFunction = CType(Marshal.GetDelegateForFunctionPointer(apiVersionAddress, GetType(GetApiVersionAddress)), GetApiVersionAddress)
                    Dim version = apiVersionAddressFunction()
                    Dim entryPointAddressFunction = CType(Marshal.GetDelegateForFunctionPointer(entryPointAddress, GetType(GetEntryPointAddress)), GetEntryPointAddress)
                    Dim entryPoint = entryPointAddressFunction(1, IntPtr.Zero)
                Catch ex As AccessViolationException
                End Try

            Catch ex As Exception
            End Try
        End Sub

    End Class

End Namespace

