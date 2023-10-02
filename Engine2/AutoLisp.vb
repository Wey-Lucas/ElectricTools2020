Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports System.Runtime.InteropServices
Imports System.Text

Namespace Engine2

    ''' <summary>
    ''' Classe para dados extendidos de entidades.
    ''' Permite que as linhas sejam editadas automaticamente.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class AutoLisp

#Region "Refeências externas"

        'ACAD 2009
        '<System.Security.SuppressUnmanagedCodeSecurity()> <DllImport("acad.exe", EntryPoint:="acedPutSym", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)> _
        'Private Shared Function acedPutSym(args As String, result As IntPtr) As Integer
        'End Function

        '<System.Security.SuppressUnmanagedCodeSecurity()> <DllImport("acad.exe", EntryPoint:="acedGetSym", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)> _
        'Private Shared Function acedGetSym(args As String, ByRef result As IntPtr) As Integer
        'End Function

        '<System.Security.SuppressUnmanagedCodeSecurity()> <DllImport("acad.exe", EntryPoint:="acedInvoke", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)> _
        'Private Shared Function acedInvoke(ByVal args As IntPtr, ByRef result As IntPtr) As Integer
        'End Function


        'ACAD 2015
        <System.Security.SuppressUnmanagedCodeSecurity()> <DllImport("acCore.dll", EntryPoint:="acedPutSym", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)> _
        Private Shared Function acedPutSym(args As String, result As IntPtr) As Integer
        End Function

        <System.Security.SuppressUnmanagedCodeSecurity()> <DllImport("acCore.dll", EntryPoint:="acedGetSym", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)> _
        Shared Function acedGetSym(ByVal args As String, <Out()> ByRef result As IntPtr) As Integer
        End Function

        <System.Security.SuppressUnmanagedCodeSecurity()> <DllImport("acCore.dll", EntryPoint:="acedInvoke", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)> _
        Private Shared Function acedInvoke(ByVal args As IntPtr, ByRef result As IntPtr) As Integer
        End Function

#End Region

        ''' <summary>
        ''' Executa funções do Autolisp
        ''' </summary>
        ''' <param name="args">Argumentos</param>
        ''' <returns>ResultBuffer</returns>
        ''' <remarks></remarks>
        Public Shared Function LispInvoke(ByVal args As ResultBuffer) As ResultBuffer
            Dim ip As IntPtr = IntPtr.Zero
            Dim status As Integer = acedInvoke(args.UnmanagedObject, ip)
            If status = CInt(PromptStatus.OK) AndAlso ip <> IntPtr.Zero Then
                Return ResultBuffer.Create(ip, True)
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Seta o valor de uma variável de sistema
        ''' </summary>
        ''' <param name="Variable">Variável</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Object</returns>
        ''' <remarks></remarks>
        Public Shared Function SetVar(Variable As String, Value As Object) As Object
            Try
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable(Variable, Value)
                Return Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable(Variable)
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Engine.Autolisp.SetVar, Motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Obtem o valor de uma variável de sistema
        ''' </summary>
        ''' <param name="Variable">Variável</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetVar(Variable As String) As Object
            Try
                Return Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable(Variable)
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Engine.Autolisp.GetVar, Motivo: " & ex.Message)
            End Try

        End Function

        ''' <summary>
        ''' Seta o valor de uma variável AutoLisp
        ''' </summary>
        ''' <param name="Variable">Variável</param>
        ''' <param name="Value">Valor</param>
        ''' <returns>Object</returns>
        ''' <remarks></remarks>
        Public Shared Function SetVarLisp(Variable As String, Value As Object) As Object
            Try
                Dim Input As New ResultBuffer
                Dim Result As Integer = 0
                With Input
                    .Add(New TypedValue(LispDataType.ListBegin))
                    Select Case Value.GetType.Name
                        Case "String"
                            .Add(New TypedValue(LispDataType.Text, Value))
                        Case "Integer", "Int32"
                            .Add(New TypedValue(LispDataType.Int32, Value))
                        Case "Short", "Int16"
                            .Add(New TypedValue(LispDataType.Int16, Value))
                        Case "Double"
                            .Add(New TypedValue(LispDataType.Double, Value))
                        Case Else
                            .Add(New TypedValue(LispDataType.None, Value))
                    End Select
                    .Add(New TypedValue(LispDataType.ListEnd))
                End With
                acedPutSym(Variable, Input.UnmanagedObject)
                Input.Dispose()
                Return Engine2.AutoLisp.GetVarLisp(Variable)
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Engine.Autolisp.SetVarLisp, Motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Obtem o valor de uma variável AutoLisp
        ''' </summary>
        ''' <param name="Variable">Variável</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetVarLisp(Variable As String) As Object
            Try
                Dim IntPtr As IntPtr = IntPtr.Zero
                Dim Status As Integer = acedGetSym(Variable, IntPtr)
                If Status = CInt(PromptStatus.OK) AndAlso IntPtr <> IntPtr.Zero Then
                    For Each TypedValue As TypedValue In ResultBuffer.Create(IntPtr, True)
                        Select Case TypedValue.TypeCode
                            Case LispDataType.Text
                                Return TypedValue.Value
                            Case LispDataType.Int32
                                Return TypedValue.Value
                            Case LispDataType.Int16
                                Return TypedValue.Value
                            Case LispDataType.Double
                                Return TypedValue.Value
                            Case LispDataType.None
                                Return TypedValue.Value
                        End Select
                    Next
                End If
                Return Nothing
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Engine.Autolisp.GetVarLisp, Motivo: " & ex.Message)
            End Try
        End Function

    End Class

End Namespace


