Imports Microsoft.Win32

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Gerencia a memória para valores diversos do sistema
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SystemMemory

        'Declarações
        Private _RegistryRoot As Engine2.Regedit.RegistryRoots
        Private _Path As String

        ''' <summary>
        ''' Adiciona o valor de memória (Adiciona item apenas se ele não existir)
        ''' </summary>
        ''' <param name="Key"></param>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Public Sub AddValue(Key As String, Value As Object, Kind As Microsoft.Win32.RegistryValueKind)
            If IsNothing(Engine2.Regedit.GetValue(Me._RegistryRoot, Me._Path, Key)) = True Then
                Engine2.Regedit.AddValue(Me._RegistryRoot, Me._Path, Key, Value, Kind)
            End If
        End Sub

        ''' <summary>
        ''' Atualiza o valor de memória
        ''' </summary>
        ''' <param name="Key"></param>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Public Function SetValue(Key As String, Value As Object) As Object
            Dim RegValue As Object = Engine2.Regedit.GetValue(Me._RegistryRoot, Me._Path, Key)
            If IsNothing(RegValue) = False Then
                Return Engine2.Regedit.AddValue(Me._RegistryRoot, Me._Path, Key, Value, Me.GetKind(Key))
            Else
                Return RegValue
            End If
        End Function

        ''' <summary>
        ''' Obtem o valor memorizado
        ''' </summary>
        ''' <param name="Key"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetValue(Key As String) As Object
            Return Engine2.Regedit.GetValue(Me._RegistryRoot, Me._Path, Key)
        End Function

        ''' <summary>
        ''' Construtor
        ''' </summary>
        ''' <param name="RegistryRoot">Chave de registro</param>
        ''' <param name="Path">Caminho do registro</param>
        ''' <remarks></remarks>
        Public Sub New(RegistryRoot As Engine2.Regedit.RegistryRoots, Path As String)
            Me._RegistryRoot = RegistryRoot
            Me._Path = Path
        End Sub

        ''' <summary>
        ''' Obtem o formato de dados da chave de registro
        ''' </summary>
        ''' <param name="Key"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetKind(Key As String) As Microsoft.Win32.RegistryValueKind
            Select Case Me._RegistryRoot
                Case Regedit.RegistryRoots.HKEY_CLASSES_ROOT
                    Return My.Computer.Registry.ClassesRoot.OpenSubKey(Me._Path).GetValueKind(Key)
                Case Regedit.RegistryRoots.HKEY_CURRENT_USER
                    Return My.Computer.Registry.CurrentUser.OpenSubKey(Me._Path).GetValueKind(Key)
                Case Regedit.RegistryRoots.HKEY_LOCAL_MACHINE
                    Return My.Computer.Registry.LocalMachine.OpenSubKey(Me._Path).GetValueKind(Key)
                Case Regedit.RegistryRoots.HKEY_PERFORMANCE_DATA
                    Return My.Computer.Registry.PerformanceData.OpenSubKey(Me._Path).GetValueKind(Key)
                Case Regedit.RegistryRoots.HKEY_USERS
                    Return My.Computer.Registry.Users.OpenSubKey(Me._Path).GetValueKind(Key)
            End Select
        End Function

    End Class

End Namespace


