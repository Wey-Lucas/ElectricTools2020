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
    ''' Classe avançada para edição de registros do Windows
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Regedit

        ''' <summary>
        ''' Enumera as raizes do registro
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum RegistryRoots
            ''' <summary>
            ''' HKEY_CLASSES_ROOT
            ''' </summary>
            ''' <remarks></remarks>
            HKEY_CLASSES_ROOT = 0
            ''' <summary>
            ''' HKEY_CURRENT_USER
            ''' </summary>
            ''' <remarks></remarks>
            HKEY_CURRENT_USER = 1
            ''' <summary>
            ''' HKEY_LOCAL_MACHINE
            ''' </summary>
            ''' <remarks></remarks>
            HKEY_LOCAL_MACHINE = 2
            ''' <summary>
            ''' HKEY_PERFORMANCE_DATA
            ''' </summary>
            ''' <remarks></remarks>
            HKEY_PERFORMANCE_DATA = 3
            ''' <summary>
            ''' HKEY_USERS
            ''' </summary>
            ''' <remarks></remarks>
            HKEY_USERS = 4
        End Enum

        ''' <summary>
        ''' Registro
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property Registry As Microsoft.VisualBasic.MyServices.RegistryProxy
            Get
                Return My.Computer.Registry
            End Get
        End Property

        ''' <summary>
        ''' Raiz HKEY_CLASSES_ROOT
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property ClassesRoot As RegistryKey
            Get
                Return My.Computer.Registry.ClassesRoot
            End Get
        End Property

        ''' <summary>
        ''' Raiz HKEY_CURRENT_USER
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property CurrentUser As RegistryKey
            Get
                Return My.Computer.Registry.CurrentUser
            End Get
        End Property

        ''' <summary>
        ''' Raiz HKEY_LOCAL_MACHINE
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property LocalMachine As RegistryKey
            Get
                Return My.Computer.Registry.LocalMachine
            End Get
        End Property

        ''' <summary>
        ''' Raiz HKEY_PERFORMANCE_DATA
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property PerformanceData As RegistryKey
            Get
                Return My.Computer.Registry.PerformanceData
            End Get
        End Property

        ''' <summary>
        ''' Raiz HKEY_USERS
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property Users As RegistryKey
            Get
                Return My.Computer.Registry.Users
            End Get
        End Property

        ''' <summary>
        ''' Consulta valor do registro
        ''' </summary>
        ''' <param name="RegistryRoot">Raiz</param>
        ''' <param name="Path">Caminho</param>
        ''' <param name="Key">Chave a ser consultada</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetValue(RegistryRoot As RegistryRoots, Path As String, Key As Object) As Object
            Return My.Computer.Registry.GetValue(RegistryRoot.ToString & "\" & Path, Key, Nothing)
        End Function

        ''' <summary>
        ''' Adicionar valor no registro (Caso o caminho não exista ele será criado)
        ''' </summary>
        ''' <param name="RegistryRoot">Raiz</param>
        ''' <param name="Path">Caminho</param>
        ''' <param name="Key">Chave a ser criada</param>
        ''' <param name="Value">Valor associado a chave</param>
        ''' <param name="Kind">Tipo de dado</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AddValue(RegistryRoot As RegistryRoots, Path As String, Key As Object, Value As Object, Kind As Microsoft.Win32.RegistryValueKind) As Object
            Try
                Select Case RegistryRoot
                    Case RegistryRoots.HKEY_CLASSES_ROOT
                        My.Computer.Registry.ClassesRoot.CreateSubKey(Path).SetValue(Key, Value, Kind)
                    Case RegistryRoots.HKEY_CURRENT_USER
                        My.Computer.Registry.CurrentUser.CreateSubKey(Path).SetValue(Key, Value, Kind)
                    Case RegistryRoots.HKEY_LOCAL_MACHINE
                        My.Computer.Registry.LocalMachine.CreateSubKey(Path).SetValue(Key, Value, Kind)
                    Case RegistryRoots.HKEY_PERFORMANCE_DATA
                        My.Computer.Registry.PerformanceData.CreateSubKey(Path).SetValue(Key, Value, Kind)
                    Case RegistryRoots.HKEY_USERS
                        My.Computer.Registry.Users.CreateSubKey(Path).SetValue(Key, Value, Kind)
                End Select
                Return Value
            Catch ex As System.Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Adicionar caminho no registro 
        ''' </summary>
        ''' <param name="RegistryRoot">Raiz</param>
        ''' <param name="Path">Caminho</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AddPath(RegistryRoot As RegistryRoots, Path As String) As Object
            Try
                Select Case RegistryRoot
                    Case RegistryRoots.HKEY_CLASSES_ROOT
                        My.Computer.Registry.ClassesRoot.CreateSubKey(Path)
                    Case RegistryRoots.HKEY_CURRENT_USER
                        My.Computer.Registry.CurrentUser.CreateSubKey(Path)
                    Case RegistryRoots.HKEY_LOCAL_MACHINE
                        My.Computer.Registry.LocalMachine.CreateSubKey(Path)
                    Case RegistryRoots.HKEY_PERFORMANCE_DATA
                        My.Computer.Registry.PerformanceData.CreateSubKey(Path)
                    Case RegistryRoots.HKEY_USERS
                        My.Computer.Registry.Users.CreateSubKey(Path)
                End Select
                Return Path
            Catch ex As System.Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Deletar caminho e chaves
        ''' </summary>
        ''' <param name="RegistryRoot">Raiz</param>
        ''' <param name="Path">Caminho</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DeletePathAndKeys(RegistryRoot As RegistryRoots, Path As String) As Boolean
            Try
                Dim SubKey As String = ""
                If Path.Contains("\") = True Then
                    Path = IO.Path.GetDirectoryName(Path)
                    SubKey = IO.Path.GetFileNameWithoutExtension(Path)
                    Select Case RegistryRoot
                        Case RegistryRoots.HKEY_CLASSES_ROOT
                            My.Computer.Registry.ClassesRoot.OpenSubKey(Path, True).DeleteSubKeyTree(SubKey)
                        Case RegistryRoots.HKEY_CURRENT_USER
                            My.Computer.Registry.CurrentUser.OpenSubKey(Path, True).DeleteSubKeyTree(SubKey)
                        Case RegistryRoots.HKEY_LOCAL_MACHINE
                            My.Computer.Registry.LocalMachine.OpenSubKey(Path, True).DeleteSubKeyTree(SubKey)
                        Case RegistryRoots.HKEY_PERFORMANCE_DATA
                            My.Computer.Registry.PerformanceData.OpenSubKey(Path, True).DeleteSubKeyTree(SubKey)
                        Case RegistryRoots.HKEY_USERS
                            My.Computer.Registry.Users.OpenSubKey(Path, True).DeleteSubKeyTree(SubKey)
                    End Select
                    Return True
                Else
                    Path = Path
                    SubKey = ""
                    Select Case RegistryRoot
                        Case RegistryRoots.HKEY_CLASSES_ROOT
                            My.Computer.Registry.ClassesRoot.DeleteSubKeyTree(Path, True)
                        Case RegistryRoots.HKEY_CURRENT_USER
                            My.Computer.Registry.CurrentUser.DeleteSubKeyTree(Path, True)
                        Case RegistryRoots.HKEY_LOCAL_MACHINE
                            My.Computer.Registry.LocalMachine.DeleteSubKeyTree(Path, True)
                        Case RegistryRoots.HKEY_PERFORMANCE_DATA
                            My.Computer.Registry.PerformanceData.DeleteSubKeyTree(Path, True)
                        Case RegistryRoots.HKEY_USERS
                            My.Computer.Registry.Users.DeleteSubKeyTree(Path, True)
                    End Select
                    Return True
                End If
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Deletar caminho e chave específica
        ''' </summary>
        ''' <param name="RegistryRoot">Raiz</param>
        ''' <param name="Path">Caminho</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DeleteKey(RegistryRoot As RegistryRoots, Path As String, Key As Object) As Boolean
            Try
                Select Case RegistryRoot
                    Case RegistryRoots.HKEY_CLASSES_ROOT
                        My.Computer.Registry.ClassesRoot.OpenSubKey(Path, True).DeleteValue(Key, False)
                    Case RegistryRoots.HKEY_CURRENT_USER
                        My.Computer.Registry.CurrentUser.OpenSubKey(Path, True).DeleteValue(Key, False)
                    Case RegistryRoots.HKEY_LOCAL_MACHINE
                        My.Computer.Registry.LocalMachine.OpenSubKey(Path, True).DeleteValue(Key, False)
                    Case RegistryRoots.HKEY_PERFORMANCE_DATA
                        My.Computer.Registry.PerformanceData.OpenSubKey(Path, True).DeleteValue(Key, False)
                    Case RegistryRoots.HKEY_USERS
                        My.Computer.Registry.Users.OpenSubKey(Path, True).DeleteValue(Key, False)
                End Select
                Return True
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Retorna a coleção de caminhos
        ''' </summary>
        ''' <param name="RegistryRoot"></param>
        ''' <param name="Path"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetPaths(RegistryRoot As RegistryRoots, Path As String) As String()
            Try
                Select Case RegistryRoot
                    Case RegistryRoots.HKEY_CLASSES_ROOT
                        Return My.Computer.Registry.ClassesRoot.OpenSubKey(Path).GetSubKeyNames()
                    Case RegistryRoots.HKEY_CURRENT_USER
                        Return My.Computer.Registry.CurrentUser.OpenSubKey(Path).GetSubKeyNames()
                    Case RegistryRoots.HKEY_LOCAL_MACHINE
                        Return My.Computer.Registry.LocalMachine.OpenSubKey(Path).GetSubKeyNames()
                    Case RegistryRoots.HKEY_PERFORMANCE_DATA
                        Return My.Computer.Registry.PerformanceData.OpenSubKey(Path).GetSubKeyNames()
                    Case RegistryRoots.HKEY_USERS
                        Return My.Computer.Registry.Users.OpenSubKey(Path).GetSubKeyNames()
                    Case Else
                        Return Nothing
                End Select
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retorna a coleção de chaves
        ''' </summary>
        ''' <param name="RegistryRoot"></param>
        ''' <param name="Path"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetKeys(RegistryRoot As RegistryRoots, Path As String) As String()
            Try
                Select Case RegistryRoot
                    Case RegistryRoots.HKEY_CLASSES_ROOT
                        Return My.Computer.Registry.ClassesRoot.OpenSubKey(Path).GetValueNames()
                    Case RegistryRoots.HKEY_CURRENT_USER
                        Return My.Computer.Registry.CurrentUser.OpenSubKey(Path).GetValueNames()
                    Case RegistryRoots.HKEY_LOCAL_MACHINE
                        Return My.Computer.Registry.LocalMachine.OpenSubKey(Path).GetValueNames()
                    Case RegistryRoots.HKEY_PERFORMANCE_DATA
                        Return My.Computer.Registry.PerformanceData.OpenSubKey(Path).GetValueNames()
                    Case RegistryRoots.HKEY_USERS
                        Return My.Computer.Registry.Users.OpenSubKey(Path).GetValueNames()
                    Case Else
                        Return Nothing
                End Select
            Catch
                Return Nothing
            End Try
        End Function

    End Class

End Namespace
