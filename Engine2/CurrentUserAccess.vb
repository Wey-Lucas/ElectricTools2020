Imports System.Threading
Imports System.Security.Permissions
Imports System.Security
Imports System.Security.Principal
Imports System.IO
Imports System.Security.AccessControl

Namespace Engine2

    ''' <summary>
    ''' Controla a segurança de acesso a arquivos e pastas
    ''' </summary>
    ''' <remarks></remarks>
    Public Class CurrentUserAccess

        ''' <summary>
        ''' Verifica acesso ao diretório
        ''' </summary>
        ''' <param name="DirectoryInfo"></param>
        ''' <param name="FileSystemRights"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function HasAccess(DirectoryInfo As DirectoryInfo, FileSystemRights As FileSystemRights) As Boolean
            Dim AuthorizationRuleCollection As AuthorizationRuleCollection = DirectoryInfo.GetAccessControl().GetAccessRules(True, True, GetType(SecurityIdentifier))
            Return HasFileOrDirectoryAccess(FileSystemRights, AuthorizationRuleCollection)
        End Function

        ''' <summary>
        ''' Verifica acesso ao arquivo
        ''' </summary>
        ''' <param name="FileInfo"></param>
        ''' <param name="FileSystemRights"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function HasAccess(FileInfo As FileInfo, FileSystemRights As FileSystemRights) As Boolean
            Dim AuthorizationRuleCollection As AuthorizationRuleCollection = FileInfo.GetAccessControl().GetAccessRules(True, True, GetType(SecurityIdentifier))
            Return HasFileOrDirectoryAccess(FileSystemRights, AuthorizationRuleCollection)
        End Function

        ''' <summary>
        ''' Função interna de verificação
        ''' </summary>
        ''' <param name="FileSystemRights"></param>
        ''' <param name="AuthorizationRuleCollection"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function HasFileOrDirectoryAccess(FileSystemRights As FileSystemRights, AuthorizationRuleCollection As AuthorizationRuleCollection) As Boolean
            Dim allow As Boolean = False
            Dim inheritedAllow As Boolean = False
            Dim inheritedDeny As Boolean = False
            For i As Integer = 0 To AuthorizationRuleCollection.Count - 1
                Dim currentRule As FileSystemAccessRule = DirectCast(AuthorizationRuleCollection(i), FileSystemAccessRule)
                Dim currentUser As WindowsIdentity = WindowsIdentity.GetCurrent()
                Dim currentPrincipal As WindowsPrincipal = New WindowsPrincipal(WindowsIdentity.GetCurrent())
                If currentUser.User.Equals(currentRule.IdentityReference) OrElse currentPrincipal.IsInRole(DirectCast(currentRule.IdentityReference, SecurityIdentifier)) Then
                    If currentRule.AccessControlType.Equals(AccessControlType.Deny) Then
                        If (currentRule.FileSystemRights And FileSystemRights) = FileSystemRights Then
                            If currentRule.IsInherited Then
                                inheritedDeny = True
                            Else
                                Return False
                            End If
                        End If
                    ElseIf currentRule.AccessControlType.Equals(AccessControlType.Allow) Then
                        If (currentRule.FileSystemRights And FileSystemRights) = FileSystemRights Then
                            If currentRule.IsInherited Then
                                inheritedAllow = True
                            Else
                                allow = True
                            End If
                        End If
                    End If
                End If
            Next
            If allow Then
                Return True
            End If
            Return inheritedAllow AndAlso Not inheritedDeny
        End Function
    End Class

End Namespace

