'=========================================================================================================='
'DESENVOLVIDO POR: LUCAS LEÔNCIO WEY DA SILVA (FORMANDO EM ANALISTA DE SISTEMAS)
'EM:2023
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: wey.lucas1@gmail.com \ (011) 99940-2202
'=========================================================================================================='

Public Class ctrElectricTools

    ''' <summary>
    ''' REGAL
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub cmdREGAL_Click(sender As Object, e As EventArgs) Handles cmdREGAL.Click
        CallCommand.CallCommand("REGAL")
    End Sub

    ''' <summary>
    ''' Abre as configurações de REGAL
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub tsmRegalConfig_Click(sender As Object, e As EventArgs) Handles tsmRegalConfig.Click
        Dim frmRegalConfig As frmRegalConfig = Nothing
        Try
            frmRegalConfig = New frmRegalConfig
            frmRegalConfig.ShowDialog()
        Catch ex As System.Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        Finally
            If IsNothing(frmRegalConfig) = False Then
                frmRegalConfig.Dispose()
            End If
        End Try
    End Sub

End Class
