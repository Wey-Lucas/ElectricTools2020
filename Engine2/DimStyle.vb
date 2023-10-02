Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Gerencia estilos de dimensionamento
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DimStyle

        ''' <summary>
        ''' Obtem o estilo de dimensionamento
        ''' </summary>
        ''' <param name="DimStyleName">Nome do estilo de dimensão</param>
        ''' <returns>DimStyleTableRecord</returns>
        ''' <remarks></remarks>
        Public Shared Function GetDimStyle(DimStyleName As String) As DimStyleTableRecord
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
                Dim DimStyleTableRecord As DimStyleTableRecord
                Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        Dim DimStyleTable As DimStyleTable = CType(Transaction.GetObject(Document.Database.DimStyleTableId, OpenMode.ForRead, False), DimStyleTable)
                        If DimStyleTable.Has(DimStyleName) = False Then
                            Return Nothing
                        End If
                        DimStyleTableRecord = TryCast(Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead), DimStyleTableRecord)
                        Return DimStyleTableRecord
                    End Using
                End Using
                Return Nothing
            Catch ex As System.Exception
                Throw New System.Exception("Erro em DimStyle.GetDimStyle, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Obtem o estilo de dimensionamento corrente
        ''' </summary>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public Shared Function GetCurrent() As String
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim DimStyleTableRecord As DimStyleTableRecord
                Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                    DimStyleTableRecord = CType(Transaction.GetObject(Document.Database.Dimstyle, OpenMode.ForRead), DimStyleTableRecord)
                    Transaction.Dispose()
                End Using
                Return DimStyleTableRecord.Name
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Obtem a coleção de estilos de dimesionamento
        ''' </summary>
        ''' <returns>ArrayList</returns>
        ''' <remarks></remarks>
        Public Shared Function GetNames() As ArrayList
            Try
                GetNames = New ArrayList
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
                Dim DimStyleTableRecord As DimStyleTableRecord
                Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        Dim DimStyleTable As DimStyleTable = CType(Transaction.GetObject(Document.Database.DimStyleTableId, OpenMode.ForRead, False), DimStyleTable)
                        For Each ObjectId As ObjectId In DimStyleTable
                            DimStyleTableRecord = DirectCast(Transaction.GetObject(ObjectId, OpenMode.ForWrite), DimStyleTableRecord)
                            GetNames.Add(DimStyleTableRecord.Name)
                        Next
                    End Using
                End Using
                GetNames.Sort()
                Return GetNames
            Catch ex As System.Exception
                Throw New System.Exception("Erro em DimStyle.GetNames, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Seta o estilo corrente
        ''' </summary>
        ''' <param name="DimStyleName">Nome do estilo de dimensão</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SetCurrent(DimStyleName As String) As Boolean
            Try
                Dim Document As Document = Application.DocumentManager.MdiActiveDocument
                Dim Database As Database = Document.Database
                Dim DimStyleTableRecord As DimStyleTableRecord
                Dim DimStyleTable As DimStyleTable
                Dim ObjectId As ObjectId
                Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        DimStyleTable = DirectCast(Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead), DimStyleTable)
                        If DimStyleTable.Has(DimStyleName) = True Then
                            ObjectId = DimStyleTable(DimStyleName)
                            DimStyleTableRecord = DirectCast(Transaction.GetObject(ObjectId, OpenMode.ForRead), DimStyleTableRecord)
                            Database.Dimstyle = DimStyleTableRecord.ObjectId
                            Database.SetDimstyleData(DimStyleTableRecord)
                            Transaction.Commit()
                            Return True
                        Else
                            Transaction.Abort()
                            Return False
                        End If
                    End Using
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Cria o estilo de dimensionamento
        ''' </summary>
        ''' <param name="DimStyleName">Nome do estilo de dimensão</param>
        ''' <param name="Dimadec">Controla o número de lugares de precisão exibidos em dimensões angulares</param>
        ''' <param name="Dimalt">Controla a exibição de unidades alternativas em dimensões</param>
        ''' <param name="Dimaltd">Controla o número de casas decimais em unidades alternativas</param>
        ''' <param name="Dimaltf">Controla o multiplicador para as unidades alternativas</param>
        ''' <param name="Dimaltrnd">Completa as unidades de dimensão alternativa</param>
        ''' <param name="Dimalttd">Define o número de casas decimais para os valores de tolerância nas unidades alternativas de uma cota</param>
        ''' <param name="Dimalttz">Controla a supressão de zeros nos valores de tolerância</param>
        ''' <param name="Dimaltu">Define o formato de unidades para as unidades alternativas de todos os subestilos de cotas exceto Angular</param>
        ''' <param name="Dimaltz">Controla a supressão de zeros para valores de dimensão de unidade alternativa</param>
        ''' <param name="Dimapost">Especifica um prefixo ou sufixo de texto (ou ambos) para a medição dimensão alternativa para todos os tipos de cotas exceto angular</param>
        ''' <param name="Dimarcsym">Controles de exibição do símbolo do arco em uma dimensão de comprimento de arco</param>
        ''' <param name="Dimasz">Controla o tamanho da linha de dimensão e de linha de pontas de flechas líder. Também controla o tamanho das linhas de gancho</param>
        ''' <param name="Dimatfit">Determina como o texto da cota e setas são organizados quando o espaço não é suficiente para colocar tanto dentro das linhas de extensão</param>
        ''' <param name="Dimaunit">Define o formato de unidades para dimensões angulares</param>
        ''' <param name="Dimazin">Suprime zeros para dimensões angulares</param>
        ''' <param name="Dimblk">Define o bloco de seta exibido nas extremidades das linhas de dimensão</param>
        ''' <param name="Dimblk1">Define a ponta de seta para a primeira extremidade da linha de cota quando DIMSAH está ligado</param>
        ''' <param name="Dimblk2">Define a ponta de seta para a segunda extremidade da linha de cota quando DIMSAH está ligado</param>
        ''' <param name="Dimcen">Controles de desenho de círculo ou centro do arco marcas e linhas de centro pela DIMCENTER, comandos DIMDIAMETER, e DIMRADIUS</param>
        ''' <param name="Dimclrd">Atribui cores para linhas de cota, pontas de flechas, e linhas de chamada de dimensão</param>
        ''' <param name="Dimclre">Atribui cores para linhas de extensão, marcas de centro e linhas de centro</param>
        ''' <param name="Dimclrt">Atribui cores para dimensionar texto</param>
        ''' <param name="Dimdec">Define o número de casas decimais exibidas para as unidades primárias de uma dimensão</param>
        ''' <param name="Dimdle">Define a distância da linha de cota se estende além da linha de extensão quando traços oblíquos são desenhados ao invés de pontas de flechas</param>
        ''' <param name="Dimdli">Controla o espaçamento das linhas de dimensão em dimensões da linha de base</param>
        ''' <param name="Dimdsep">Especifica um separador decimal de caractere simples para usar ao criar cotas cujo formato de unidade é decimal</param>
        ''' <param name="Dimexe">Especifica o quão longe para estender a linha de extensão para além da linha de cota</param>
        ''' <param name="Dimexo">Especifica até que ponto as linhas de extensão são compensados ​​a partir de pontos de origem</param>
        ''' <param name="Dimfrac">Define o formato de fração quando DIMLUNIT está definido para 4 (Arquitetura) ou 5 (fracionário)</param>
        ''' <param name="Dimfxlen">Não documentado</param>
        ''' <param name="DimfxlenOn">Não documentado</param>
        ''' <param name="Dimgap">Define a distância em torno do texto da cota quando a linha de cota quebra para acomodar o texto da cota</param>
        ''' <param name="Dimjogang">Determina o ângulo do segmento transverso da linha de cota em uma cota de raio com desvio</param>
        ''' <param name="Dimjust">Controla o posicionamento horizontal do texto da cota</param>
        ''' <param name="Dimldrblk">Especifica o tipo de seta para linhas de chamadas</param>
        ''' <param name="Dimlfac">Define um fator de escala para medições de cota linear</param>
        ''' <param name="Dimlim">Gera limites de cota como o texto padrão</param>
        ''' <param name="Dimltex1">Define o tipo de linha da primeira linha de extensão</param>
        ''' <param name="Dimltex2">Define o tipo de linha da segunda linha de extensão</param>
        ''' <param name="Dimltype">Define o tipo de linha da linha de cota</param>
        ''' <param name="Dimlunit">Define as unidades para todos os tipos de cota, exceto Angular</param>
        ''' <param name="Dimlwd">Atribui espessura de linha às linhas da cota</param>
        ''' <param name="Dimlwe">Atribui espessura de linha às linhas de extensão</param>
        ''' <param name="Dimpost">Especifica um prefixo ou sufixo de texto (ou ambos) para a medição da cota</param>
        ''' <param name="Dimrnd">Arredonda todas as distâncias de cotagem para o valor especificado</param>
        ''' <param name="Dimsah">Controla a exibição dos blocos de ponta de seta de linha de cota</param>
        ''' <param name="Dimscale">Define o fator de escala geral aplicado a variáveis de cotagem que especificam tamanhos, distâncias ou deslocamentos</param>
        ''' <param name="Dimsd1">Controla a supressão da primeira linha e da ponta de seta de cota</param>
        ''' <param name="Dimsd2">Controla a supressão da segunda linha e ponta de seta de cota</param>
        ''' <param name="Dimse1">Suprime a exibição da primeira linha de extensão</param>
        ''' <param name="Dimse2">Suprime a exibição da segunda linha de extensão</param>
        ''' <param name="Dimsoxd">Suprime pontas de seta, se não houver espaço suficiente dentro das linhas de extensão</param>
        ''' <param name="Dimtad">Controla o posicionamento vertical do texto em relação à linha de cota</param>
        ''' <param name="Dimtdec">Define o número de casas decimais exibidas em valores de tolerância para as unidades principais em uma cota</param>
        ''' <param name="Dimtfac">Especifica um fator de escala para a altura do texto de valores de frações e tolerância relativo à altura do texto da cota, como definido por DIMTXT</param>
        ''' <param name="Dimtfill">Controla o plano de fundo do texto da cota</param>
        ''' <param name="Dimtfillclr">Define a cor para o plano de fundo do texto nas cotas</param>
        ''' <param name="Dimtih">Controla a posição do texto da cota dentro das linhas de extensão para todos os tipos de cotas exceto Ordinate</param>
        ''' <param name="Dimtix">Desenha o texto entre as linhas de extensão</param>
        ''' <param name="Dimtm">Define o limite mínimo (ou inferior) de tolerância para o texto da cota quando DIMTOL ou DIMLIM está ativada</param>
        ''' <param name="Dimtmove">Define as regras de movimento de texto da cota</param>
        ''' <param name="Dimtofl">Controla se uma linha de cota é desenhada entre as linhas de extensão, mesmo quando o texto é colocado externamente</param>
        ''' <param name="Dimtoh">Controla a posição do texto da cota fora das linhas de extensão</param>
        ''' <param name="Dimtol">Anexa tolerâncias ao texto da cota</param>
        ''' <param name="Dimtolj">Define a justificação vertical para valores de tolerância relativa ao texto de cota nominal</param>
        ''' <param name="Dimtp">Define o limite máximo (ou superior) de tolerância para o texto da cota quando DIMTOL ou DIMLIM está ativada</param>
        ''' <param name="Dimtsz">Especifica o tamanho de riscos oblíquos desenhados ao invés de pontas de seta para a cotagem linear, radial e de diâmetro</param>
        ''' <param name="Dimtvp">Controla o posicionamento vertical do texto da cota acima ou abaixo da linha de cota</param>
        ''' <param name="Dimtxsty">Especifica a o estilo de texto para a cota</param>
        ''' <param name="Dimtxt">Especifica a altura do texto da cota, a não ser que o estilo de texto atual tenha uma altura fixa</param>
        ''' <param name="Dimtxtdirection">Especifica a direção de leitura do texto da cota</param>
        ''' <param name="Dimtzin">Controla a supressão de zeros em valores de tolerância</param>
        ''' <param name="Dimupt">Controla as opções para texto posicionado pelo usuário</param>
        ''' <param name="Dimzin">Controla a supressão de zeros no valor da unidade principal</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Create(ByVal DimStyleName As String, Optional Dimadec As Object = Nothing, Optional Dimalt As Object = Nothing, Optional Dimaltd As Object = Nothing, Optional Dimaltf As Object = Nothing, Optional Dimaltrnd As Object = Nothing, Optional Dimalttd As Object = Nothing, Optional Dimalttz As Object = Nothing, Optional Dimaltu As Object = Nothing, Optional Dimaltz As Object = Nothing, Optional Dimapost As Object = Nothing, Optional Dimarcsym As Object = Nothing, Optional Dimasz As Object = Nothing, Optional Dimatfit As Object = Nothing, Optional Dimaunit As Object = Nothing, Optional Dimazin As Object = Nothing, Optional Dimblk As Object = Nothing, Optional Dimblk1 As Object = Nothing, Optional Dimblk2 As Object = Nothing, Optional Dimcen As Object = Nothing, Optional Dimclrd As Object = Nothing, Optional Dimclre As Object = Nothing, Optional Dimclrt As Object = Nothing, Optional Dimdec As Object = Nothing, Optional Dimdle As Object = Nothing, Optional Dimdli As Object = Nothing, Optional Dimdsep As Object = Nothing, Optional Dimexe As Object = Nothing, Optional Dimexo As Object = Nothing, Optional Dimfrac As Object = Nothing, Optional Dimfxlen As Object = Nothing, Optional DimfxlenOn As Object = Nothing, Optional Dimgap As Object = Nothing, Optional Dimjogang As Object = Nothing, Optional Dimjust As Object = Nothing, Optional Dimldrblk As Object = Nothing, Optional Dimlfac As Object = Nothing, Optional Dimlim As Object = Nothing, Optional Dimltex1 As Object = Nothing, Optional Dimltex2 As Object = Nothing, Optional Dimltype As Object = Nothing, Optional Dimlunit As Object = Nothing, Optional Dimlwd As Object = Nothing, Optional Dimlwe As Object = Nothing, Optional Dimpost As Object = Nothing, Optional Dimrnd As Object = Nothing, Optional Dimsah As Object = Nothing, Optional Dimscale As Object = Nothing, Optional Dimsd1 As Object = Nothing, Optional Dimsd2 As Object = Nothing, Optional Dimse1 As Object = Nothing, Optional Dimse2 As Object = Nothing, Optional Dimsoxd As Object = Nothing, Optional Dimtad As Object = Nothing, Optional Dimtdec As Object = Nothing, Optional Dimtfac As Object = Nothing, Optional Dimtfill As Object = Nothing, Optional Dimtfillclr As Object = Nothing, Optional Dimtih As Object = Nothing, Optional Dimtix As Object = Nothing, Optional Dimtm As Object = Nothing, Optional Dimtmove As Object = Nothing, Optional Dimtofl As Object = Nothing, Optional Dimtoh As Object = Nothing, Optional Dimtol As Object = Nothing, Optional Dimtolj As Object = Nothing, Optional Dimtp As Object = Nothing, Optional Dimtsz As Object = Nothing, Optional Dimtvp As Object = Nothing, Optional Dimtxsty As Object = Nothing, Optional Dimtxt As Object = Nothing, Optional Dimtxtdirection As Object = Nothing, Optional Dimtzin As Object = Nothing, Optional Dimupt As Object = Nothing, Optional Dimzin As Object = Nothing) As Boolean
            Try
                Dim Database As Database = Application.DocumentManager.MdiActiveDocument.Database
                Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        Dim DimStyleTable As DimStyleTable = DirectCast(Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead), DimStyleTable)
                        If Not DimStyleTable.Has(DimStyleName) Then
                            DimStyleTable.UpgradeOpen()
                            Dim DimStyleTableRecord As New DimStyleTableRecord()
                            With DimStyleTableRecord
                                .Name = DimStyleName
                                If IsNothing(Dimadec) = False Then
                                    .Dimadec = Dimadec
                                End If
                                If IsNothing(Dimalt) = False Then
                                    .Dimalt = Dimalt
                                End If
                                If IsNothing(Dimaltd) = False Then
                                    .Dimaltd = Dimaltd
                                End If
                                If IsNothing(Dimaltf) = False Then
                                    .Dimaltf = Dimaltf
                                End If
                                If IsNothing(Dimaltrnd) = False Then
                                    .Dimaltrnd = Dimaltrnd
                                End If
                                If IsNothing(Dimalttd) = False Then
                                    .Dimalttd = Dimalttd
                                End If
                                If IsNothing(Dimalttz) = False Then
                                    .Dimalttz = Dimalttz
                                End If
                                If IsNothing(Dimaltu) = False Then
                                    .Dimaltu = Dimaltu
                                End If
                                If IsNothing(Dimaltz) = False Then
                                    .Dimaltz = Dimaltz
                                End If
                                If IsNothing(Dimapost) = False Then
                                    .Dimapost = Dimapost
                                End If
                                If IsNothing(Dimarcsym) = False Then
                                    .Dimarcsym = Dimarcsym
                                End If
                                If IsNothing(Dimasz) = False Then
                                    .Dimasz = Dimasz
                                End If
                                If IsNothing(Dimatfit) = False Then
                                    .Dimatfit = Dimatfit
                                End If
                                If IsNothing(Dimaunit) = False Then
                                    .Dimaunit = Dimaunit
                                End If
                                If IsNothing(Dimazin) = False Then
                                    .Dimazin = Dimazin
                                End If
                                If IsNothing(Dimblk) = False Then
                                    .Dimblk = Dimblk
                                End If
                                If IsNothing(Dimblk1) = False Then
                                    .Dimblk1 = Dimblk1
                                End If
                                If IsNothing(Dimblk2) = False Then
                                    .Dimblk2 = Dimblk2
                                End If
                                If IsNothing(Dimcen) = False Then
                                    .Dimcen = Dimcen
                                End If
                                If IsNothing(Dimclrd) = False Then
                                    .Dimclrd = Dimclrd
                                End If
                                If IsNothing(Dimclre) = False Then
                                    .Dimclre = Dimclre
                                End If
                                If IsNothing(Dimclrt) = False Then
                                    .Dimclrt = Dimclrt
                                End If
                                If IsNothing(Dimdec) = False Then
                                    .Dimdec = Dimdec
                                End If
                                If IsNothing(Dimdle) = False Then
                                    .Dimdle = Dimdle
                                End If
                                If IsNothing(Dimdli) = False Then
                                    .Dimdli = Dimdli
                                End If
                                If IsNothing(Dimdsep) = False Then
                                    .Dimdsep = Dimdsep
                                End If
                                If IsNothing(Dimexe) = False Then
                                    .Dimexe = Dimexe
                                End If
                                If IsNothing(Dimexo) = False Then
                                    .Dimexo = Dimexo
                                End If
                                If IsNothing(Dimfrac) = False Then
                                    .Dimfrac = Dimfrac
                                End If
                                If IsNothing(Dimfxlen) = False Then
                                    .Dimfxlen = Dimfxlen
                                End If
                                If IsNothing(DimfxlenOn) = False Then
                                    .DimfxlenOn = DimfxlenOn
                                End If
                                If IsNothing(Dimgap) = False Then
                                    .Dimgap = Dimgap
                                End If
                                If IsNothing(Dimjogang) = False Then
                                    .Dimjogang = Dimjogang
                                End If
                                If IsNothing(Dimjust) = False Then
                                    .Dimjust = Dimjust
                                End If
                                If IsNothing(Dimldrblk) = False Then
                                    .Dimldrblk = Dimldrblk
                                End If
                                If IsNothing(Dimlfac) = False Then
                                    .Dimlfac = Dimlfac
                                End If
                                If IsNothing(Dimlim) = False Then
                                    .Dimlim = Dimlim
                                End If
                                If IsNothing(Dimltex1) = False Then
                                    .Dimltex1 = Dimltex1
                                End If
                                If IsNothing(Dimltex2) = False Then
                                    .Dimltex2 = Dimltex2
                                End If
                                If IsNothing(Dimltype) = False Then
                                    .Dimltype = Dimltype
                                End If
                                If IsNothing(Dimlunit) = False Then
                                    .Dimlunit = Dimlunit
                                End If
                                If IsNothing(Dimlwd) = False Then
                                    .Dimlwd = Dimlwd
                                End If
                                If IsNothing(Dimlwe) = False Then
                                    .Dimlwe = Dimlwe
                                End If
                                If IsNothing(Dimpost) = False Then
                                    .Dimpost = Dimpost
                                End If
                                If IsNothing(Dimrnd) = False Then
                                    .Dimrnd = Dimrnd
                                End If
                                If IsNothing(Dimsah) = False Then
                                    .Dimsah = Dimsah
                                End If
                                If IsNothing(Dimscale) = False Then
                                    .Dimscale = Dimscale
                                End If
                                If IsNothing(Dimsd1) = False Then
                                    .Dimsd1 = Dimsd1
                                End If
                                If IsNothing(Dimsd2) = False Then
                                    .Dimsd2 = Dimsd2
                                End If
                                If IsNothing(Dimse1) = False Then
                                    .Dimse1 = Dimse1
                                End If
                                If IsNothing(Dimse2) = False Then
                                    .Dimse2 = Dimse2
                                End If
                                If IsNothing(Dimsoxd) = False Then
                                    .Dimsoxd = Dimsoxd
                                End If
                                If IsNothing(Dimtad) = False Then
                                    .Dimtad = Dimtad
                                End If
                                If IsNothing(Dimtdec) = False Then
                                    .Dimtdec = Dimtdec
                                End If
                                If IsNothing(Dimtfac) = False Then
                                    .Dimtfac = Dimtfac
                                End If
                                If IsNothing(Dimtfill) = False Then
                                    .Dimtfill = Dimtfill
                                End If
                                If IsNothing(Dimtfillclr) = False Then
                                    .Dimtfillclr = Dimtfillclr
                                End If
                                If IsNothing(Dimtih) = False Then
                                    .Dimtih = Dimtih
                                End If
                                If IsNothing(Dimtix) = False Then
                                    .Dimtix = Dimtix
                                End If
                                If IsNothing(Dimtm) = False Then
                                    .Dimtm = Dimtm
                                End If
                                If IsNothing(Dimtmove) = False Then
                                    .Dimtmove = Dimtmove
                                End If
                                If IsNothing(Dimtofl) = False Then
                                    .Dimtofl = Dimtofl
                                End If
                                If IsNothing(Dimtoh) = False Then
                                    .Dimtoh = Dimtoh
                                End If
                                If IsNothing(Dimtol) = False Then
                                    .Dimtol = Dimtol
                                End If
                                If IsNothing(Dimtolj) = False Then
                                    .Dimtolj = Dimtolj
                                End If
                                If IsNothing(Dimtp) = False Then
                                    .Dimtp = Dimtp
                                End If
                                If IsNothing(Dimtsz) = False Then
                                    .Dimtsz = Dimtsz
                                End If
                                If IsNothing(Dimtvp) = False Then
                                    .Dimtvp = Dimtvp
                                End If
                                If IsNothing(Dimtxsty) = False Then
                                    .Dimtxsty = Dimtxsty
                                End If
                                If IsNothing(Dimtxt) = False Then
                                    .Dimtxt = Dimtxt
                                End If
                                If IsNothing(Dimtxtdirection) = False Then
                                    .Dimtxtdirection = Dimtxtdirection
                                End If
                                If IsNothing(Dimtzin) = False Then
                                    .Dimtzin = Dimtzin
                                End If
                                If IsNothing(Dimupt) = False Then
                                    .Dimupt = Dimupt
                                End If
                                If IsNothing(Dimzin) = False Then
                                    .Dimzin = Dimzin
                                End If
                            End With
                            DimStyleTable.Add(DimStyleTableRecord)
                            Transaction.AddNewlyCreatedDBObject(DimStyleTableRecord, True)
                        End If
                        Transaction.Commit()
                    End Using
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function


        ''' <summary>
        ''' Modifica o dimensionamento
        ''' </summary>
        ''' <param name="DimStyleName">Nome do estilo de dimensão</param>
        ''' <param name="Dimadec">Controla o número de lugares de precisão exibidos em dimensões angulares</param>
        ''' <param name="Dimalt">Controla a exibição de unidades alternativas em dimensões</param>
        ''' <param name="Dimaltd">Controla o número de casas decimais em unidades alternativas</param>
        ''' <param name="Dimaltf">Controla o multiplicador para as unidades alternativas</param>
        ''' <param name="Dimaltrnd">Completa as unidades de dimensão alternativa</param>
        ''' <param name="Dimalttd">Define o número de casas decimais para os valores de tolerância nas unidades alternativas de uma cota</param>
        ''' <param name="Dimalttz">Controla a supressão de zeros nos valores de tolerância</param>
        ''' <param name="Dimaltu">Define o formato de unidades para as unidades alternativas de todos os subestilos de cotas exceto Angular</param>
        ''' <param name="Dimaltz">Controla a supressão de zeros para valores de dimensão de unidade alternativa</param>
        ''' <param name="Dimapost">Especifica um prefixo ou sufixo de texto (ou ambos) para a medição dimensão alternativa para todos os tipos de cotas exceto angular</param>
        ''' <param name="Dimarcsym">Controles de exibição do símbolo do arco em uma dimensão de comprimento de arco</param>
        ''' <param name="Dimasz">Controla o tamanho da linha de dimensão e de linha de pontas de flechas líder. Também controla o tamanho das linhas de gancho</param>
        ''' <param name="Dimatfit">Determina como o texto da cota e setas são organizados quando o espaço não é suficiente para colocar tanto dentro das linhas de extensão</param>
        ''' <param name="Dimaunit">Define o formato de unidades para dimensões angulares</param>
        ''' <param name="Dimazin">Suprime zeros para dimensões angulares</param>
        ''' <param name="Dimblk">Define o bloco de seta exibido nas extremidades das linhas de dimensão</param>
        ''' <param name="Dimblk1">Define a ponta de seta para a primeira extremidade da linha de cota quando DIMSAH está ligado</param>
        ''' <param name="Dimblk2">Define a ponta de seta para a segunda extremidade da linha de cota quando DIMSAH está ligado</param>
        ''' <param name="Dimcen">Controles de desenho de círculo ou centro do arco marcas e linhas de centro pela DIMCENTER, comandos DIMDIAMETER, e DIMRADIUS</param>
        ''' <param name="Dimclrd">Atribui cores para linhas de cota, pontas de flechas, e linhas de chamada de dimensão</param>
        ''' <param name="Dimclre">Atribui cores para linhas de extensão, marcas de centro e linhas de centro</param>
        ''' <param name="Dimclrt">Atribui cores para dimensionar texto</param>
        ''' <param name="Dimdec">Define o número de casas decimais exibidas para as unidades primárias de uma dimensão</param>
        ''' <param name="Dimdle">Define a distância da linha de cota se estende além da linha de extensão quando traços oblíquos são desenhados ao invés de pontas de flechas</param>
        ''' <param name="Dimdli">Controla o espaçamento das linhas de dimensão em dimensões da linha de base</param>
        ''' <param name="Dimdsep">Especifica um separador decimal de caractere simples para usar ao criar cotas cujo formato de unidade é decimal</param>
        ''' <param name="Dimexe">Especifica o quão longe para estender a linha de extensão para além da linha de cota</param>
        ''' <param name="Dimexo">Especifica até que ponto as linhas de extensão são compensados ​​a partir de pontos de origem</param>
        ''' <param name="Dimfrac">Define o formato de fração quando DIMLUNIT está definido para 4 (Arquitetura) ou 5 (fracionário)</param>
        ''' <param name="Dimfxlen">Não documentado</param>
        ''' <param name="DimfxlenOn">Não documentado</param>
        ''' <param name="Dimgap">Define a distância em torno do texto da cota quando a linha de cota quebra para acomodar o texto da cota</param>
        ''' <param name="Dimjogang">Determina o ângulo do segmento transverso da linha de cota em uma cota de raio com desvio</param>
        ''' <param name="Dimjust">Controla o posicionamento horizontal do texto da cota</param>
        ''' <param name="Dimldrblk">Especifica o tipo de seta para linhas de chamadas</param>
        ''' <param name="Dimlfac">Define um fator de escala para medições de cota linear</param>
        ''' <param name="Dimlim">Gera limites de cota como o texto padrão</param>
        ''' <param name="Dimltex1">Define o tipo de linha da primeira linha de extensão</param>
        ''' <param name="Dimltex2">Define o tipo de linha da segunda linha de extensão</param>
        ''' <param name="Dimltype">Define o tipo de linha da linha de cota</param>
        ''' <param name="Dimlunit">Define as unidades para todos os tipos de cota, exceto Angular</param>
        ''' <param name="Dimlwd">Atribui espessura de linha às linhas da cota</param>
        ''' <param name="Dimlwe">Atribui espessura de linha às linhas de extensão</param>
        ''' <param name="Dimpost">Especifica um prefixo ou sufixo de texto (ou ambos) para a medição da cota</param>
        ''' <param name="Dimrnd">Arredonda todas as distâncias de cotagem para o valor especificado</param>
        ''' <param name="Dimsah">Controla a exibição dos blocos de ponta de seta de linha de cota</param>
        ''' <param name="Dimscale">Define o fator de escala geral aplicado a variáveis de cotagem que especificam tamanhos, distâncias ou deslocamentos</param>
        ''' <param name="Dimsd1">Controla a supressão da primeira linha e da ponta de seta de cota</param>
        ''' <param name="Dimsd2">Controla a supressão da segunda linha e ponta de seta de cota</param>
        ''' <param name="Dimse1">Suprime a exibição da primeira linha de extensão</param>
        ''' <param name="Dimse2">Suprime a exibição da segunda linha de extensão</param>
        ''' <param name="Dimsoxd">Suprime pontas de seta, se não houver espaço suficiente dentro das linhas de extensão</param>
        ''' <param name="Dimtad">Controla o posicionamento vertical do texto em relação à linha de cota</param>
        ''' <param name="Dimtdec">Define o número de casas decimais exibidas em valores de tolerância para as unidades principais em uma cota</param>
        ''' <param name="Dimtfac">Especifica um fator de escala para a altura do texto de valores de frações e tolerância relativo à altura do texto da cota, como definido por DIMTXT</param>
        ''' <param name="Dimtfill">Controla o plano de fundo do texto da cota</param>
        ''' <param name="Dimtfillclr">Define a cor para o plano de fundo do texto nas cotas</param>
        ''' <param name="Dimtih">Controla a posição do texto da cota dentro das linhas de extensão para todos os tipos de cotas exceto Ordinate</param>
        ''' <param name="Dimtix">Desenha o texto entre as linhas de extensão</param>
        ''' <param name="Dimtm">Define o limite mínimo (ou inferior) de tolerância para o texto da cota quando DIMTOL ou DIMLIM está ativada</param>
        ''' <param name="Dimtmove">Define as regras de movimento de texto da cota</param>
        ''' <param name="Dimtofl">Controla se uma linha de cota é desenhada entre as linhas de extensão, mesmo quando o texto é colocado externamente</param>
        ''' <param name="Dimtoh">Controla a posição do texto da cota fora das linhas de extensão</param>
        ''' <param name="Dimtol">Anexa tolerâncias ao texto da cota</param>
        ''' <param name="Dimtolj">Define a justificação vertical para valores de tolerância relativa ao texto de cota nominal</param>
        ''' <param name="Dimtp">Define o limite máximo (ou superior) de tolerância para o texto da cota quando DIMTOL ou DIMLIM está ativada</param>
        ''' <param name="Dimtsz">Especifica o tamanho de riscos oblíquos desenhados ao invés de pontas de seta para a cotagem linear, radial e de diâmetro</param>
        ''' <param name="Dimtvp">Controla o posicionamento vertical do texto da cota acima ou abaixo da linha de cota</param>
        ''' <param name="Dimtxt">Especifica a altura do texto da cota, a não ser que o estilo de texto atual tenha uma altura fixa</param>
        ''' <param name="Dimtxtdirection">Especifica a direção de leitura do texto da cota</param>
        ''' <param name="Dimtzin">Controla a supressão de zeros em valores de tolerância</param>
        ''' <param name="Dimupt">Controla as opções para texto posicionado pelo usuário</param>
        ''' <param name="Dimzin">Controla a supressão de zeros no valor da unidade principal</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Change(ByVal DimStyleName As String, Optional Dimadec As Object = Nothing, Optional Dimalt As Object = Nothing, Optional Dimaltd As Object = Nothing, Optional Dimaltf As Object = Nothing, Optional Dimaltrnd As Object = Nothing, Optional Dimalttd As Object = Nothing, Optional Dimalttz As Object = Nothing, Optional Dimaltu As Object = Nothing, Optional Dimaltz As Object = Nothing, Optional Dimapost As Object = Nothing, Optional Dimarcsym As Object = Nothing, Optional Dimasz As Object = Nothing, Optional Dimatfit As Object = Nothing, Optional Dimaunit As Object = Nothing, Optional Dimazin As Object = Nothing, Optional Dimblk As Object = Nothing, Optional Dimblk1 As Object = Nothing, Optional Dimblk2 As Object = Nothing, Optional Dimcen As Object = Nothing, Optional Dimclrd As Object = Nothing, Optional Dimclre As Object = Nothing, Optional Dimclrt As Object = Nothing, Optional Dimdec As Object = Nothing, Optional Dimdle As Object = Nothing, Optional Dimdli As Object = Nothing, Optional Dimdsep As Object = Nothing, Optional Dimexe As Object = Nothing, Optional Dimexo As Object = Nothing, Optional Dimfrac As Object = Nothing, Optional Dimfxlen As Object = Nothing, Optional DimfxlenOn As Object = Nothing, Optional Dimgap As Object = Nothing, Optional Dimjogang As Object = Nothing, Optional Dimjust As Object = Nothing, Optional Dimldrblk As Object = Nothing, Optional Dimlfac As Object = Nothing, Optional Dimlim As Object = Nothing, Optional Dimltex1 As Object = Nothing, Optional Dimltex2 As Object = Nothing, Optional Dimltype As Object = Nothing, Optional Dimlunit As Object = Nothing, Optional Dimlwd As Object = Nothing, Optional Dimlwe As Object = Nothing, Optional Dimpost As Object = Nothing, Optional Dimrnd As Object = Nothing, Optional Dimsah As Object = Nothing, Optional Dimscale As Object = Nothing, Optional Dimsd1 As Object = Nothing, Optional Dimsd2 As Object = Nothing, Optional Dimse1 As Object = Nothing, Optional Dimse2 As Object = Nothing, Optional Dimsoxd As Object = Nothing, Optional Dimtad As Object = Nothing, Optional Dimtdec As Object = Nothing, Optional Dimtfac As Object = Nothing, Optional Dimtfill As Object = Nothing, Optional Dimtfillclr As Object = Nothing, Optional Dimtih As Object = Nothing, Optional Dimtix As Object = Nothing, Optional Dimtm As Object = Nothing, Optional Dimtmove As Object = Nothing, Optional Dimtofl As Object = Nothing, Optional Dimtoh As Object = Nothing, Optional Dimtol As Object = Nothing, Optional Dimtolj As Object = Nothing, Optional Dimtp As Object = Nothing, Optional Dimtsz As Object = Nothing, Optional Dimtvp As Object = Nothing, Optional Dimtxt As Object = Nothing, Optional Dimtxtdirection As Object = Nothing, Optional Dimtzin As Object = Nothing, Optional Dimupt As Object = Nothing, Optional Dimzin As Object = Nothing) As Boolean
            Try
                Dim Database As Database = Application.DocumentManager.MdiActiveDocument.Database
                Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                    Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                        Dim DimStyleTable As DimStyleTable = DirectCast(Transaction.GetObject(Database.DimStyleTableId, OpenMode.ForRead), DimStyleTable)
                        If DimStyleTable.Has(DimStyleName) = False Then
                            Dim DimStyleTableRecord As DimStyleTableRecord = Transaction.GetObject(DimStyleTable(DimStyleName), OpenMode.ForRead)
                            Dim ObjectIdCollection As ObjectIdCollection = DimStyleTableRecord.GetPersistentReactorIds()
                            For Each ObjectId As ObjectId In ObjectIdCollection
                                If ObjectId.ObjectClass.IsDerivedFrom(RXClass.GetClass(GetType(Dimension))) = True Then
                                    Dim Dimension As Dimension = Transaction.GetObject(ObjectId, OpenMode.ForWrite)
                                    With Dimension
                                        If IsNothing(Dimadec) = False Then
                                            .Dimadec = Dimadec
                                        End If
                                        If IsNothing(Dimalt) = False Then
                                            .Dimalt = Dimalt
                                        End If
                                        If IsNothing(Dimaltd) = False Then
                                            .Dimaltd = Dimaltd
                                        End If
                                        If IsNothing(Dimaltf) = False Then
                                            .Dimaltf = Dimaltf
                                        End If
                                        If IsNothing(Dimaltrnd) = False Then
                                            .Dimaltrnd = Dimaltrnd
                                        End If
                                        If IsNothing(Dimalttd) = False Then
                                            .Dimalttd = Dimalttd
                                        End If
                                        If IsNothing(Dimalttz) = False Then
                                            .Dimalttz = Dimalttz
                                        End If
                                        If IsNothing(Dimaltu) = False Then
                                            .Dimaltu = Dimaltu
                                        End If
                                        If IsNothing(Dimaltz) = False Then
                                            .Dimaltz = Dimaltz
                                        End If
                                        If IsNothing(Dimapost) = False Then
                                            .Dimapost = Dimapost
                                        End If
                                        If IsNothing(Dimarcsym) = False Then
                                            .Dimarcsym = Dimarcsym
                                        End If
                                        If IsNothing(Dimasz) = False Then
                                            .Dimasz = Dimasz
                                        End If
                                        If IsNothing(Dimatfit) = False Then
                                            .Dimatfit = Dimatfit
                                        End If
                                        If IsNothing(Dimaunit) = False Then
                                            .Dimaunit = Dimaunit
                                        End If
                                        If IsNothing(Dimazin) = False Then
                                            .Dimazin = Dimazin
                                        End If
                                        If IsNothing(Dimblk) = False Then
                                            .Dimblk = Dimblk
                                        End If
                                        If IsNothing(Dimblk1) = False Then
                                            .Dimblk1 = Dimblk1
                                        End If
                                        If IsNothing(Dimblk2) = False Then
                                            .Dimblk2 = Dimblk2
                                        End If
                                        If IsNothing(Dimcen) = False Then
                                            .Dimcen = Dimcen
                                        End If
                                        If IsNothing(Dimclrd) = False Then
                                            .Dimclrd = Dimclrd
                                        End If
                                        If IsNothing(Dimclre) = False Then
                                            .Dimclre = Dimclre
                                        End If
                                        If IsNothing(Dimclrt) = False Then
                                            .Dimclrt = Dimclrt
                                        End If
                                        If IsNothing(Dimdec) = False Then
                                            .Dimdec = Dimdec
                                        End If
                                        If IsNothing(Dimdle) = False Then
                                            .Dimdle = Dimdle
                                        End If
                                        If IsNothing(Dimdli) = False Then
                                            .Dimdli = Dimdli
                                        End If
                                        If IsNothing(Dimdsep) = False Then
                                            .Dimdsep = Dimdsep
                                        End If
                                        If IsNothing(Dimexe) = False Then
                                            .Dimexe = Dimexe
                                        End If
                                        If IsNothing(Dimexo) = False Then
                                            .Dimexo = Dimexo
                                        End If
                                        If IsNothing(Dimfrac) = False Then
                                            .Dimfrac = Dimfrac
                                        End If
                                        If IsNothing(Dimfxlen) = False Then
                                            .Dimfxlen = Dimfxlen
                                        End If
                                        If IsNothing(DimfxlenOn) = False Then
                                            .DimfxlenOn = DimfxlenOn
                                        End If
                                        If IsNothing(Dimgap) = False Then
                                            .Dimgap = Dimgap
                                        End If
                                        If IsNothing(Dimjogang) = False Then
                                            .Dimjogang = Dimjogang
                                        End If
                                        If IsNothing(Dimjust) = False Then
                                            .Dimjust = Dimjust
                                        End If
                                        If IsNothing(Dimldrblk) = False Then
                                            .Dimldrblk = Dimldrblk
                                        End If
                                        If IsNothing(Dimlfac) = False Then
                                            .Dimlfac = Dimlfac
                                        End If
                                        If IsNothing(Dimlim) = False Then
                                            .Dimlim = Dimlim
                                        End If
                                        If IsNothing(Dimltex1) = False Then
                                            .Dimltex1 = Dimltex1
                                        End If
                                        If IsNothing(Dimltex2) = False Then
                                            .Dimltex2 = Dimltex2
                                        End If
                                        If IsNothing(Dimltype) = False Then
                                            .Dimltype = Dimltype
                                        End If
                                        If IsNothing(Dimlunit) = False Then
                                            .Dimlunit = Dimlunit
                                        End If
                                        If IsNothing(Dimlwd) = False Then
                                            .Dimlwd = Dimlwd
                                        End If
                                        If IsNothing(Dimlwe) = False Then
                                            .Dimlwe = Dimlwe
                                        End If
                                        If IsNothing(Dimpost) = False Then
                                            .Dimpost = Dimpost
                                        End If
                                        If IsNothing(Dimrnd) = False Then
                                            .Dimrnd = Dimrnd
                                        End If
                                        If IsNothing(Dimsah) = False Then
                                            .Dimsah = Dimsah
                                        End If
                                        If IsNothing(Dimscale) = False Then
                                            .Dimscale = Dimscale
                                        End If
                                        If IsNothing(Dimsd1) = False Then
                                            .Dimsd1 = Dimsd1
                                        End If
                                        If IsNothing(Dimsd2) = False Then
                                            .Dimsd2 = Dimsd2
                                        End If
                                        If IsNothing(Dimse1) = False Then
                                            .Dimse1 = Dimse1
                                        End If
                                        If IsNothing(Dimse2) = False Then
                                            .Dimse2 = Dimse2
                                        End If
                                        If IsNothing(Dimsoxd) = False Then
                                            .Dimsoxd = Dimsoxd
                                        End If
                                        If IsNothing(Dimtad) = False Then
                                            .Dimtad = Dimtad
                                        End If
                                        If IsNothing(Dimtdec) = False Then
                                            .Dimtdec = Dimtdec
                                        End If
                                        If IsNothing(Dimtfac) = False Then
                                            .Dimtfac = Dimtfac
                                        End If
                                        If IsNothing(Dimtfill) = False Then
                                            .Dimtfill = Dimtfill
                                        End If
                                        If IsNothing(Dimtfillclr) = False Then
                                            .Dimtfillclr = Dimtfillclr
                                        End If
                                        If IsNothing(Dimtih) = False Then
                                            .Dimtih = Dimtih
                                        End If
                                        If IsNothing(Dimtix) = False Then
                                            .Dimtix = Dimtix
                                        End If
                                        If IsNothing(Dimtm) = False Then
                                            .Dimtm = Dimtm
                                        End If
                                        If IsNothing(Dimtmove) = False Then
                                            .Dimtmove = Dimtmove
                                        End If
                                        If IsNothing(Dimtofl) = False Then
                                            .Dimtofl = Dimtofl
                                        End If
                                        If IsNothing(Dimtoh) = False Then
                                            .Dimtoh = Dimtoh
                                        End If
                                        If IsNothing(Dimtol) = False Then
                                            .Dimtol = Dimtol
                                        End If
                                        If IsNothing(Dimtolj) = False Then
                                            .Dimtolj = Dimtolj
                                        End If
                                        If IsNothing(Dimtp) = False Then
                                            .Dimtp = Dimtp
                                        End If
                                        If IsNothing(Dimtsz) = False Then
                                            .Dimtsz = Dimtsz
                                        End If
                                        If IsNothing(Dimtvp) = False Then
                                            .Dimtvp = Dimtvp
                                        End If
                                        If IsNothing(Dimtxt) = False Then
                                            .Dimtxt = Dimtxt
                                        End If
                                        If IsNothing(Dimtxtdirection) = False Then
                                            .Dimtxtdirection = Dimtxtdirection
                                        End If
                                        If IsNothing(Dimtzin) = False Then
                                            .Dimtzin = Dimtzin
                                        End If
                                        If IsNothing(Dimupt) = False Then
                                            .Dimupt = Dimupt
                                        End If
                                        If IsNothing(Dimzin) = False Then
                                            .Dimzin = Dimzin
                                        End If
                                    End With
                                End If
                            Next
                            Transaction.Commit()
                        End If
                    End Using
                End Using
                Return True
            Catch
                Return False
            End Try
        End Function

    End Class

End Namespace

