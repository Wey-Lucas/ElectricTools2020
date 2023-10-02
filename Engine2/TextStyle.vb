'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.GraphicsInterface
Imports System.IO
Imports System.Drawing

Namespace Engine2

    ''' <summary>
    ''' Classe para gerenciamento de estilos de texto
    ''' </summary>
    ''' <remarks></remarks>
    Public Class TextStyle

        ''' <summary>
        '''  Avalia se o estilo de texto existe
        ''' </summary>
        ''' <param name="TextStyleName">Nome do estilo de texto</param>
        ''' <param name="Transaction">Transação</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Contains(TextStyleName As String, Optional Transaction As Transaction = Nothing) As Boolean
            Contains = False
            Dim Document As Document
            Dim TextStyleTable As TextStyleTable
            Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
            If IsNothing(Transaction) = False Then
                TextStyleTable = Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForWrite, False)
                Contains = TextStyleTable.Has(TextStyleName)
            Else
                Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                    Transaction = Document.Database.TransactionManager.StartTransaction()
                    Using Transaction
                        Contains = Contains(TextStyleName, Transaction)
                        Transaction.Abort()
                    End Using
                End Using
            End If
            Return Contains
        End Function

        ''' <summary>
        ''' Obtem o estilo de texto corrente
        ''' </summary>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public Shared Function GetCurrent() As String
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim TextStyleTableRecord As TextStyleTableRecord
                Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                    TextStyleTableRecord = Transaction.GetObject(Document.Database.Textstyle, OpenMode.ForRead)
                    Transaction.Dispose()
                End Using
                Return TextStyleTableRecord.Name
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Seta o estilo de texto corrente
        ''' </summary>
        ''' <param name="TextStyleName">Nome do estilo de texto</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function SetCurrent(ByVal TextStyleName As String) As Boolean
            Try
                If TextStyleName.Trim <> "" Then
                    Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                    Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                        Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                            Dim TextStyleTable As TextStyleTable = Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForWrite)
                            Document.Database.Textstyle = TextStyleTable(TextStyleName)
                            Transaction.Commit()
                        End Using
                    End Using
                End If
                Return True
            Catch Ex As System.Exception
                Throw New System.Exception("Erro em TextStyle.SetCurrent, motivo: " & Ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Adiciona o estilo de texto
        ''' </summary>
        ''' <param name="TextStyleName">Nome do estilo de texto</param>
        ''' <param name="FontName">Nome da fonte</param>
        ''' <param name="FontSize">Tamanho da fonte</param>
        ''' <param name="Bold">Define se é negrito</param>
        ''' <param name="Italic">Define se a fonte é itálica</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Create(TextStyleName As String, FontName As String, FontSize As Double, Optional Bold As Boolean = False, Optional Italic As Boolean = False) As Boolean
            Dim CurrentStyle As String = Engine2.TextStyle.GetCurrent
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
            Dim FontFile As Object = Engine2.TextStyle.GetFontFileName(FontName)
            If IsNothing(FontFile) = True Then
                FontFile = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("FONTALT") 'Fonte alternativa
            End If
            Try
                Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                    Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                        Dim TextStyleTable As TextStyleTable = CType(Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)
                        Dim TextStyleTableRecord As TextStyleTableRecord = New TextStyleTableRecord()
                        If TextStyleTable.Has(TextStyleName) = False Then
                            With TextStyleTableRecord
                                .FileName = System.IO.Path.GetFileName(FontFile).ToUpper.Trim
                                .Name = TextStyleName
                                .XScale = 1
                                .TextSize = FontSize
                                .Font = New FontDescriptor(FontName, Bold, Italic, Nothing, Nothing)
                                TextStyleTable.Add(TextStyleTableRecord)
                            End With
                            Transaction.AddNewlyCreatedDBObject(TextStyleTableRecord, True)
                            Document.Database.Textstyle = TextStyleTableRecord.ObjectId
                            Transaction.Commit()
                            Return True
                        Else
                            Transaction.Dispose()
                            Return True
                        End If
                    End Using
                End Using
            Catch Ex As System.Exception
                Throw New System.Exception("Erro em TextStyle.Create, motivo: " & Ex.Message)
            Finally
                Engine2.TextStyle.SetCurrent(CurrentStyle)
            End Try
        End Function

        ''' <summary>
        ''' Adiciona o estilo de texto
        ''' </summary>
        ''' <param name="Font">Fonte</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Create(Font As System.Drawing.Font) As Boolean
            Dim CurrentStyle As String = Engine2.TextStyle.GetCurrent
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
            Dim FontFile As Object = Engine2.TextStyle.GetFontFileName(Font.FontFamily.Name)
            If IsNothing(FontFile) = True Then
                FontFile = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("FONTALT") 'Fonte alternativa
            End If
            Try
                Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                    Using Transaction As Transaction = Document.Database.TransactionManager.StartTransaction()
                        Dim TextStyleTable As TextStyleTable = CType(Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)
                        Dim TextStyleTableRecord As TextStyleTableRecord = New TextStyleTableRecord()
                        If TextStyleTable.Has(Font.FontFamily.Name) = False Then
                            With TextStyleTableRecord
                                .FileName = System.IO.Path.GetFileName(FontFile).ToUpper.Trim
                                .Name = .Name
                                .XScale = 1
                                .TextSize = Font.Size
                                .Font = New FontDescriptor(Font.FontFamily.Name, Font.Bold, Font.Italic, Nothing, Nothing)
                                TextStyleTable.Add(TextStyleTableRecord)
                            End With
                            Transaction.AddNewlyCreatedDBObject(TextStyleTableRecord, True)
                            Document.Database.Textstyle = TextStyleTableRecord.ObjectId
                            Transaction.Commit()
                            Return True
                        Else
                            Transaction.Dispose()
                            Return True
                        End If
                    End Using
                End Using
            Catch Ex As System.Exception
                Throw New System.Exception("Erro em TextStyle.Create, motivo: " & Ex.Message)
            Finally
                Engine2.TextStyle.SetCurrent(CurrentStyle)
            End Try
        End Function

        ''' <summary>
        ''' Obtem o ID estilo de texto
        ''' </summary>
        ''' <param name="TextStyleName">Nome do estilo de texto</param>
        ''' <returns>TextStyleTableRecord</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextStyle(TextStyleName As String) As TextStyleTableRecord
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
                Dim TextStyleTableRecord As TextStyleTableRecord
                Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        Dim TextStyleTable As TextStyleTable = CType(Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForRead, False), TextStyleTable)
                        If TextStyleTable.Has(TextStyleName) = False Then
                            Return Nothing
                        End If
                        TextStyleTableRecord = TryCast(Transaction.GetObject(TextStyleTable(TextStyleName), OpenMode.ForRead), TextStyleTableRecord)
                        Return TextStyleTableRecord
                    End Using
                End Using
                Return Nothing
            Catch ex As System.Exception
                Throw New System.Exception("Erro em TextStyle.GetTextStyle, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Obtem o ID estilo de texto
        ''' </summary>
        ''' <param name="ObjectID">ID do estilo de texto</param>
        ''' <returns>TextStyleTableRecord</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextStyle(ObjectID As ObjectId) As TextStyleTableRecord
            Try
                Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
                Dim TextStyleTableRecord As TextStyleTableRecord
                Using DocumentLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                    Using Transaction As Transaction = Document.TransactionManager.StartTransaction()
                        Dim TextStyleTable As TextStyleTable = CType(Transaction.GetObject(Document.Database.TextStyleTableId, OpenMode.ForRead, False), TextStyleTable)
                        If TextStyleTable.Has(ObjectID) = False Then
                            Return Nothing
                        End If
                        TextStyleTableRecord = TryCast(Transaction.GetObject(ObjectID, OpenMode.ForRead), TextStyleTableRecord)
                        Return TextStyleTableRecord
                    End Using
                End Using
                Return Nothing
            Catch ex As System.Exception
                Throw New System.Exception("Erro em TextStyle.GetTextStyle, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Verifica se o arquivo de fonte existe
        ''' </summary>
        ''' <param name="FontName"></param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function ExistsFontFile(FontName As String) As Boolean
            Dim fontsFolderPath As Object = Environment.GetFolderPath(Environment.SpecialFolder.Fonts)
            If Directory.GetFiles(fontsFolderPath, FontName & ".*").Length > 0 Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Retorna o nome do arquivo de fonte dado o nome da fonte do texto
        ''' </summary>
        ''' <param name="FontName"></param>
        ''' <returns>Object</returns>
        ''' <remarks></remarks>
        Public Shared Function GetFontFileName(FontName As String) As Object
            Dim fontsFolderPath As Object = Environment.GetFolderPath(Environment.SpecialFolder.Fonts)
            Dim Fonts() As String = Directory.GetFiles(fontsFolderPath, FontName & ".*")
            If Fonts.Length > 0 Then
                Return Fonts(0)
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Obtem a coleção de fontes
        ''' </summary>
        ''' <returns>ArrayList</returns>
        ''' <remarks></remarks>
        Private Shared Function GetFonts() As ArrayList
            GetFonts = New ArrayList
            Dim InstalledFontCollection As New Text.InstalledFontCollection
            Dim FontFamily() As FontFamily = InstalledFontCollection.Families()
            For Each FontName As FontFamily In FontFamily
                If GetFonts.Contains(FontName.Name) = False Then
                    GetFonts.Add(FontName.Name)
                End If
                If GetFonts.Contains(FontName.Name.ToUpper) = False Then
                    GetFonts.Add(FontName.Name.ToUpper)
                End If
                If GetFonts.Contains(FontName.Name.ToLower) = False Then
                    GetFonts.Add(FontName.Name.ToLower)
                End If
            Next
            Return GetFonts
        End Function

        ''' <summary>
        ''' Retorna a coleção de estilos de texto
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function TextStyles() As ArrayList
            TextStyles = New ArrayList
            Dim Document As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument()
            Dim Database As Database = Document.Database
            Dim SymbolTable As SymbolTable
            Dim TextStyleTableRecord As TextStyleTableRecord
            Using Transaction As Transaction = Database.TransactionManager.StartTransaction()
                SymbolTable = Transaction.GetObject(Database.TextStyleTableId, OpenMode.ForRead)
                For Each ObjectId As ObjectId In SymbolTable
                    TextStyleTableRecord = Transaction.GetObject(ObjectId, OpenMode.ForRead)
                    If TextStyleTableRecord.Name.Contains("|") = False And TextStyleTableRecord.Name.Trim <> "" Then
                        TextStyles.Add(TextStyleTableRecord.Name)
                    End If
                Next
                Transaction.Abort()
            End Using
            TextStyles.Sort()
            Return TextStyles
        End Function

    End Class

End Namespace
