Imports System.Runtime.Serialization.Formatters.Binary
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Text
Imports System.Text.RegularExpressions

'=========================================================================================================='
'DESENVOLVIDO POR: DANIEL TRAVAGLIA DOS SANTOS (ANALISTA DE SISTEMAS)
'EM:2014
'PARA:SPHE ENGENHARIA
'=========================================================================================================='
'CONTATO: dtstravaglia@yahoo.com.br \ (011) 98135-6040
'=========================================================================================================='

Namespace Engine2

    ''' <summary>
    ''' Permite serializar de objetos.
    ''' Armazena as informações em arquivos XML ou Binário.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Serialize

        ''' <summary>
        ''' Enumera formatos para serialização.
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum SerializeFormats

            ''' <summary>
            ''' Formato XML
            ''' </summary>
            ''' <remarks></remarks>
            Xml

            ''' <summary>
            ''' Formato binário
            ''' </summary>
            ''' <remarks></remarks>
            Bin

        End Enum

        ''' <summary>
        ''' Serializa um objeto.
        ''' </summary>
        ''' <param name="SerializeFormat">Formato de armazenagem.</param>
        ''' <param name="FileName">Nome do arquivo de armazenagem.</param>
        ''' <param name="SerializedObject">Objeto a ser serializado.</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function Serialize(ByVal SerializeFormat As SerializeFormats, ByVal FileName As String, ByVal SerializedObject As Object) As Boolean
            Dim DF As FileStream = Nothing
            Try
                Select Case SerializeFormat
                    Case SerializeFormats.Bin
                        Dim SR As New BinaryFormatter()
                        DF = New FileStream(FileName, FileMode.Create, FileAccess.Write, FileShare.None)
                        SR.Serialize(DF, SerializedObject)
                        DF.Close()
                        Return True
                    Case SerializeFormats.Xml
                        Dim SX As New Serialization.XmlSerializer(SerializedObject.GetType)
                        DF = New FileStream(FileName, FileMode.Create, FileAccess.Write, FileShare.None)
                        SX.Serialize(DF, SerializedObject)
                        DF.Close()
                        Return True
                    Case Else
                        Return Nothing
                End Select
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Serialize.Serialize, motivo: " & ex.Message)
            Finally
                DisposeObject(DF)
            End Try
        End Function

        ''' <summary>
        ''' Reverte a serialização de um objeto.
        ''' </summary>
        ''' <param name="SerializeFormat">Formato de armazenagem.</param>
        ''' <param name="FileName">Nome do arquivo de armazenagem.</param>
        ''' <param name="SerializedObject">Objeto que recebe o item serializado.</param>
        ''' <returns>Object</returns>
        ''' <remarks></remarks>
        Public Shared Function DeSerialize(ByVal SerializeFormat As SerializeFormats, ByVal FileName As String, ByVal SerializedObject As Object) As Object
            Dim DF As FileStream = Nothing
            Try
                Select Case SerializeFormat
                    Case SerializeFormats.Bin
                        Dim DS As New BinaryFormatter()
                        DF = New FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.None)
                        SerializedObject = DS.Deserialize(DF)
                        DF.Close()
                        Return SerializedObject
                    Case SerializeFormats.Xml
                        Dim DX As New Serialization.XmlSerializer(SerializedObject.GetType)
                        DF = New FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.None)
                        SerializedObject = DX.Deserialize(DF)
                        DF.Close()
                        Return SerializedObject
                    Case Else
                        Return Nothing
                End Select
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Serialize.DeSerialize, motivo: " & ex.Message)
            Finally
                DisposeObject(DF)
            End Try
        End Function

        ''' <summary>
        ''' Salva o XML em arquivo
        ''' </summary>
        ''' <param name="XML">String XML</param>
        ''' <param name="FileName">Caminho para salvamento</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Public Shared Function XmlToFile(XML As String, ByVal FileName As String) As Boolean
            Try
                If System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(FileName)) = False Then
                    Dim DirectoryInfo As System.IO.DirectoryInfo = System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(FileName))
                    If DirectoryInfo.Exists = False Then
                        Throw New System.Exception("Não é possível criar o diretório '" & System.IO.Path.GetDirectoryName(FileName) & "'")
                    End If
                End If
                If System.IO.File.Exists(FileName) = True Then
                    System.IO.File.Copy(FileName, FileName.Replace(".xml", ".$xml"), True)
                End If
                Dim FileStream As FileStream = System.IO.File.Create(FileName, FileMode.OpenOrCreate)
                Using fluxoTexto As New IO.StreamWriter(FileStream)
                    With fluxoTexto
                        .Write(XML)
                        .Close()
                    End With
                End Using
                Return True
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Serialize.XmlToFile, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Converte objeto em xml
        ''' </summary>
        ''' <param name="ObjectItem">Objeto a ser serializado</param>
        ''' <returns>String</returns>
        ''' <remarks></remarks>
        Public Shared Function ObjectToXml(ByVal ObjectItem As Object) As String
            Try
                Dim XmlDocument As New XmlDocument
                Dim XmlSerializer As New Serialization.XmlSerializer(ObjectItem.GetType)
                Using stream As Stream = New MemoryStream
                    XmlSerializer.Serialize(stream, ObjectItem)
                    stream.Flush()
                    stream.Seek(0, SeekOrigin.Begin)
                    XmlDocument.Load(stream)
                End Using
                Return XmlDocument.InnerXml
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Serialize.ObjectToXml, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Converte Xml em objeto
        ''' </summary>
        ''' <param name="Xml">String XML</param>
        ''' <param name="ObjectItem">Objeto de origem</param>
        ''' <returns>Object</returns>
        ''' <remarks></remarks>
        Public Shared Function XmlToObject(ByVal Xml As String, ByVal ObjectItem As Object) As Object
            Try
                Dim XmlSerializer As New XmlSerializer(ObjectItem.GetType)
                Dim XmlReader As System.Xml.XmlReader = System.Xml.XmlReader.Create(New System.IO.StringReader(Xml))
                ObjectItem = XmlSerializer.Deserialize(XmlReader)
                Return ObjectItem
            Catch ex As System.Exception
                Throw New System.Exception("Erro em Serialize.XmlToObject, motivo: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Destroi objetos
        ''' </summary>
        ''' <param name="DisposeableObject">Objeto a ser destruído</param>
        ''' <returns>Boolean</returns>
        ''' <remarks></remarks>
        Private Shared Function DisposeObject(ByVal DisposeableObject As Object) As Boolean
            Try
                If IsNothing(DisposeableObject) = False Then
                    DisposeableObject.Dispose()
                End If
                Return True
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Determina se um tag com ou sem valor associado existe no XML
        ''' </summary>
        ''' <param name="Xml">Xml</param>
        ''' <param name="TagName">Nome da propriedade</param>
        ''' <param name="Value">Caso seja informado o valor associado a tag será avaliado</param>
        ''' <returns>Xml\Nothing</returns>
        ''' <remarks></remarks>
        Public Shared Function XmlContains(Xml As String, TagName As String, Optional Value As Object = Nothing) As Boolean
            Try
                Dim GroupCollection As GroupCollection
                GroupCollection = Regex.Match(Xml, "<" & TagName & ">([^>]*)<\/" & TagName & ">").Groups
                If GroupCollection.Count > 0 Then
                    If IsNothing(Value) = False Then
                        If Engine2.Serialize.XmlGetValue(Xml, TagName) = Value Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        Return True
                    End If
                End If
                Return False
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Obtem o valor de uma tag no XML
        ''' </summary>
        ''' <param name="Xml">Xml</param>
        ''' <param name="TagName">Nome da tag</param>
        ''' <returns>Xml\Nothing</returns>
        ''' <remarks></remarks>
        Public Shared Function XmlGetValue(Xml As String, TagName As String) As Object
            Try
                Dim GroupCollection As GroupCollection
                GroupCollection = Regex.Match(Xml, "<" & TagName & ">([^>]*)<\/" & TagName & ">").Groups
                If GroupCollection.Count > 0 Then
                    Return GroupCollection(1)
                End If
                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Atualiza o valor associado a uma tag no XML
        ''' </summary>
        ''' <param name="Xml">Xml</param>
        ''' <param name="TagName">Nome da tag</param>
        ''' <param name="Value">Valor da tag</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function XmlUpdate(Xml As String, TagName As String, Value As Object) As Object
            Try
                Return Regex.Replace(Xml, Regex.Match(Xml, "<" & TagName & ">([^>]*)<\/" & TagName & ">").Groups(1).ToString, Value)
            Catch
                Return Nothing
            End Try
        End Function

    End Class

End Namespace

