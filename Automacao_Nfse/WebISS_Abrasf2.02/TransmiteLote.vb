Imports CClass
Imports System.Security.Cryptography.X509Certificates
Imports System.Web.Services
Imports System.Xml
Imports System.Text.RegularExpressions
Imports CertificadoDigital
Imports NFSeWebService
Imports System.IO
Imports NFSeCore
Imports System.Text
Imports System.Globalization

Namespace cWebISS2
    Public Class TransmiteLote
        Inherits ATransmiteLote

        Public Sub New(ByRef conn As IDbConnection, _
                       ByRef trans As IDbTransaction, _
                       ByRef factory As DBFactory, _
                       ByRef publicVar As PublicVar)
            MyBase.factory = factory
            MyBase.publicVar = publicVar
            MyBase.conn = conn
            MyBase.trans = trans
        End Sub

        'Public Overrides Sub SetParams(ByRef params() As Object, ByVal xml As String, ByVal codgCidade As Integer, ByVal idEmpresa As Integer)
        '    Dim PropriedadesXML As IPropriedadesXML = FactoryCore.GetPropriedadesXML(codgCidade)
        '    Dim cabecalho As String = PropriedadesXML.GetCabecalho()
        '    params = New Object() {cabecalho, xml}
        'End Sub

        Public Overrides Function processar(ByRef xml As String, ByVal codgCidade As Integer, ByVal numrCnpj As Long, ByVal idEmpresa As Long) As String
            Dim ambiente As TipoAmbiente = Nothing
            Dim certificado As X509Certificate = Nothing
            Dim usarProxy As Boolean = False
            Dim result As Object = Nothing
            Dim responseSOAP As New XmlDocument()
            Dim propCertificado As IPropriedadeCertificado = Nothing
            Dim url As String = Nothing
            Dim soapAction As String = ""
            Dim cidadeData As NfseCidadesData = GetCidade(codgCidade)
            Dim xmlRetorno As String
            Dim xmlEnvio As String


            Try
                ' obter o ambiente producao ou homologacao
                ambiente = ParametroUtil.GetParametro(idEmpresa, "AMBIENTE")
                LogInfo(idEmpresa, "TransmiteLote Processar 1")
                ' obter a classe especifica para gerenciar o certificado para esta cidade
                propCertificado = FactoryCore.GetPropriedadeCertificado(conn, trans, factory, publicVar, codgCidade)
                LogInfo(idEmpresa, "TransmiteLote Processar 2")
                certificado = propCertificado.GetCertificadoTransmissao(idEmpresa, numrCnpj)

                If certificado Is Nothing Then
                    Throw New FalhaException("Erro ao obter certificado digital, para TRANSMISSÃO, através do nº do documento informado : " + numrCnpj.ToString)
                End If
                LogInfo(idEmpresa, "TransmiteLote Processar 3")
                If ambiente = TipoAmbiente.HOMOLOGACAO Then
                    url = ParametroUtil.GetParametroCidade(codgCidade, "URL_AMBIENTE_HOMOLOGACAO")
                Else
                    url = ParametroUtil.GetParametroCidade(codgCidade, "URL_AMBIENTE_PRODUCAO")
                End If
                LogInfo(idEmpresa, "TransmiteLote Processar 4")
                Dim PropriedadesXML As IPropriedadesXML = FactoryCore.GetPropriedadesXML(codgCidade)
                xmlEnvio = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:nfse=""http://nfse.abrasf.org.br"">" + _
                          "<soapenv:Header/>" + _
                          "<soapenv:Body>" + _
                              "<nfse:GerarNfseRequest>" + _
                                  "<nfseCabecMsg><![CDATA[" + PropriedadesXML.GetCabecalho() + "]]></nfseCabecMsg>" + _
                                  "<nfseDadosMsg><![CDATA[<?xml version=""1.0"" encoding=""UTF-8""?>" + xml + "]]></nfseDadosMsg>" + _
                              "</nfse:GerarNfseRequest>" + _
                          "</soapenv:Body>" + _
                      "</soapenv:Envelope>"

                LogInfo(idEmpresa, "TransmiteLote Processar 5: " + xmlEnvio)
                Dim doc As XmlDocument = New XmlDocument()
                doc.PreserveWhitespace = True
                doc.LoadXml(xmlEnvio)
                LogInfo(idEmpresa, "TransmiteLote Processar 6")

                xmlRetorno = SendRequestChilkat(xmlEnvio, url, "http://nfse.abrasf.org.br/GerarNfse", numrCnpj, idEmpresa)

                LogInfo(idEmpresa, "TransmiteLote Processar 7 :" + xmlRetorno)
                xmlRetorno = RemoveAccents(xmlRetorno)
                xmlRetorno = xmlRetorno.Replace("<?xml version=""1.0"" encoding=""utf-8""?>", "").Replace("&lt;?xml version=""1.0"" encoding=""utf-8""?&gt;", "")
                responseSOAP.LoadXml(xmlRetorno)
                Dim node As XmlNode = responseSOAP.SelectSingleNode(GetXpathFromString("Envelope, Body, GerarNfseResponse, outputXML"))
                If node IsNot Nothing Then
                    node.InnerXml = node.InnerXml.Replace("&lt;", "<").Replace("&gt;", ">")

                End If
                If node Is Nothing Then
                    Throw New FalhaException("Retorno da transmissão do lote a prefeitura não pode ser processado: objeto veio nulo. Retorno: " + xmlRetorno)
                End If
                node.InnerXml = node.InnerXml.Replace("<![CDATA[", "").Replace("]]>", "")
                xmlRetorno = node.InnerXml.ToString
                xmlRetorno = CClass.StringUtil.tirarAcentos(xmlRetorno)
            Catch ex As Exception
                If TypeOf ex Is FalhaException Then
                    Throw ex
                Else
                    Throw New FalhaException("Erro ao tentar transmitir lote de RPS: " + ex.Message, ex)
                End If
            End Try

            Return xmlRetorno.Replace("> <", ">" & vbNewLine & "<")
        End Function

        Private Shared Function RemoveAccents(ByVal text As String) As String
            Dim sbReturn As StringBuilder = New StringBuilder()
            Dim arrayText = text.Normalize(NormalizationForm.FormD).ToCharArray()
            For Each letter As Char In arrayText
                If CharUnicodeInfo.GetUnicodeCategory(letter) <> UnicodeCategory.NonSpacingMark Then sbReturn.Append(letter)
            Next
            Return sbReturn.ToString()
        End Function

        Protected Overrides Function GetMotivoRejeicao(ByVal xmlRetorno As String) As String
            Dim xpath As String = APropriedadesXML.GetXpathFromCollection(New String() {"GerarNfseResposta", "ListaMensagemRetorno", "MensagemRetorno"})

            Return GetMotivoRejeicaoFunction(xmlRetorno, xpath)
        End Function

        Protected Function GetMotivoRejeicaoFunction(ByVal xml As String, ByVal xpath As String) As String
            Dim xmldoc As New XmlDocument()
            Dim result As String = ""
            Dim nodes As XmlNodeList = Nothing
            Dim primeiraLinha As Boolean = True

            Try
                xmldoc.LoadXml(xml)

                nodes = xmldoc.SelectNodes(xpath)

                If nodes IsNot Nothing AndAlso nodes.Count > 0 Then
                    For Each node As XmlNode In nodes
                        If node.ChildNodes.Count > 0 Then
                            For Each cNode As XmlNode In node.ChildNodes
                                If Not primeiraLinha Then
                                    result += Chr(13) + Chr(10) + cNode.Name + " - " + cNode.InnerText
                                Else
                                    result += cNode.Name + " - " + cNode.InnerText

                                    primeiraLinha = False
                                End If
                            Next
                        Else
                            result += node.Name + " - " + node.InnerText
                        End If
                    Next
                Else
                    result = xml
                End If

                If result Is Nothing OrElse result.Trim.Length = 0 Then
                    result = xml
                End If

                Return result
            Catch ex As Exception
                ' retornar o xml que veio da prefeitura, pois o erro que aconteceu nao era esperado
                Return xml
            Finally
                xmldoc = Nothing
                nodes = Nothing
                xpath = Nothing
            End Try
        End Function

        Public Overrides Function isRetornoEnvioOk(ByRef xml As String) As Boolean
            ' verificar se o retorno é XML

            Dim xmldoc As XmlDocument = New XmlDocument
            Try

                xmldoc.LoadXml(xml)
            Catch ex As Exception
                Throw New FalhaException("Erro ao processar retorno: O conteúdo retornado não é XML - " + ex.Message + "Conteúdo retornado: " + xml, ex)
            End Try
            Dim codgVerificacao As String = XmlUtil.getValorTag(xmldoc, "CodigoVerificacao")

            If codgVerificacao.Trim.Length = 0 Then
                Return False
            End If

            Return True
        End Function
    End Class
End Namespace
