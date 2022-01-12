Imports CClass
Imports System.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Web.Services.Protocols
Imports CertificadoDigital
Imports System.text.RegularExpressions
Imports NFSeWebService
Imports System.IO
Imports NFSeCore

Namespace cCenti2
    Public Class ConsultarLote
        Inherits AConsultarLote


        Private isExternal As Boolean = False
        Private codgCidade As Integer
        Dim result As Object = ""
        Private numrCnpj As Long
        Private protocolo As String
        Private idEmpresa As Integer
        Private tipoConsulta As Boolean
        Dim ambiente As TipoAmbiente = Nothing

        Public Sub New(ByRef conn As IDbConnection, _
                       ByRef trans As IDbTransaction, _
                       ByRef factory As DBFactory, _
                       ByRef publicVar As PublicVar)
            MyBase.factory = factory
            MyBase.publicVar = publicVar
            MyBase.conn = conn
            MyBase.trans = trans
        End Sub

        Public Overrides Function processar(ByVal codgCidade As Integer, ByVal numrCnpj As Long, _
                                             ByVal protocolo As String, ByVal inscMunicipal As String, ByVal idEmpresa As Integer) As String
            Me.codgCidade = codgCidade
            Me.numrCnpj = numrCnpj
            Me.protocolo = protocolo
            Dim certificado As X509Certificate = Nothing
            Dim propCertificado As IPropriedadeCertificado = Nothing
            Dim urlSituacao As String = Nothing
            Dim urlConsulta As String = Nothing
            Dim cidadeData As NfseCidadesData = GetCidade(codgCidade)
            Try
                Me.idEmpresa = idEmpresa

                ' obter o ambiente producao ou homologacao
                ambiente = ParametroUtil.GetParametro(idEmpresa, "AMBIENTE")

                ' gerar o xml de consulta de protocolo
                Dim xml As String = GetXmlConsulta(protocolo, numrCnpj, inscMunicipal)
                Dim PropriedadesXML As IPropriedadesXML = FactoryCore.GetPropriedadesXML(codgCidade)
                ' gerar o xml de consultar situacao protocolo

                xml = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:nfse=""http://www.abrasf.org.br/ABRASF/arquivos/nfse.xsd"">" + _
                         "<soapenv:Header/>" + _
                         "<soapenv:Body>" + _
                            "<nfse:ConsultarLoteRps>" + _
                                "<header><![CDATA[" + PropriedadesXML.GetCabecalho + "]]></header>" + _
                                "<parameters><![CDATA[" + xml + "]]></parameters>" + _
                            "</nfse:ConsultarLoteRps>" + _
                        "</soapenv:Body>" + _
                      "</soapenv:Envelope>"

                propCertificado = FactoryCore.GetPropriedadeCertificado(conn, trans, factory, publicVar, codgCidade)
                'If ambiente = TipoAmbiente.HOMOLOGACAO Then
                '    numrCnpj = 2348447000181
                '    inscMunicipal = "30"
                'End If
                certificado = propCertificado.GetCertificadoTransmissao(idEmpresa, numrCnpj)

                If certificado Is Nothing Then
                    Throw New FalhaException("Erro ao obter certificado digital, para CONSULTAR LOTE, através do nº do documento informado : " + numrCnpj.ToString)
                End If

                If ambiente = TipoAmbiente.HOMOLOGACAO Then
                    urlConsulta = ParametroUtil.GetParametroCidade(codgCidade, "URL_AMBIENTE_HOMOLOGACAO")
                Else
                    urlConsulta = ParametroUtil.GetParametroCidade(codgCidade, "URL_AMBIENTE_PRODUCAO")
                End If

                Dim doc As XmlDocument = New XmlDocument()
                doc.PreserveWhitespace = True
                doc.LoadXml(xml)
                xmlRetorno = HTTPSoapTextRequest(certificado, doc, urlConsulta, idEmpresa, "http://nfse.abrasf.org.br/ConsultarLoteRps")
                doc = New XmlDocument
                ' processar retorno, tirar as tags soap
                doc.LoadXml(xmlRetorno)
                Dim nodeConsultaLote As XmlNode = doc.SelectSingleNode(GetXpathFromString("Envelope, Body, ConsultarLoteRpsResponse, return"))
                nodeConsultaLote.InnerXml = CClass.StringUtil.tirarAcentos(nodeConsultaLote.InnerXml.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&#xD;", ""))
                doc.LoadXml(nodeConsultaLote.InnerXml)

                If nodeConsultaLote Is Nothing Then
                    Throw New FalhaException("Retorno da prefeitura não pode ser processado: objeto resposta veio nulo ou vazio. Retorno: " + xmlRetorno)
                End If

                result = nodeConsultaLote.InnerXml

            Catch ex As Exception
                If TypeOf ex Is FalhaException Then
                    Throw ex
                Else
                    LogInfo(ex)
                    Throw New FalhaException("Erro ao tentar transmitir consulta de lote de NFSE: " + ex.Message, ex)
                End If
            End Try

            Return CClass.StringUtil.tirarAcentos(result.ToString)
        End Function

        Protected Overrides Function GetCodgVerificacao(ByVal xml As String) As String
            Dim xmldoc As XmlDocument = New XmlDocument
            Dim result As String = Nothing

            Try
                xmldoc.LoadXml(xml)
            Catch ex As Exception
                Throw New FalhaException("Erro ao ler XML de retorno: " + ex.Message + _
                                         " - conteudo do retorno: " + xml)
            End Try

            Return XmlUtil.getValorTag(xmldoc, "CodigoVerificacao")

            Throw New FalhaException("Erro ao ler XML de retorno: o xml retornado " + _
                                         "pela prefeitura não contem a tag ""CodigoVerificacao"". " + _
                                         "Esta fora do padrão esperado.")
        End Function

        Protected Overrides Function isRetornoProcessamentoOk(ByVal xml As String) As Boolean
            Dim xmldoc As XmlDocument = New XmlDocument

            LogInfo("AConsultarLote isRetornoProcessamentoOk xml: " + xml)

            Try
                xmldoc.LoadXml(xml)
            Catch ex As Exception
                Throw New FalhaException("Erro ao ler XML de retorno: " + ex.Message + _
                                         " - conteudo do retorno: " + xml)
            End Try

            Return CClass.XmlUtil.getDocByTag(xmldoc, "Nfse") IsNot Nothing
        End Function

        Protected Overrides Function GetXmlConsulta(ByVal protocolo As String, ByVal numrDocumento As Long, ByVal inscMunicipal As String) As String
            If ambiente = TipoAmbiente.HOMOLOGACAO Then
                numrDocumento = 2348447000181
                inscMunicipal = "30"
            End If

            Dim xmlConsultaLote As String = "<?xml version=""1.0"" encoding=""utf-8""?>" + _
                                            "<ConsultarLoteRpsEnvio Id="""" xmlns=""http://ws.speedgov.com.br/consultar_lote_rps_envio_v1.xsd"">" + _
                                              "<Prestador>" + _
                                                  "<Cnpj xmlns=""http://ws.speedgov.com.br/tipos_v1.xsd"">" + numrDocumento.ToString("00000000000000") + "</Cnpj>" + _
                                                  "<InscricaoMunicipal xmlns=""http://ws.speedgov.com.br/tipos_v1.xsd"">" + inscMunicipal + "</InscricaoMunicipal>" + _
                                              "</Prestador>" + _
                                              "<Protocolo>" + protocolo + "</Protocolo>" + _
                                            "</ConsultarLoteRpsEnvio>"

            Return xmlConsultaLote
        End Function

        Protected Overrides Function GetMotivoRejeicao() As String
            Dim xpath As String = APropriedadesXML.GetXpathFromCollection(New String() {"ConsultarLoteRpsResposta"})

            Dim resultrejeicao As String
            resultrejeicao = GetMotivoRejeicao1(result, xpath)

            Return resultrejeicao
        End Function

        Private Function GetMotivoRejeicao1(ByVal xml As String, ByVal xpath As String) As String
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
                                    result += Chr(13) + Chr(10) + cNode.Name + " - " + cNode.FirstChild.InnerText + " " + cNode.InnerText.Replace(cNode.FirstChild.InnerText, "")
                                Else
                                    result += cNode.Name + " - " + +cNode.FirstChild.InnerText + " " + cNode.InnerText.Replace(cNode.FirstChild.InnerText, "")
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
    End Class
End Namespace
