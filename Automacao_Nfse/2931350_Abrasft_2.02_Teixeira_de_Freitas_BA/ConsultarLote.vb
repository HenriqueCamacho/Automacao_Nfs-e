Imports CClass
Imports System.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Web.Services.Protocols
Imports CertificadoDigital
Imports System.text.RegularExpressions
Imports NFSeWebService
Imports System.IO
Imports NFSeCore

Namespace c2931350
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

        Public Overrides Sub SetParams(ByRef params() As Object, ByVal xml As String, ByVal codgCidade As Integer, ByVal idEmpresa As Integer)
            Dim PropriedadesXML As IPropriedadesXML = FactoryCore.GetPropriedadesXML(codgCidade)
            Dim cabecalho As String = PropriedadesXML.GetCabecalho()
            params = New Object() {cabecalho, xml}
        End Sub

        ' Nao usa este serviço por estar utilizando nfse envio.

        Public Overrides Function processar(ByVal codgCidade As Integer, ByVal numrCnpj As Long, _
                                             ByVal protocolo As String, ByVal inscMunicipal As String, ByVal idEmpresa As Integer) As String
            Me.codgCidade = codgCidade
            Me.numrCnpj = numrCnpj
            Me.protocolo = protocolo
            Dim certificado As X509Certificate = Nothing
            Dim propCertificado As IPropriedadeCertificado = Nothing
            Dim url As String = Nothing
            Dim cidadeData As NfseCidadesData = GetCidade(codgCidade)
            Try
                Me.idEmpresa = idEmpresa

                ' obter o ambiente producao ou homologacao
                ambiente = ParametroUtil.GetParametro(idEmpresa, "AMBIENTE")

                ' gerar o xml de consulta de protocolo
                Dim xml As String = GetXmlConsulta(protocolo, numrCnpj, inscMunicipal)
                Dim PropriedadesXML As IPropriedadesXML = FactoryCore.GetPropriedadesXML(codgCidade)
                ' gerar o xml de consultar situacao protocolo
                Dim xmlSitucao As String = GetXmlConsultaSituacaoLote(protocolo, numrCnpj, inscMunicipal)
                xmlSitucao = "<?xml version=""1.0"" encoding=""utf-8""?>" + _
                             "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" + _
                                      "<soap:Body>" + _
                                                  "<ConsultarSituacaoLoteRps xmlns=""http://tempuri.org/"">" + _
                                                     "<cabec>" + PropriedadesXML.GetCabecalho.Replace("<", "&lt;").Replace(">", "&gt;") + "</cabec>" + _
                                                      "<msg>" + xmlSitucao.Replace("<", "&lt;").Replace(">", "&gt;") + "</msg>" + _
                                              "</ConsultarSituacaoLoteRps>" + _
                                      "</soap:Body>" + _
                              "</soap:Envelope>"

                xml = "<?xml version=""1.0"" encoding=""utf-8""?>" + _
                             "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" + _
                                      "<soap:Body>" + _
                                                  "<ConsultarLoteRps xmlns=""http://tempuri.org/"">" + _
                                                     "<cabec>" + PropriedadesXML.GetCabecalho.Replace("<", "&lt;").Replace(">", "&gt;") + "</cabec>" + _
                                                      "<msg>" + xml.Replace("<", "&lt;").Replace(">", "&gt;") + "</msg>" + _
                                                  "</ConsultarLoteRps>" + _
                                      "</soap:Body>" + _
                              "</soap:Envelope>"
                propCertificado = FactoryCore.GetPropriedadeCertificado(conn, trans, factory, publicVar, codgCidade)

                certificado = propCertificado.GetCertificadoTransmissao(idEmpresa, numrCnpj)

                If certificado Is Nothing Then
                    Throw New FalhaException("Erro ao obter certificado digital, para CONSULTAR LOTE, através do nº do documento informado : " + numrCnpj.ToString)
                End If

                If ambiente = TipoAmbiente.HOMOLOGACAO Then
                    url = ParametroUtil.GetParametroCidade(codgCidade, "URL_AMBIENTE_HOMOLOGACAO")
                Else
                    url = ParametroUtil.GetParametroCidade(codgCidade, "URL_AMBIENTE_PRODUCAO")
                End If
                Dim doc As XmlDocument = New XmlDocument()
                doc.PreserveWhitespace = True
                doc.LoadXml(xmlSitucao)
                xmlRetorno = HTTPSoapTextRequest(certificado, doc, url, idEmpresa, "http://tempuri.org/INfseServices/ConsultarSituacaoLoteRps")
                LogInfo("ConsultaLote processar retornoSituaçao: " + xmlRetorno)
                doc = New XmlDocument
                doc.LoadXml(xmlRetorno)
                Dim nodeConsultaSituacao As XmlNode = doc.SelectSingleNode(GetXpathFromString("Envelope, Body, ConsultarSituacaoLoteRpsResponse, ConsultarSituacaoLoteRpsResult"))
                nodeConsultaSituacao.InnerXml = nodeConsultaSituacao.InnerXml.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&#xD;", "")
                doc.LoadXml(nodeConsultaSituacao.InnerXml)
                Dim nodeConsultaSituacao1 As XmlNode = doc.SelectSingleNode(GetXpathFromString("ConsultarSituacaoLoteRpsResposta, Situacao"))
                Dim nodeConsultaSituacao2 As XmlNode = doc.SelectSingleNode(GetXpathFromString("ConsultarSituacaoLoteRpsResposta, ListaMensagemRetorno, MensagemRetorno, Codigo"))

                If nodeConsultaSituacao Is Nothing Then
                    Throw New FalhaException("Retorno da prefeitura não pode ser processado: objeto resposta veio nulo ou vazio. Retorno: " + xmlRetorno)
                End If

                If nodeConsultaSituacao2 IsNot Nothing Then
                    tipoConsulta = True
                    result = nodeConsultaSituacao.InnerXml
                    Return CClass.StringUtil.tirarAcentos(result.ToString)
                End If

                ' realizar a consulta do lote
                xmlRetorno = ""
                LogInfo("ConsultaLote processar retornoSituaçao 2: " + nodeConsultaSituacao1.InnerText)
                'If nodeConsultaSituacao1.InnerText = "1" Then
                '    Throw New FalhaException("LOTE NÃO ENVIADO.")
                'End If
                If nodeConsultaSituacao1.InnerText = "2" OrElse nodeConsultaSituacao1.InnerText = "1" Then
                    Return nodeConsultaSituacao.InnerXml.ToString
                End If
                If nodeConsultaSituacao1.InnerText = "3" OrElse nodeConsultaSituacao1.InnerText = "4" Then
                    Dim doc1 As XmlDocument = New XmlDocument()
                    doc1.PreserveWhitespace = True
                    doc1.LoadXml(xml)
                    xmlRetorno = HTTPSoapTextRequest(certificado, doc1, url, idEmpresa, "http://tempuri.org/INfseServices/ConsultarLoteRps")
                    LogInfo("ConsultaLote processar retornoLote: " + xmlRetorno)
                    doc1 = New XmlDocument
                    ' processar retorno, tirar as tags soap
                    doc1.LoadXml(xmlRetorno)

                    Dim nodeConsultaLote As XmlNode = doc1.SelectSingleNode(GetXpathFromString("Envelope, Body, ConsultarLoteRpsResponse,ConsultarLoteRpsResult"))
                    nodeConsultaLote.InnerXml = nodeConsultaLote.InnerXml.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&#xD;", "")
                    doc1.LoadXml(nodeConsultaLote.InnerXml)

                    If nodeConsultaLote Is Nothing Then
                        Throw New FalhaException("Retorno da prefeitura não pode ser processado: objeto resposta veio nulo ou vazio. Retorno: " + xmlRetorno)
                    End If

                    result = nodeConsultaLote.InnerXml
                End If

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

        Protected Overrides Function LoteEmProcessamento(ByVal xml As String) As Boolean
            ' E92	Esse RPS foi enviado para a nossa base de dados, mas ainda não foi processado	
            ' Faça uma nova consulta mais tarde.
            Dim retorno As Boolean = False
            Dim xmldoc As XmlDocument = New XmlDocument
            Try
                xmldoc.LoadXml(xml)

                Dim node As XmlNode = xmldoc.SelectSingleNode(GetXpathFromString("ConsultarSituacaoLoteRpsResposta, Situacao"))
                LogInfo("ConsultaLote LoteEmProcessamento 1 XML: " + xml)
                If node IsNot Nothing Then
                    LogInfo("ConsultaLote LoteEmProcessamento 2")
                    Dim codigo As String = node.InnerText.Trim()
                    If codigo = "2" OrElse codigo = "1" Then
                        LogInfo("ConsultaLote LoteEmProcessamento 3")
                        retorno = True
                    Else
                        LogInfo("ConsultaLote LoteEmProcessamento 4")
                        retorno = False
                    End If
                End If

                Return retorno
            Catch ex As Exception
                Throw New FalhaException("Erro ao ler XML de retorno de consulta de lotes: " + ex.Message)
            End Try
        End Function

        Protected Overrides Function GetXmlConsulta(ByVal protocolo As String, ByVal numrDocumento As Long, ByVal inscMunicipal As String) As String
            'If ambiente = TipoAmbiente.HOMOLOGACAO Then
            '    numrDocumento = CLng("2974456001230")
            '    inscMunicipal = "9096"
            'End If
            Dim xmlConsultaLote As String = _
            "<ConsultarLoteRpsEnvio xmlns=""http://www.abrasf.org.br/nfse"">" + _
               "<Prestador>" + _
                  "<Cnpj>" + numrDocumento.ToString("00000000000000") + "</Cnpj>" + _
                  "<InscricaoMunicipal>" + Regex.Replace(inscMunicipal, "[^0-9a-zA-Z]", "") + "</InscricaoMunicipal>" + _
               "</Prestador>" + _
               "<Protocolo>" + protocolo + "</Protocolo>" + _
            "</ConsultarLoteRpsEnvio>"

            Return xmlConsultaLote
        End Function

        Private Function GetXmlConsultaSituacaoLote(ByVal protocolo As String, ByVal numrDocumento As Long, ByVal inscMunicipal As String) As String
            'If ambiente = TipoAmbiente.HOMOLOGACAO Then
            '    numrDocumento = CLng("02974456001230")
            '    inscMunicipal = "9096"
            'End If
            Dim xmlConsultaSitucaoLote As String = _
                        "<ConsultarSituacaoLoteRpsEnvio xmlns=""http://www.abrasf.org.br/nfse"">" + _
                           "<Prestador>" + _
                              "<Cnpj>" + numrDocumento.ToString("00000000000000") + "</Cnpj>" + _
                              "<InscricaoMunicipal>" + Regex.Replace(inscMunicipal, "[^0-9a-zA-Z]", "") + "</InscricaoMunicipal>" + _
                           "</Prestador>" + _
                           "<Protocolo>" + protocolo + "</Protocolo>" + _
                        "</ConsultarSituacaoLoteRpsEnvio>"

            Return xmlConsultaSitucaoLote
        End Function

        Protected Overrides Function GetMotivoRejeicao() As String
            Dim xpath As String = APropriedadesXML.GetXpathFromCollection(New String() {"ConsultarLoteRpsResposta", "ListaMensagemRetorno", "MensagemRetorno"})
            Dim xpath2 As String = APropriedadesXML.GetXpathFromCollection(New String() {"ConsultarLoteRpsResposta", "ListaMensagemRetornoLote", "MensagemRetorno"})

            Dim resultrejeicao As String = GetMotivoRejeicao1(result, xpath)

            If resultrejeicao Is Nothing OrElse resultrejeicao.Trim.Length = 0 Then
                result = GetMotivoRejeicao1(result, xpath2)
            End If

            If resultrejeicao Is Nothing OrElse resultrejeicao.Trim.Length = 0 Then
                resultrejeicao = resultrejeicao
            End If

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
    End Class
End Namespace
