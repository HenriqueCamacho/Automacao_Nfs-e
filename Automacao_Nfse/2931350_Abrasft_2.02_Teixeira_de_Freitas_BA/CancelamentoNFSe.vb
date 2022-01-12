Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Security.Cryptography.X509Certificates
Imports CertificadoDigital
Imports System.text.RegularExpressions
Imports CClass
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Text
Imports System.Buffer
Imports System.Net
Imports NFSeCore

Namespace c2931350
    Public Class CancelamentoNFSe
        Inherits ACancelamento

        Public Sub New(ByRef conn As IDbConnection, _
                       ByRef trans As IDbTransaction, _
                       ByRef factory As DBFactory, _
                       ByRef publicVar As PublicVar)
            MyBase.factory = factory
            MyBase.publicVar = publicVar
            MyBase.conn = conn
            MyBase.trans = trans
        End Sub

        Public Overrides Function processar(ByVal codgCancelamento As Integer, _
                                                ByVal codgMunicipio As Integer, _
                                                ByVal movimentoData As NfseMovimentoData, _
                                                ByVal inscMunicipal As String, Optional ByVal usarUriRef As Boolean = True) As String
            Dim result As String = Nothing
            Dim propCertificado As IPropriedadeCertificado = Nothing
            Dim ambiente As TipoAmbiente = ParametroUtil.GetParametro(movimentoData.idEmpresa, "AMBIENTE")

            Try
                Dim templateFilePath As String = FactoryCore.GetTemplateCancelamento(codgMunicipio)

                documentoXML = New Xml.XmlDataDocument()
                documentoXML.DataSet.ReadXml(templateFilePath)
            Catch ex As Exception
                Throw New FalhaException("Erro ao processar arquivo XSD para cidade: " + codgMunicipio.ToString, ex)
            End Try

            dsCancelamento = documentoXML.DataSet

            dtIdentificacaoNfse = dsCancelamento.Tables("IdentificacaoNfse")
            dtInfPedidoCancelamento = dsCancelamento.Tables("InfPedidoCancelamento")
            dtCpfCnpj = dsCancelamento.Tables("CpfCnpj")
            LogInfo("Make Cancelamento de NFSE: 1")
            adicionaValor("Id", dtInfPedidoCancelamento, "NFSE" + GetIdCancelamento())
            adicionaValor("CodigoCancelamento", dtInfPedidoCancelamento, 1)
            adicionaValor("Numero", dtIdentificacaoNfse, movimentoData.numrNfse)
            LogInfo("Make Cancelamento de NFSE: 2")
            adicionaValor("Cnpj", dtCpfCnpj, movimentoData.numrCnpj.ToString("00000000000000"))
            adicionaValor("InscricaoMunicipal", dtIdentificacaoNfse, Regex.Replace(inscMunicipal, "[^0-9a-zA-Z]", ""))
            adicionaValor("CodigoMunicipio", dtIdentificacaoNfse, codgMunicipio)
            LogInfo("Make Cancelamento de NFSE: 3")
            xmlCancelamento = dsToXML(dsCancelamento)
            LogInfo("Make Cancelamento de NFSE: 4")
            Dim PropriedadesXML As IPropriedadesXML = FactoryCore.GetPropriedadesXML(codgMunicipio)
            Dim servico As NfseServicoData = GetIntegracaoServico(movimentoData.tipoRps, movimentoData.serieRps, movimentoData.numeroRps, movimentoData.numrCnpj)
            LogInfo("Make Cancelamento de NFSE: 5")
            ' obter a classe especifica para gerenciar o certificado para esta cidade
            propCertificado = FactoryCore.GetPropriedadeCertificado(conn, trans, factory, publicVar, codgMunicipio)
            xmlCancelamento = ordernarXmlCancelamento(xmlCancelamento)
            xmlCancelamento = propCertificado.assinarXMLCancelamento(xmlCancelamento, movimentoData.idEmpresa, codgMunicipio)
            xmlCancelamento = "<Envelope xmlns=""http://schemas.xmlsoap.org/soap/envelope/"">" + _
                                  "<Header/>" + _
                                  "<Body>" + _
                                      "<CancelarNfse xmlns=""http://trbteixeira.kbfsistemas.com.br/webrun/webservices/NFEServices.jws"">" + _
                                          "<Nfsecabecmsg><![CDATA[" + PropriedadesXML.GetCabecalho() + "]]></Nfsecabecmsg>" + _
                                          "<Nfsedadosmsg><![CDATA[" + xmlCancelamento + "]]></Nfsedadosmsg>" + _
                                      "</CancelarNfse>" + _
                                  "</Body>" + _
                              "</Envelope>"
            ' transmitir
            result = transmitirCancelamento(codgMunicipio, movimentoData.numrCnpj, xmlCancelamento, movimentoData.idEmpresa)
            LogInfo("Make Cancelamento de NFSE: 6" + xmlCancelamento)
            Return result
        End Function

        Protected Overrides Function transmitirCancelamento(ByVal codgCidade As Integer, ByVal numrCnpj As Long, ByVal xml As String, ByVal idempresa As Integer) As String
            Dim wsdl As String = Nothing
            Dim ambiente As TipoAmbiente = Nothing
            Dim certificado As X509Certificate = Nothing
            Dim result As String = ""
            Dim xmlRetorno As String = Nothing
            Dim responseSOAP As New XmlDocument()
            Dim propCertificado As IPropriedadeCertificado = Nothing
            Dim url As String = Nothing
            Dim cidadeData As NfseCidadesData = GetCidade(codgCidade)

            Try
                ' obter o ambiente producao ou homologacao
                ambiente = CInt(ParametroUtil.GetParametro(idempresa, "AMBIENTE"))
                LogInfo("Make Cancelamento de NFSE: 7")
                ' obter a classe especifica para gerenciar o certificado para esta cidade
                propCertificado = FactoryCore.GetPropriedadeCertificado(conn, trans, factory, publicVar, codgCidade)

                certificado = propCertificado.GetCertificadoTransmissao(idempresa, numrCnpj)
                If certificado Is Nothing Then
                    Throw New FalhaException("Erro ao obter certificado digital, para TRANSMISSÃO, através do nº do documento informado : " + numrCnpj.ToString)
                End If
                LogInfo("Make Cancelamento de NFSE: 8")
                If ambiente = TipoAmbiente.HOMOLOGACAO Then
                    url = ParametroUtil.GetParametroCidade(codgCidade, "URL_AMBIENTE_HOMOLOGACAO")
                Else
                    url = ParametroUtil.GetParametroCidade(codgCidade, "URL_AMBIENTE_PRODUCAO")
                End If
                LogInfo("Make Cancelamento de NFSE: 9")
                Dim doc As New XmlDocument
                doc.PreserveWhitespace = True
                doc.LoadXml(xmlCancelamento)
                LogInfo("Make Cancelamento de NFSE: 9.1" + xmlCancelamento)
                xmlRetorno = HTTPSoapTextRequest(certificado, doc, url, idempresa, "")
                responseSOAP.LoadXml(xmlRetorno)
                LogInfo("Make Cancelamento de NFSE: 10" + xmlRetorno)
                Dim node As XmlNode = responseSOAP.SelectSingleNode(GetXpathFromString("Envelope, Body, CancelarNfseResponse, CancelarNfseReturn"))
                node.InnerXml = node.InnerXml.Replace("<![CDATA[", "").Replace("]]>", "").Replace("&lt;", "<").Replace("&gt;", ">")
                If node Is Nothing Then
                    Throw New FalhaException("Retorno do cancelamento da prefeitura não pode ser processado: objeto veio nulo. Retorno: " + xmlRetorno)
                End If
                LogInfo("Make Cancelamento de NFSE: 11")
                result = node.InnerXml
                xmlCancelamentoRetorno = result.ToString
                LogInfo("Make Cancelamento de NFSE: 12")
                Return result.ToString
            Catch ex As Exception
                If TypeOf ex Is FalhaException Then
                    Throw ex
                Else
                    Throw New FalhaException("Erro ao tentar transmitir cancelamento de NFSE: " + ex.Message, ex)
                End If
            End Try
            LogInfo("Make Cancelamento de NFSE: 13")
        End Function

        Protected Overrides Function IsCancelamentoHomologado() As Boolean
            Dim xmldoc As New Xml.XmlDocument

            Try
                xmldoc.LoadXml(xmlCancelamentoRetorno)
                LogInfo("Make Cancelamento de NFSE: 14")
                ' se a tag DataHoraCancelamento existir eh prq o cancelamento foi homologado
                Dim xpath As String = APropriedadesXML.GetXpathFromCollection(New String() _
                            {"CancelarNfseResposta", "RetCancelamento", "NfseCancelamento", "Confirmacao", "DataHora"})
                LogInfo("Make Cancelamento de NFSE: 16")
                Dim xpath2 As String = APropriedadesXML.GetXpathFromCollection(New String() _
                            {"CancelarNfseResposta", "ListaMensagemRetorno", "MensagemRetorno", "Codigo"})
                LogInfo("Make Cancelamento de NFSE: 17")
                If XmlUtil.getDocByTagXPath(xmldoc, xpath) Is Nothing Then
                    Dim codigo As String = XmlUtil.getValorTagXPath(xmldoc, xpath2)
                    LogInfo("Make Cancelamento de NFSE: 18: " + codigo)
                    Return codigo.Trim = "E79" ' codigo E79 = Essa NFS-e já está cancelada
                End If
                LogInfo("Make Cancelamento de NFSE: 19")
                Return XmlUtil.getDocByTagXPath(xmldoc, xpath) IsNot Nothing
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Function ordernarXmlCancelamento(ByVal xml As String) As String
            Dim doc As New XmlDocument()
            doc.LoadXml(xml)

            Dim xPathInf = "/*[local-name()='CancelarNfseEnvio']/*[local-name()='Pedido']/*[local-name()='InfPedidoCancelamento']/*[local-name()='"

            moveAfterTag(doc, xPathInf + "CodigoCancelamento']", xPathInf + "IdentificacaoNfse']")
            moveAfterTag(doc, xPathInf + "IdentificacaoNfse']/*[local-name()='CpfCnpj']", xPathInf + "IdentificacaoNfse']/*[local-name()='Numero']")

            Return doc.OuterXml
        End Function
    End Class
End Namespace
