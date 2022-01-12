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

Namespace cCenti2
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
            Dim cancelamento As NfseCancelamentoData = GetIntegracaoCancelamento(movimentoData.idMovimento)
            Dim InfNfse As String = Nothing
            Dim codVer As String = Nothing
            Try
                InfNfse = getInfNfse(movimentoData.xmlNota)
                codVer = getCodigoVerificacao(movimentoData.xmlNota)
            Catch ex As Exception
                Throw New FalhaException("XML Nota salva na movimento esta invalida. " + ex.Message, ex)
            End Try

            'Gerar XML
            xmlCancelamento = "<?xml version=\""1.0\"" encoding=\""UTF-8\""?>" + _
                              "<CancelarNfseEnvio xmlns=\""http://www.centi.com.br/files/nfse.xsd\"" xmlns:xsd=\""http://www.w3.org/2001/XMLSchema\"" xmlns:xsi=\""http://www.w3.org/2001/XMLSchema-instance\"">" + _
                                  "<Pedido>" + _
                                      "<InfPedidoCancelamento>" + _
                                          "<IdentificacaoNfse>" + _
                                              "<Id>" & InfNfse & "</Id>" + _
                                              "<Numero>" & movimentoData.numrNfse & "</Numero>" + _
                                              "<CpfCnpj>" + _
                                                  "<Cnpj>" & movimentoData.numrCnpj.ToString("00000000000000") & "</Cnpj>" + _
                                              "</CpfCnpj>" + _
                                              "<InscricaoMunicipal>" & Regex.Replace(inscMunicipal, "[^0-9a-zA-Z]", "") & "</InscricaoMunicipal>" + _
                                              "<CodigoMunicipio>5218805</CodigoMunicipio>" + _
                                              "<CodigoVerificacao>" & codVer & "</CodigoVerificacao>" + _
                                              "<DescricaoCancelamento>" & cancelamento.motivoCancelamento & "</DescricaoCancelamento>" + _
                                          "</IdentificacaoNfse>" + _
                                          "<CodigoCancelamento>" & codgCancelamento & "</CodigoCancelamento>" + _
                                      "</InfPedidoCancelamento>" + _
                                  "</Pedido>" + _
                              "</CancelarNfseEnvio>"

            LogInfo("CancelamentoNFSe processar 1: " & xmlCancelamento)            

            Dim PropriedadesXML As IPropriedadesXML = FactoryCore.GetPropriedadesXML(codgMunicipio)
            Dim servico As NfseServicoData = GetIntegracaoServico(movimentoData.tipoRps, movimentoData.serieRps, movimentoData.numeroRps, movimentoData.numrCnpj)

            propCertificado = FactoryCore.GetPropriedadeCertificado(conn, trans, factory, publicVar, codgMunicipio)

            Dim usuario As String = Nothing
            Dim senha As String = Nothing

            Try
                usuario = ParametroUtil.GetParametro(movimentoData.idEmpresa, "USUARIO")
                senha = ParametroUtil.GetParametro(movimentoData.idEmpresa, "SENHA")
            Catch ex As Exception
                Throw New FalhaException("Parametrize o usuario e senha no NFSe Service. ")
            End Try

            LogInfo("CancelamentoNFSe processar 2: " & xmlCancelamento)
            result = transmitirCancelamento(codgMunicipio, movimentoData.numrCnpj, xmlCancelamento, movimentoData.idEmpresa)

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
            Dim empresaData As NfseEmpresasData = GetIntegracaoEmpresa(idempresa)

            Try
                ambiente = CInt(ParametroUtil.GetParametro(idempresa, "AMBIENTE"))

                propCertificado = FactoryCore.GetPropriedadeCertificado(conn, trans, factory, publicVar, codgCidade)

                certificado = propCertificado.GetCertificadoTransmissao(idempresa, numrCnpj)
                If certificado Is Nothing Then
                    Throw New FalhaException("Erro ao obter certificado digital, para TRANSMISS�O, atrav�s do n� do documento informado : " + numrCnpj.ToString)
                End If

                Dim usuario As String = Nothing
                Dim senha As String = Nothing

                Try
                    usuario = ParametroUtil.GetParametro(idempresa, "USUARIO")
                    senha = ParametroUtil.GetParametro(idempresa, "SENHA")
                Catch ex As Exception
                    Throw New FalhaException("Parametrize o usuario e senha no NFSe Service. ")
                End Try

                If ambiente = TipoAmbiente.HOMOLOGACAO Then
                    url = ParametroUtil.GetParametroCidade(codgCidade, "URL_AMBIENTE_HOMOLOGACAO")
                Else
                    url = "https://api.centi.com.br/nfe/cancelar/" + empresaData.uf.ToUpper + "/" + empresaData.municipio.Replace(" ", "").ToLower

                End If

                Dim jsonEnvio As String = "{""usuario"": """ + usuario + """, ""senha"": """ + senha + """, ""xml"": """ + xml + """}"

                LogInfo(idempresa, "Cancela JsonEnvio: " + jsonEnvio)
                LogInfo(idempresa, "Cancela URL: " + url)

                xmlRetorno = HTTPPostRequestJson(url, jsonEnvio, idempresa, numrCnpj)
                LogInfo(idempresa, "Cancela xmlRetorno crú: " + xmlRetorno)

                xmlRetorno = (CClass.StringUtil.tirarAcentos(xmlRetorno))
                LogInfo(idempresa, "Cancela xmlRetorno sem acentos: " + xmlRetorno)

                xmlRetorno = xmlRetorno.Replace("<?xml version=""1.0"" encoding=""utf-8""?>", "")

                xmlCancelamentoRetorno = xmlRetorno

                Return xmlRetorno
            Catch ex As Exception
                If TypeOf ex Is FalhaException Then
                    Throw ex
                Else
                    Throw New FalhaException("Erro ao tentar transmitir cancelamento de NFSE: " + ex.Message, ex)
                End If
            End Try
        End Function

        Protected Overrides Function IsCancelamentoHomologado() As Boolean
            Dim xmldoc As New Xml.XmlDocument

            Try
                xmldoc.LoadXml(xmlCancelamentoRetorno)
                LogInfo("xmlCancelamentoRetorno: " + xmlCancelamentoRetorno.ToUpper)
                If xmlCancelamentoRetorno.ToUpper.Contains("NF JA ENCONTRA-SE CANCELADA.") Then
                    Return True
                End If
                ' se a tag DataHoraCancelamento existir eh prq o cancelamento foi homologado
                Dim xpath As String = APropriedadesXML.GetXpathFromCollection(New String() _
                            {"GerarNfseResposta", "ListaNfse", "CompNfse", "Nfse", "InfNfse", "Status"})

                If XmlUtil.getDocByTagXPath(xmldoc, xpath) IsNot Nothing Then
                    Dim status As String = XmlUtil.getValorTagXPath(xmldoc, xpath)

                    Return status = "3" 'Essa NFS-e j� est� cancelada
                End If

                Return XmlUtil.getDocByTagXPath(xmldoc, xpath) IsNot Nothing
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Function ordernarXmlCancelamento(ByVal xml As String) As String
            Dim doc As New XmlDocument()
            doc.LoadXml(xml)

            Dim xPathInf = "/*[local-name()='GerarNfseResposta']/*[local-name()='Pedido']/*[local-name()='InfPedidoCancelamento']/*[local-name()='"
            moveAfterTag(doc, xPathInf + "CodigoCancelamento']", xPathInf + "IdentificacaoNfse']")
            moveAfterTag(doc, xPathInf + "IdentificacaoNfse']/*[local-name()='InscricaoMunicipal']", xPathInf + "IdentificacaoNfse']/*[local-name()='CpfCnpj']")
            moveAfterTag(doc, xPathInf + "IdentificacaoNfse']/*[local-name()='CodigoMunicipio']", xPathInf + "IdentificacaoNfse']/*[local-name()='InscricaoMunicipal']")

            Return doc.OuterXml
        End Function


        Private Function getInfNfse(ByVal nota As String) As String
            Dim doc As New XmlDocument()
            doc.LoadXml(nota)
            Return CClass.XmlUtil.getValorAtributo(doc, "InfNfse", "Id")
        End Function

        Private Function getCodigoVerificacao(ByVal nota As String) As String
            Dim doc As New XmlDocument()
            doc.LoadXml(nota)
            Return XmlUtil.getValorTag(doc, "CodigoVerificacao")

        End Function
    End Class
End Namespace
