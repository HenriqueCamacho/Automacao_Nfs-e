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

Namespace c2931350
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

        Public Overrides Sub processar(ByRef movimento As NfseMovimentoData)
            Dim ambiente As TipoAmbiente = Nothing
            Dim certificado As X509Certificate = Nothing
            Dim usarProxy As Boolean = False
            Dim result As Object = Nothing
            Dim responseSOAP As New XmlDocument()
            Dim propCertificado As IPropriedadeCertificado = Nothing
            Dim url As String = Nothing
            Dim soapAction As String = ""
            Dim empresaData As NfseEmpresasData = GetIntegracaoEmpresa(movimento.numrCnpj)
            Dim xmlRetorno As String
            Dim xml As String

            Try
                ' obter o ambiente producao ou homologacao
                ambiente = ParametroUtil.GetParametro(empresaData.idEmpresa, "AMBIENTE")

                ' obter a classe especifica para gerenciar o certificado para esta cidade
                propCertificado = FactoryCore.GetPropriedadeCertificado(conn, trans, factory, publicVar, empresaData.codgCidade)
                certificado = propCertificado.GetCertificadoTransmissao(empresaData.idEmpresa, empresaData.numrCnpj)

                If certificado Is Nothing Then
                    Throw New FalhaException("Erro ao obter certificado digital, para TRANSMISSÃO, através do nº do documento informado : " + empresaData.numrCnpj.ToString)
                End If

                If ambiente = TipoAmbiente.HOMOLOGACAO Then
                    url = ParametroUtil.GetParametroCidade(empresaData.codgCidade, "URL_AMBIENTE_HOMOLOGACAO")
                Else
                    url = ParametroUtil.GetParametroCidade(empresaData.codgCidade, "URL_AMBIENTE_PRODUCAO")
                End If

                Dim PropriedadesXML As IPropriedadesXML = FactoryCore.GetPropriedadesXML(empresaData.codgCidade)
                'xml = "<?xml version=""1.0"" encoding=""utf-8""?>" + movimento.xmlNota
                xml = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" + movimento.xmlNota

                'xml = "<Envelope xmlns=""http://schemas.xmlsoap.org/soap/envelope/"">" + _
                '            "<Header/>" + _
                '            "<Body>" + _
                '                "<GerarNfse xmlns=""http://trbteixeira.kbfsistemas.com.br/webrun/webservices/NFEServices.jws"">" + _
                '                    "<Nfsecabecmsg><![CDATA[" + PropriedadesXML.GetCabecalho() + "]]></Nfsecabecmsg>" + _
                '                    "<Nfsedadosmsg><![CDATA[<?xml version=""1.0"" encoding=""UTF-8""?>" + xml + "]]></Nfsedadosmsg>" + _
                '                "</GerarNfse>" + _
                '            "</Body>" + _
                '        "</Envelope>"

                xml = "<Envelope xmlns=""http://schemas.xmlsoap.org/soap/envelope/"">" + _
            "<Header/>" + _
            "<Body>" + _
                "<GerarNfse xmlns=""http://nfse.eunapolis.ba.gov.br/webrun/webservices/NFEServices.jws"">" + _
                    "<Nfsecabecmsg><![CDATA[" + PropriedadesXML.GetCabecalho() + "]]></Nfsecabecmsg>" + _
                    "<Nfsedadosmsg><![CDATA[" + xml + "]]></Nfsedadosmsg>" + _
                "</GerarNfse>" + _
            "</Body>" + _
        "</Envelope>"

                'Se for homologação muda o xmlns do GerarNfse
                If ambiente = TipoAmbiente.HOMOLOGACAO Then
                    xml = xml.Replace("<GerarNfse xmlns=""http://nfse.eunapolis.ba.gov.br/webrun/webservices/NFEServices.jws"">", "<GerarNfse xmlns=""http://177.93.108.119:4040/nfsews/webservices/NFEServices.jws"">")
                End If

                LogInfo(empresaData.idEmpresa, "TransmiteLote XML envio: " + xml)
                Dim doc As XmlDocument = New XmlDocument()
                doc.PreserveWhitespace = True
                doc.LoadXml(xml)
                xmlRetorno = HTTPSoapTextRequest(certificado, doc, url, empresaData.idEmpresa, "")
                LogInfo(empresaData.idEmpresa, "TransmiteLote XML retorno: " + xmlRetorno)
                responseSOAP.LoadXml(xmlRetorno)
                Dim node As XmlNode = responseSOAP.SelectSingleNode(GetXpathFromString("Envelope, Body, GerarNfseResponse, GerarNfseReturn"))
                node.InnerXml = node.InnerXml.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&#xD;", "")
                node.InnerXml = CClass.StringUtil.tirarAcentos(node.InnerXml)

                'Para a cidade Lauro de freitas, no primeiro retorno ele já retorna que o rps já existe
                If node Is Nothing Then
                    Throw New FalhaException("Retorno da transmissão do lote a prefeitura não pode ser processado: objeto veio nulo. Retorno: " + xmlRetorno)
                End If

                movimento.xmlRetorno = node.InnerXml
            Catch ex As Exception
                Dim docErro As XmlDocument = New XmlDocument()
                Try
                    docErro.LoadXml(ex.InnerException.Message)
                Catch e As Exception
                    Throw New Exception(ex.ToString)
                End Try
                If XmlUtil.getDocByTag(docErro, "faultstring") IsNot Nothing Then
                    Throw New FalhaException("Erro no processo da PREFEITURA: " + XmlUtil.getValorTag(docErro, "faultstring"))
                End If

                If TypeOf ex Is FalhaException Then
                    Throw ex
                Else
                    Throw New FalhaException("Erro ao tentar transmitir lote de RPS: " + ex.Message, ex)
                End If
            End Try

        End Sub

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
