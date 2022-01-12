Imports CClass
Imports System.Security.Cryptography.X509Certificates
Imports System.Web.Services
Imports System.Xml
Imports System.Text.RegularExpressions
Imports CertificadoDigital
Imports NFSeWebService
Imports System.IO
Imports NFSeCore

Namespace cRlz
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
                xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" + movimento.xmlNota '
                xml = "<soapenv:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:urn=""urn:server.issqn"">" + _
                       "<soapenv:Header/>" + _
                       "<soapenv:Body>" + _
                             "<urn:gravaNotaXML soapenv:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">" + _
                                "<params xsi:type=""xsd:string""><![CDATA[" + xml + "]]></params>" + _
                             "</urn:gravaNotaXML>" + _
                        "</soapenv:Body>" + _
                        "</soapenv:Envelope>"
                LogInfo(empresaData.idEmpresa, "TransmiteLote XML envio: " + xml)
                Dim doc As XmlDocument = New XmlDocument()
                doc.PreserveWhitespace = False
                doc.LoadXml(xml)
                xmlRetorno = HTTPSoapTextRequest(certificado, doc, url, empresaData.idEmpresa, "urn:server.issqn#gravaNotaXML")
                LogInfo(empresaData.idEmpresa, "TransmiteLote XML retorno: " + xmlRetorno)
                responseSOAP.LoadXml(xmlRetorno)
                Dim node As XmlNode = responseSOAP.SelectSingleNode(GetXpathFromString("Envelope, Body, gravaNotaXMLResponse, return"))
                node.InnerXml = node.InnerXml.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&#xD;", "")
                node.InnerXml = CClass.StringUtil.tirarAcentos(node.InnerXml)
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

        Public Overrides Function isRetornoEnvioOk(ByRef xml As String) As Boolean
            ' verificar se o retorno é XML
            Dim xmldoc As XmlDocument = New XmlDocument

            Try
                xmldoc.LoadXml(xml)
            Catch ex As Exception
                Throw New FalhaException("Erro ao processar retorno: O conteúdo retornado não é XML - " + ex.Message + "Conteúdo retornado: " + xml, ex)

            End Try

            Dim protocolo As String = XmlUtil.getValorTag(xmldoc, "url")

            If protocolo.Trim.Length = 0 Then
                Return False
            End If

            Return True
        End Function

        Public Overrides Function GetProtocolo(ByRef xml As String) As String
            ' verificar se o retorno é XML
            Dim xmldoc As XmlDocument = New XmlDocument

            Try
                xmldoc.LoadXml(xml)
            Catch ex As Exception
                Throw New FalhaException("Erro ao processar retorno: O conteúdo retornado não é XML - " + ex.Message, ex)
            End Try

            Dim protocolo As String = XmlUtil.getValorTag(xmldoc, "Protocolo")

            If protocolo.Length = 0 Then
                Throw New FalhaException("Erro ao processar retorno: O conteúdo retornado não está no formato esperado: " + xml)
            End If

            Return protocolo
        End Function

        Protected Overrides Function GetMotivoRejeicao(ByVal xmlRetorno As String) As String
            Dim xpath As String = APropriedadesXML.GetXpathFromCollection(New String() {"nota"})
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
                                result += Chr(13) + Chr(10) + cNode.Name + " : " + cNode.InnerText
                            Next
                        Else
                            result += node.Name + " : " + node.InnerText
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
