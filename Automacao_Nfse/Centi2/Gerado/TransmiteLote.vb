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
Imports System.Net

Namespace cCenti3
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


        Public Overrides Function processar(ByRef xml As String, ByVal codgCidade As Integer, ByVal numrCnpj As Long, ByVal idEmpresa As Long) As String
            Dim ambiente As TipoAmbiente
            Dim certificado As X509Certificate = Nothing
            Dim wsdl As String = Nothing
            Dim ws As Object = Nothing
            Dim usarProxy As Boolean = False
            Dim result As Object = Nothing
            Dim responseSOAP As New XmlDocument()
            Dim propCertificado As IPropriedadeCertificado = Nothing
            Dim url As String = Nothing
            Dim soapAction As String = ""
            Dim empresaData As NfseEmpresasData = GetIntegracaoEmpresa(numrCnpj)
            Dim xmlRetorno As String
            Dim usuario As String = ""
            Dim senha As String = ""

            Try
                ambiente = ParametroUtil.GetParametro(idEmpresa, "AMBIENTE")

                propCertificado = FactoryCore.GetPropriedadeCertificado(conn, trans, factory, publicVar, empresaData.codgCidade)
                certificado = propCertificado.GetCertificadoTransmissao(empresaData.idEmpresa, empresaData.numrCnpj)

                If certificado Is Nothing Then
                    Throw New FalhaException("Erro ao obter certificado digital, para TRANSMISSÃO, através do nº do documento informado : " + empresaData.numrCnpj.ToString)
                End If

                If ambiente = TipoAmbiente.HOMOLOGACAO Then
                    url = "https://api.centi.com.br/nfe/gerar/homologacao/GO/rioverde"
                    usuario = "fiscal@carpal.com.br"
                    senha = "carpal@2016"
                Else
                    url = "https://api.centi.com.br/nfe/gerar/" + empresaData.uf.ToUpper + "/" + empresaData.municipio.Replace(" ", "").ToLower
                    usuario = ParametroUtil.GetParametro(empresaData.idEmpresa, "USUARIO")
                    senha = ParametroUtil.GetParametro(empresaData.idEmpresa, "SENHA")
                End If


                Dim jsonEnvio As String = "{""usuario"": """ + usuario + """, ""senha"": """ + senha + """, ""xml"": """ + xml.Replace("""", "'") + """}"


                LogInfo(idEmpresa, "TransmiteLote processar 1: " + jsonEnvio)

                xmlRetorno = HTTPPostRequestJson(url, jsonEnvio, idEmpresa, numrCnpj)
                LogInfo(idEmpresa, "TransmiteLote processar 2: " + xmlRetorno)

                xmlRetorno = xmlRetorno.Replace("<?xml version=""1.0"" encoding=""utf-8""?>", "<?xml version=""1.0"" encoding=""ISO-8859-1""?>")
                xmlRetorno = (CClass.StringUtil.tirarAcentos(xmlRetorno))
                LogInfo(idEmpresa, "TransmiteLote processar 3 Tratado: " + xmlRetorno)


                Return xmlRetorno
            Catch ex As Exception
                LogInfo(idEmpresa, "TransmiteLote Exception: " + ex.Message)
                Dim docErro As New XmlDocument()
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

        End Function


        Public Overrides Function isRetornoEnvioOk(ByRef xml As String) As Boolean
            ' verificar se o retorno é XML
            Dim xmldoc As XmlDocument = New XmlDocument

            Try
                xmldoc.LoadXml(xml.Trim)
            Catch ex As Exception
                Throw New FalhaException("Erro ao processar retorno: O conteúdo retornado não é XML - " + ex.Message + "Conteúdo retornado: " + xml, ex)

            End Try

            Dim protocolo As String = XmlUtil.getValorTag(xmldoc, "Numero")

            If protocolo.Trim.Length = 0 Then
                Return False
            End If

            Return True
        End Function

        Public Overrides Function GetProtocolo(ByRef xml As String) As String
            ' verificar se o retorno é XML
            Dim xmldoc As XmlDocument = New XmlDocument

            Try
                xmldoc.LoadXml(xml.Trim())
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
            Dim xpath As String = APropriedadesXML.GetXpathFromCollection(New String() {"GerarNfseResposta", "ListaMensagemRetorno"})
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
                                result += Chr(13) + Chr(10) + cNode.Name + " - " + cNode.InnerText
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
