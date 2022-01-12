Imports System.Text
Imports NFSeCore

Namespace cCenti3
    Public Class PropriedadesXML
        Inherits APropriedadesXML

        Private Shared nsXsdLoaded As Boolean = False

        Public Sub New()
        End Sub

        Public Overrides Function GetFirstTagLote() As String
            Return "p:EnviarLoteRpsEnvio"
        End Function

        Public Overrides Function GetFirstTagRps() As String
            Return "Rps"
        End Function

        Public Overrides Function GetSecondTagLote() As String
            Return "p:LoteRps"
        End Function

        Public Overrides Function GetSecondTagRps() As String
            Return "InfDeclaracaoPrestacaoServico"
        End Function

        Public Overrides Function GetSTagCanc() As String
            Return "p1:InfPedidoCancelamento"
        End Function

        Public Overrides Function GetXmlConsultaNfseRps(ByVal tipoRps As String, ByVal numeroRps As String, ByVal serieRps As String, ByVal tipoDocumento As Integer, ByVal numrDocumento As String, ByVal inscMunicipal As String) As String
            Dim result As New StringBuilder()
            result.Append("<?xml version=""1.0"" encoding=""UTF-8""?>")
            result.Append("<p:ConsultarNfseRpsEnvio xmlns:ds=""http://www.w3.org/2000/09/xmldsig#"" xmlns:p=""http://ws.speedgov.com.br/consultar_nfse_rps_envio_v1.xsd"" xmlns:p1=""http://ws.speedgov.com.br/tipos_v1.xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://ws.speedgov.com.br/consultar_nfse_rps_envio_v1.xsd consultar_nfse_rps_envio_v1.xsd "">")
            result.Append("<p:IdentificacaoRps>")
            result.Append("<p1:Numero>" + numeroRps + "</p1:Numero>")
            result.Append("<p1:Serie>" + serieRps + "</p1:Serie>")
            result.Append("<p1:Tipo>" + tipoRps + "</p1:Tipo>")
            result.Append("</p:IdentificacaoRps>")
            result.Append("<p:Prestador>")
            result.Append("<p1:Cnpj>" + numrDocumento + "</p1:Cnpj>")
            result.Append("<p1:InscricaoMunicipal>" + inscMunicipal + "</p1:InscricaoMunicipal>")
            result.Append("</p:Prestador>")
            result.Append("</p:ConsultarNfseRpsEnvio>")
            Return result.ToString()
        End Function

        Public Overrides Function GetCabecalho() As String
            Dim result As String = "<?xml version=""1.0"" encoding=""UTF-8""?><p:cabecalho versao=""1"" xmlns:ds=""http://www.w3.org/2000/09/xmldsig#"" xmlns:p=""http://ws.speedgov.com.br/cabecalho_v1.xsd"" xmlns:p1=""http://ws.speedgov.com.br/tipos_v1.xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://ws.speedgov.com.br/cabecalho_v1.xsd cabecalho_v1.xsd ""><versaoDados>1</versaoDados></p:cabecalho>"
            Return result
        End Function

        Public Overrides Function GetXsdNamespaceMap(ByVal codgCidade As Integer) As System.Collections.Hashtable
            ' obter o caminho dos arquivos XSD
            Dim xsdnfse As String = FactoryCore.GetFilePath(codgCidade, TipoArquivo.schemas) + "nfse.xsd"
            Dim urinfse As String = GetDefaultNamespace(xsdnfse)
            _nsXsd.Add(urinfse, xsdnfse)

            'xsdnfse = FactoryCore.GetFilePath(codgCidade, TipoArquivo.schemas) + "tipos_v1.xsd"
            'urinfse = GetDefaultNamespace(xsdnfse)
            '_nsXsd.Add(urinfse, xsdnfse)

            'xsdnfse = FactoryCore.GetFilePath(codgCidade, TipoArquivo.schemas) + "xmldsig-core-schema20020212.xsd"
            'urinfse = GetDefaultNamespace(xsdnfse)
            '_nsXsd.Add(urinfse, xsdnfse)
            Return _nsXsd
        End Function

    End Class
End Namespace
