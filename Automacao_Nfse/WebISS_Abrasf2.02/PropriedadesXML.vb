Imports System.Text
Imports NFSeCore

Namespace cWebISS2
    Public Class PropriedadesXML
        Inherits APropriedadesXML

        Private Shared nsXsdLoaded As Boolean = False

        Public Sub New()
        End Sub

        Public Overrides Function GetFirstTagLote() As String ' nao usa
            Return ""
        End Function

        Public Overrides Function GetFirstTagRps() As String
            Return "Rps"
        End Function

        Public Overrides Function GetSecondTagLote() As String ' nao usa
            Return ""
        End Function

        Public Overrides Function GetSecondTagRps() As String
            Return "InfDeclaracaoPrestacaoServico"
        End Function

        Public Overrides Function GetSTagCanc() As String
            Return "InfPedidoCancelamento"
        End Function
        Public Overrides Function GetXsdNamespaceMap(ByVal codgCidade As Integer) As System.Collections.Hashtable
            ' obter o caminho dos arquivos XSD
            Dim xsdnfse As String = FactoryCore.GetFilePath(codgCidade, TipoArquivo.schemas) + "nfse v2 02.xsd"
            Dim urinfse As String = GetDefaultNamespace(xsdnfse)

            _nsXsd.Add(urinfse, xsdnfse)

            Return _nsXsd
        End Function

        Public Overrides Function GetXmlConsultaNfseRps(ByVal tipoRps As String, ByVal numeroRps As String, ByVal serieRps As String, ByVal tipoDocumento As Integer, ByVal numrDocumento As String, ByVal inscMunicipal As String) As String
            Dim result As New StringBuilder()

            result.Append("<ConsultarNfseRpsEnvio xmlns=""http://www.abrasf.org.br/nfse.xsd"">")
            result.Append("<IdentificacaoRps>")
            result.Append("<Numero>" + numeroRps + "</Numero>")
            result.Append("<Serie>" + serieRps + "</Serie>")
            result.Append("<Tipo>" + tipoRps + "</Tipo>")
            result.Append("</IdentificacaoRps>")
            result.Append("<Prestador>")
            result.Append("<CpfCnpj>")
            result.Append("<Cnpj>" + numrDocumento + "</Cnpj>")
            result.Append("</CpfCnpj>")
            result.Append("<InscricaoMunicipal>" + inscMunicipal + "</InscricaoMunicipal>")
            result.Append("</Prestador>")
            result.Append("</ConsultarNfseRpsEnvio>")

            Return result.ToString()
        End Function

        Public Overrides Function GetCabecalho() As String
            Dim result As String = "<cabecalho versao=""2.02"" xmlns=""http://www.abrasf.org.br/nfse.xsd"">" + _
                                              "<versaoDados>2.02</versaoDados>" + _
                                            "</cabecalho>"
            Return result
        End Function
    End Class
End Namespace
