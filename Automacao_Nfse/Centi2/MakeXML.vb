Imports CClass
Imports System.Xml
Imports System.Text.RegularExpressions
Imports NFSeCore
Imports System.Text

Namespace cCenti2
    Public Class MakeXML
        Inherits AMakeXML


        Public Sub New(ByRef conn As IDbConnection, _
                       ByRef trans As IDbTransaction, _
                       ByRef factory As DBFactory, _
                       ByRef publicVar As PublicVar)
            MyBase.factory = factory
            MyBase.publicVar = publicVar
            MyBase.conn = conn
            MyBase.trans = trans
        End Sub
        ''' <summary>
        ''' Interpreta os dados da integração e os converte para um XML de RPS ja 
        ''' em formato de LOTE
        ''' </summary>
        ''' <param name="capa">Dados da capa da nota</param>
        ''' <param name="servico">Dados do servico referenciado na nota</param>
        ''' <param name="tomador">Dados do tomador do servico</param>
        ''' <returns>String - XML do lote de RPS</returns>
        Public Overrides Function processar(ByRef capa As NfseCapaData, ByRef servico As NfseServicoData, ByRef tomador As NfseTomadorData, ByRef substituido As NfseSubstituidoData, ByVal idempresa As Integer) As String

            Dim result As String
            Try
                result = popularXML(capa, servico, tomador, substituido)
            Catch ex As Exception
                Throw New FalhaException("Erro ao obter dados da integracao: " + ex.Message, ex)
            End Try

            result = CClass.Util.TString(result).Replace("&", "").Replace("#", "").Replace("¿", "")
            LogInfo("MakeXML processar result: " + result)
            Return CClass.StringUtil.tirarAcentos(result)

        End Function

        ''' <summary>
        ''' Atraves dos dados provenientes da camada de integracao é preenchido o dataset que sera
        ''' convertido em XML
        ''' </summary>
        Private Function popularXML(ByRef capa As NfseCapaData, ByRef servico As NfseServicoData, ByRef tomador As NfseTomadorData, ByRef substituido As NfseSubstituidoData) As String
            Dim xmlStream As New StringBuilder()
            Dim empresaData As NfseEmpresasData = GetIntegracaoEmpresa(capa.numrCnpj)
            '
            Dim nfseEmpresaData As New NfseEmpresasData
            Dim nfseservicoDataRN As New NfseEmpresasRN(factory, publicVar, Nothing)
            '
            Dim ambiente As TipoAmbiente = ParametroUtil.GetParametro(empresaData.idEmpresa, "AMBIENTE")

            nfseEmpresaData = GetIntegracaoEmpresa(capa.numrCnpj)

            Dim serieNfse As String
            Try
                serieNfse = ParametroUtil.GetParametro(CInt(nfseEmpresaData.idEmpresa), "SERIE_NFSE")
                If serieNfse = "" Then
                    serieNfse = capa.serieRps
                End If
            Catch ex As Exception
                serieNfse = capa.serieRps
            End Try

            LogInfo("MakeXML processar popularXML.: 1")

            'xmlStream.Append("<?xml version=""1.0"" encoding=""UTF-8""?>")

            xmlStream.Append("<GerarNfseEnvio xmlns=""http://www.centi.com.br/files/nfse.xsd"">")
            xmlStream.Append("<Rps>")
            xmlStream.Append("<InfDeclaracaoPrestacaoServico>")
            xmlStream.Append("<Rps Id=""lote" & GetNumeroLote().ToString & """>")
            xmlStream.Append("<IdentificacaoRps>")
            xmlStream.Append("<Numero>" & capa.numeroRps & "</Numero>")
            xmlStream.Append("<Serie>" & capa.serieRps & "</Serie>")
            xmlStream.Append("<Tipo>" & capa.tipoRps & "</Tipo>")
            xmlStream.Append("</IdentificacaoRps>")
            xmlStream.Append("<DataEmissao>" & capa.dataEmissao.ToString("yyyy-MM-dd") & "</DataEmissao>")
            xmlStream.Append("<Status>1</Status>")
            xmlStream.Append("</Rps>")

            LogInfo("MakeXML processar popularXML.: 2")

            xmlStream.Append("<Competencia>" & capa.dataEmissao.ToString("yyyy-MM-dd") & "</Competencia>")
            xmlStream.Append("<Servico>")
            xmlStream.Append("<Valores>")
            xmlStream.Append("<ValorServicos>" & servico.valrServicos.ToString("#########0.00").Replace(",", ".") & "</ValorServicos>")
            xmlStream.Append("<ValorDeducoes>" & servico.valrDeducoes.ToString("#########0.00").Replace(",", ".") & "</ValorDeducoes>")
            xmlStream.Append("<ValorPis>" & servico.valrPisRetido.ToString("#########0.00").Replace(",", ".") & "</ValorPis>")
            xmlStream.Append("<ValorCofins>" & servico.valrCofinsRetido.ToString("#########0.00").Replace(",", ".") & "</ValorCofins>")
            xmlStream.Append("<ValorInss>" & servico.valrInss.ToString("#########0.00").Replace(",", ".") & "</ValorInss>")
            xmlStream.Append("<ValorIr>" & servico.valrIr.ToString("#########0.00").Replace(",", ".") & "</ValorIr>")
            xmlStream.Append("<ValorCsll>" & servico.valrCsll.ToString("#########0.00").Replace(",", ".") & "</ValorCsll>")
            xmlStream.Append("<OutrasRetencoes>" & servico.outrasRetencoes.ToString("#########0.00").Replace(",", ".") & "</OutrasRetencoes>")
            If servico.issRetido = TipoIss.RETIDO Then
                xmlStream.Append("<ValorIss>" & servico.valrIssRetido.ToString("#########0.00").Replace(",", ".") & "</ValorIss>")
            Else
                xmlStream.Append("<ValorIss>" & servico.valrIss.ToString("#########0.00").Replace(",", ".") & "</ValorIss>")
            End If

            LogInfo("MakeXML processar popularXML.: 3")

            If ambiente = TipoAmbiente.HOMOLOGACAO Then
                xmlStream.Append("<Aliquota>4.00</Aliquota>")
            Else
                xmlStream.Append("<Aliquota>" & servico.aliquota.ToString("#########0.00").Replace(",", ".") & "</Aliquota>")
            End If            
            xmlStream.Append("<DescontoIncondicionado>" & servico.valrDescIncondicionado.ToString("#########0.00").Replace(",", ".") & "</DescontoIncondicionado>")
            xmlStream.Append("<DescontoCondicionado>" & servico.valrDescCondicionado.ToString("#########0.00").Replace(",", ".") & "</DescontoCondicionado>")
            xmlStream.Append("</Valores>")

            If ambiente = TipoAmbiente.HOMOLOGACAO Then
                xmlStream.Append("<IssRetido>2</IssRetido>")
            Else
                xmlStream.Append("<IssRetido>" & servico.issRetido & "</IssRetido>")
            End If

            LogInfo("MakeXML processar popularXML.: 4")

            If servico.issRetido = TipoIss.RETIDO Then
                xmlStream.Append("<ResponsavelRetencao>1</ResponsavelRetencao>")
            End If

            If ambiente = TipoAmbiente.HOMOLOGACAO Then
                xmlStream.Append("<ItemListaServico>1401</ItemListaServico>")
            Else
                If servico.codgServico.Contains(".") Then
                    xmlStream.Append("<ItemListaServico>" & servico.codgServico & "</ItemListaServico>")
                Else
                    xmlStream.Append("<ItemListaServico>" & servico.codgServico.Insert(2, ".") & "</ItemListaServico>")
                End If
            End If

            LogInfo("MakeXML processar popularXML.: 5")

            '
            If ambiente = TipoAmbiente.HOMOLOGACAO Then
                xmlStream.Append("<CodigoCnae>0</CodigoCnae>")
                xmlStream.Append("<CodigoTributacaoMunicipio>14.01</CodigoTributacaoMunicipio>")
            Else
                xmlStream.Append("<CodigoCnae>" & servico.codgCnae & "</CodigoCnae>")
                xmlStream.Append("<CodigoTributacaoMunicipio>" & servico.codgTributacao & "</CodigoTributacaoMunicipio>")
            End If

            LogInfo("MakeXML processar popularXML.: 6")

            xmlStream.Append("<Discriminacao>" & GetDiscriminacaoServico(capa.tipoRps, capa.serieRps, capa.numeroRps, capa.numrCnpj) & "</Discriminacao>")
            xmlStream.Append("<CodigoMunicipio>" & servico.codgMunicipio & "</CodigoMunicipio>")
            'xmlStream.Append("<CodigoPais>1058</CodigoPais>")
            xmlStream.Append("<ExigibilidadeISS>1</ExigibilidadeISS>")
            xmlStream.Append("<MunicipioIncidencia>" & servico.codgMunicipio & "</MunicipioIncidencia>")
            xmlStream.Append("</Servico>")
            xmlStream.Append("<Prestador>")

            LogInfo("MakeXML processar popularXML.: 7")

            xmlStream.Append("<CpfCnpj>")
            If ambiente = TipoAmbiente.PRODUCAO Then
                xmlStream.Append("<Cnpj>" & empresaData.numrCnpj.ToString("00000000000000") & "</Cnpj>")
            Else
                xmlStream.Append("<Cnpj>23403611000348</Cnpj>")
            End If
            xmlStream.Append("</CpfCnpj>")

            If ambiente = TipoAmbiente.PRODUCAO Then
                xmlStream.Append("<InscricaoMunicipal>" & empresaData.inscricaoMunicipal & "</InscricaoMunicipal>")
            Else
                xmlStream.Append("<InscricaoMunicipal>51463</InscricaoMunicipal>")
            End If

            LogInfo("MakeXML processar popularXML.: 8")

            xmlStream.Append("</Prestador>")
            xmlStream.Append("<Tomador>")
            xmlStream.Append("<IdentificacaoTomador>")
            xmlStream.Append("<CpfCnpj>")
            If tomador.tipoDocumento = TipoDocumento.Jurídica Then
                xmlStream.Append("<Cnpj>" & tomador.numrDocumento.ToString("00000000000000") & "</Cnpj>")
            Else
                xmlStream.Append("<Cpf>" & tomador.numrDocumento.ToString("00000000000") & "</Cpf>")
            End If
            xmlStream.Append("</CpfCnpj>")
            xmlStream.Append("</IdentificacaoTomador>")
            xmlStream.Append("<RazaoSocial>" & tomador.razaoSocial & "</RazaoSocial>")
            xmlStream.Append("<Endereco>")

            LogInfo("MakeXML processar popularXML.: 9")

            If tomador.endereco IsNot Nothing AndAlso tomador.endereco.Trim.Length > 0 Then
                xmlStream.Append("<Endereco>" & tomador.endereco & "</Endereco>")
            End If

            If tomador.numero IsNot Nothing AndAlso tomador.numero.Trim.Length > 0 Then
                xmlStream.Append("<Numero>" & tomador.numero & "</Numero>")
            End If

            If tomador.bairro IsNot Nothing AndAlso tomador.bairro.Trim.Length > 0 Then
                xmlStream.Append("<Bairro>" & tomador.bairro & "</Bairro>")
            End If

            LogInfo("MakeXML processar popularXML.: 11")

            xmlStream.Append("<CodigoMunicipio>" & tomador.codgMunicipio & "</CodigoMunicipio>")
            xmlStream.Append("<Uf>" & tomador.uf & "</Uf>")
            'xmlStream.Append("<CodigoPais>1058</CodigoPais>")
            If tomador.cep IsNot Nothing AndAlso tomador.cep.Trim.Length > 0 Then
                xmlStream.Append("<Cep>" & tomador.cep & "</Cep>")
            End If
            xmlStream.Append("</Endereco>")
            xmlStream.Append("<Contato>")
            If tomador.telefone IsNot Nothing AndAlso tomador.telefone.Trim.Length > 0 Then
                xmlStream.Append("<Telefone>" & tomador.ddd & tomador.telefone & "</Telefone>")
            End If

            LogInfo("MakeXML processar popularXML.: 12")

            If tomador.email IsNot Nothing AndAlso tomador.email.Trim.Length > 0 Then
                xmlStream.Append("<Email>" & tomador.email & "</Email>")
            End If
            xmlStream.Append("</Contato>")
            xmlStream.Append("</Tomador>")
            xmlStream.Append("<OptanteSimplesNacional>" & capa.optanteSimplesNacional & "</OptanteSimplesNacional>")
            xmlStream.Append("<IncentivoFiscal>" & capa.incentivadorCultural & "</IncentivoFiscal>")
            xmlStream.Append("</InfDeclaracaoPrestacaoServico>")
            ''ASSINATURA
            xmlStream.Append("</Rps>")
            xmlStream.Append("</GerarNfseEnvio>")

            Return xmlStream.ToString

        End Function

    End Class
End Namespace