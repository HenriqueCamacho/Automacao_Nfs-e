Imports CClass
Imports System.Xml
Imports System.text.RegularExpressions
Imports NFSeCore

Namespace cWebISS2
    Public Class MakeXML
        Inherits AMakeXML

        Dim dtLoteRps As DataTable = Nothing
        Dim dtListaRps As DataTable = Nothing
        Dim dtRps As DataTable = Nothing
        Dim dtInfRps As DataTable = Nothing
        Dim dtIdentificacaoRps As DataTable = Nothing
        Dim dtInfDeclaracaoPrestacaoServico As DataTable = Nothing
        Dim dtRpsSubstituido As DataTable = Nothing
        Dim dtServico As DataTable = Nothing
        Dim dtValores As DataTable = Nothing
        Dim dtPrestador As DataTable = Nothing
        Dim dtCpfCnpjPrestador As DataTable = Nothing
        Dim dtTomador As DataTable = Nothing
        Dim dtIdentificacaoTomador As DataTable = Nothing
        Dim dtCpfCnpjTomador As DataTable = Nothing
        Dim dtEnderecoTomador As DataTable = Nothing
        Dim dtContato As DataTable = Nothing
        Dim dtIntermediario As DataTable = Nothing
        'Dim dtIntermediarioServicoIdentificacao As DataTable = Nothing
        Dim dtIdentificacaoIntermediario As DataTable = Nothing

        Dim dsLote As DataSet = Nothing
        Dim documentoXML As Xml.XmlDataDocument = Nothing

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
            Try
                Dim schemaFilePath As String = FactoryCore.GetTemplateEnvioLote(servico.codgMunicipio)

                documentoXML = New Xml.XmlDataDocument()
                documentoXML.DataSet.ReadXml(schemaFilePath)
            Catch ex As Exception
                Throw New FalhaException("Erro ao processar arquivo XSD para cidade: " + servico.codgMunicipio.ToString, ex)
            End Try
            ' cria e inicializa os datatables
            initDataTables()

            ' preenche com dados os datatables
            Try
                popularXML(capa, servico, tomador, substituido)
            Catch ex As Exception
                Throw New FalhaException("Erro ao obter dados da integracao: " + ex.Message, ex)
            End Try

            Dim result As String = dsToXML(dsLote)

            result = ordenarXML(result, substituido IsNot Nothing)
            result = result.Replace("Endereco1", "Endereco")
            result = result.Replace("Rps1", "Rps")
            result = result.Replace("CpfCnpj1", "CpfCnpj")
            result = CClass.Util.TString(result).Replace("&", "").Replace("#", "")
            LogInfo("MakeXML processar XML: " + CClass.StringUtil.tirarAcentos(result).Replace("<GerarNfseEnvio xmlns=""http://www.abrasf.org.br/nfse.xsd"">", "<GerarNfseEnvio xmlns=""http://www.abrasf.org.br/nfse.xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.abrasf.org.br/nfse.xsd"">"))
            Return CClass.StringUtil.tirarAcentos(result).Replace("<GerarNfseEnvio xmlns=""http://www.abrasf.org.br/nfse.xsd"">", "<GerarNfseEnvio xmlns=""http://www.abrasf.org.br/nfse.xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.abrasf.org.br/nfse.xsd"">")
        End Function

        ''' <summary>
        ''' Atraves dos dados provenientes da camada de integracao é preenchido o dataset que sera
        ''' convertido em XML
        ''' </summary>
        Private Sub popularXML(ByRef capa As NfseCapaData, ByRef servico As NfseServicoData, ByRef tomador As NfseTomadorData, ByRef substituido As NfseSubstituidoData)
            Dim tpSistema As TipoSistema = CInt(ParametroUtil.GetParametro(ParametroUtil.PARAM_GERAL, "TIPO_SISTEMA"))
            Dim numrLote As Integer = GetNumeroLote()
            Dim empresaData As NfseEmpresasData = GetIntegracaoEmpresa(capa.numrCnpj)
            Dim ambiente As TipoAmbiente = ParametroUtil.GetParametro(empresaData.idEmpresa, "AMBIENTE")
            If ambiente = TipoAmbiente.HOMOLOGACAO Then
                servico.codgMunicipio = 9999999
                empresaData.codgCidade = 9999999
                tomador.codgMunicipio = 9999999
                tomador.email = "henrique.camacho@nbsi.com.br"
            End If

            LogInfo("MakeXML popularXML: 1")
            adicionaValor("Id", dtInfDeclaracaoPrestacaoServico, "Lote" + numrLote.ToString("000000000000000"))
            LogInfo("MakeXML popularXML: 2")
            adicionaValor("Id", dtRps, "NFSe" + capa.numeroRps.ToString)
            adicionaValor("Numero", dtIdentificacaoRps, capa.numeroRps)
            adicionaValor("Serie", dtIdentificacaoRps, capa.serieRps)
            adicionaValor("Tipo", dtIdentificacaoRps, capa.tipoRps)
            If ambiente = TipoAmbiente.HOMOLOGACAO Then
                adicionaValor("DataEmissao", dtRps, "2021-09-16")
            Else
                adicionaValor("DataEmissao", dtRps, capa.dataEmissao.ToString("yyyy-MM-dd"))
            End If
            adicionaValor("Status", dtRps, 1)
            LogInfo("MakeXML popularXML: 3")
            'RpsSubstituido
            If substituido IsNot Nothing AndAlso tpSistema = TipoSistema.DELPHI Then
                adicionaValor("Numero", dtRpsSubstituido, substituido.numeroRpsSubstituido)
                adicionaValor("Serie", dtRpsSubstituido, substituido.serieRpsSubstituido)
                adicionaValor("Tipo", dtRpsSubstituido, substituido.tipoRpsSubstituido)
            Else
                dtRpsSubstituido.Rows(0).Delete()
            End If
            adicionaValor("Competencia", dtInfDeclaracaoPrestacaoServico, capa.dataEmissao.ToString("yyyy-MM-dd"))
            'Servico
            LogInfo("MakeXML popularXML: 4")
            adicionaValor("ValorServicos", dtValores, servico.valrServicos.ToString("#########0.00").Replace(",", "."))
            adicionaValorDbNullOnEmpty("ValorDeducoes", dtValores, servico.valrDeducoes.ToString("#########0.00").Replace(",", "."), True)
            adicionaValorDbNullOnEmpty("ValorPis", dtValores, servico.valrPisRetido.ToString("#########0.00").Replace(",", "."), True)
            adicionaValorDbNullOnEmpty("ValorCofins", dtValores, servico.valrCofinsRetido.ToString("#########0.00").Replace(",", "."), True)
            adicionaValorDbNullOnEmpty("ValorInss", dtValores, servico.valrInss.ToString("#########0.00").Replace(",", "."), True)
            adicionaValorDbNullOnEmpty("ValorIr", dtValores, servico.valrIr.ToString("#########0.00").Replace(",", "."), True)
            adicionaValorDbNullOnEmpty("ValorCsll", dtValores, servico.valrCsll.ToString("#########0.00").Replace(",", "."), True)
            adicionaValorDbNullOnEmpty("OutrasRetencoes", dtValores, servico.outrasRetencoes.ToString("#########0.00").Replace(",", "."), True)
            LogInfo("MakeXML popularXML: 5")
            If servico.issRetido = TipoIss.NAO_RETIDO Then
                adicionaValorDbNullOnEmpty("ValorIss", dtValores, servico.valrIss.ToString("#########0.00").Replace(",", "."), True)
            Else
                adicionaValorDbNullOnEmpty("ValorIss", dtValores, servico.valrIssRetido.ToString("#########0.00").Replace(",", "."), True)
            End If
            LogInfo("MakeXML popularXML: 6")
            If servico.aliquota > 0 Then
                'adicionaValorDbNullOnEmpty("Aliquota", dtValores, CDec((servico.aliquota / 100)).ToString("#########0.0000").Replace(",", "."), True)
                adicionaValorDbNullOnEmpty("Aliquota", dtValores, servico.aliquota.ToString("#########0.00").Replace(",", "."), True)
            Else
                adicionaValorDbNullOnEmpty("Aliquota", dtValores, 0, True)
            End If
            LogInfo("MakeXML popularXML: 7")
            adicionaValorDbNullOnEmpty("DescontoIncondicionado", dtValores, servico.valrDescIncondicionado.ToString("#########0.00").Replace(",", "."), True)
            adicionaValorDbNullOnEmpty("DescontoCondicionado", dtValores, servico.valrDescCondicionado.ToString("#########0.00").Replace(",", "."), True)
            LogInfo("MakeXML popularXML: 8")
            adicionaValor("IssRetido", dtServico, servico.issRetido)

            If servico.issRetido = TipoIss.RETIDO Then 'Segundo prefeitura se o tipoiss for = 2 esta tag nao deverá ser informada.
                adicionaValor("ResponsavelRetencao", dtServico, 1)
            Else
                adicionaValor("ResponsavelRetencao", dtServico, DBNull.Value)
            End If
            adicionaValor("ItemListaServico", dtServico, servico.codgServico.ToString())
            If ambiente = TipoAmbiente.HOMOLOGACAO Then
                adicionaValorDbNullOnEmpty("CodigoCnae", dtServico, 1401, False)
            Else
                adicionaValorDbNullOnEmpty("CodigoCnae", dtServico, Util.retirarNonNumeros(servico.codgCnae), False)
            End If
            LogInfo("MakeXML popularXML: 9")
            adicionaValorDbNullOnEmpty("CodigoTributacaoMunicipio", dtServico, servico.codgTributacao, False)
            If ambiente = TipoAmbiente.HOMOLOGACAO Then
                adicionaValor("Discriminacao", dtServico, CopyString(GetDiscriminacaoServico(capa.tipoRps, capa.serieRps, capa.numeroRps, 2348447000181), 2000))
            Else
                adicionaValor("Discriminacao", dtServico, CopyString(GetDiscriminacaoServico(capa.tipoRps, capa.serieRps, capa.numeroRps, capa.numrCnpj), 2000))
            End If
            adicionaValor("CodigoMunicipio", dtServico, servico.codgMunicipio)
            LogInfo("MakeXML popularXML: 10")
            adicionaValor("ExigibilidadeISS", dtServico, 1)
            adicionaValor("MunicipioIncidencia", dtServico, servico.codgMunicipio)
            LogInfo("MakeXML popularXML: 11")
            'Prestador
            If ambiente = TipoAmbiente.HOMOLOGACAO Then
                adicionaValor("Cnpj", dtCpfCnpjPrestador, "01451027000244")
                adicionaValor("InscricaoMunicipal", dtPrestador, "72919")
            Else
                adicionaValor("Cnpj", dtCpfCnpjPrestador, capa.numrCnpj.ToString("00000000000000"))

                If capa.inscMunicipal IsNot Nothing Then
                    adicionaValor("InscricaoMunicipal", dtPrestador, Regex.Replace(capa.inscMunicipal, "[^0-9a-zA-Z]", ""))
                Else
                    adicionaValor("InscricaoMunicipal", dtLoteRps, Nothing)
                End If
            End If



            LogInfo("MakeXML popularXML: 12")
            'Tomador
            If tomador.tipoDocumento = TipoDocumento.Jurídica Then
                adicionaValor("Cnpj", dtCpfCnpjTomador, tomador.numrDocumento.ToString("00000000000000"))
                adicionaValor("Cpf", dtCpfCnpjTomador, DBNull.Value)
            Else
                adicionaValor("Cpf", dtCpfCnpjTomador, tomador.numrDocumento.ToString("00000000000"))
                adicionaValor("Cnpj", dtCpfCnpjTomador, DBNull.Value)
            End If
            LogInfo("MakeXML popularXML: 13")
            If tomador.inscMunicipal IsNot Nothing Then
                adicionaValor("InscricaoMunicipal", dtIdentificacaoTomador, Regex.Replace(tomador.inscMunicipal, "[^0-9a-zA-Z]", ""))
            Else
                adicionaValor("InscricaoMunicipal", dtIdentificacaoTomador, Nothing)
            End If
            LogInfo("MakeXML popularXML: 14")
            adicionaValorDbNullOnEmpty("RazaoSocial", dtTomador, tomador.razaoSocial, False)
            'Endereco Tomador
            adicionaValorDbNullOnEmpty("Endereco1", dtEnderecoTomador, tomador.endereco, False)
            LogInfo("MakeXML popularXML: 15")
            If tomador.numero IsNot Nothing AndAlso tomador.numero.Trim.Length > 0 Then
                adicionaValorDbNullOnEmpty("Numero", dtEnderecoTomador, tomador.numero, False)
            Else
                adicionaValorDbNullOnEmpty("Numero", dtEnderecoTomador, "0", False)
            End If
            LogInfo("MakeXML popularXML: 16")
            adicionaValorDbNullOnEmpty("Complemento", dtEnderecoTomador, tomador.complemento, False)
            adicionaValorDbNullOnEmpty("Bairro", dtEnderecoTomador, tomador.bairro, False)
            adicionaValorDbNullOnEmpty("CodigoMunicipio", dtEnderecoTomador, tomador.codgMunicipio, False)
            LogInfo("MakeXML popularXML: 17")
            adicionaValorDbNullOnEmpty("Uf", dtEnderecoTomador, tomador.uf, False)
            'adicionaValor("CodigoPais", dtEnderecoTomador, getCodigoPaisBacen(tomador.codNacionalidade, True))
            LogInfo("MakeXML popularXML: 18")
            adicionaValorDbNullOnEmpty("Cep", dtEnderecoTomador, tomador.cep, True) 'Fim endereço Tomador
            'Contato Tomador
            LogInfo("MakeXML popularXML: 19")
            Try
                If tomador.telefone Is Nothing AndAlso tomador.email Is Nothing Then
                    dtContato.Rows(0).Delete()
                ElseIf tomador.email.Trim.Length = 0 AndAlso tomador.telefone.Trim.Length = 0 Then
                    dtContato.Rows(0).Delete()
                Else
                    If tomador.ddd IsNot Nothing AndAlso tomador.telefone IsNot Nothing Then
                        If tomador.ddd.Trim.Length = 0 AndAlso tomador.telefone.Trim.Length = 0 Then
                            adicionaValorDbNullOnEmpty("Telefone", dtContato, Nothing, False)
                        Else
                            Dim _telefone As String = tomador.ddd.ToString + tomador.telefone.ToString
                            adicionaValorDbNullOnEmpty("Telefone", dtContato, _telefone, False)
                        End If
                    Else
                        adicionaValorDbNullOnEmpty("Telefone", dtContato, Nothing, False)
                    End If

                    If tomador.email IsNot Nothing AndAlso tomador.email.Trim.Length > 0 Then
                        adicionaValorDbNullOnEmpty("Email", dtContato, tomador.email.ToLower, False)
                    Else
                        adicionaValorDbNullOnEmpty("Email", dtContato, Nothing, False)
                    End If
                End If
            Catch ex As Exception
                dtContato.Rows(0).Delete()
            End Try
            LogInfo("MakeXML popularXML: 20")
            'Intermediario
            Dim intermediario As NfseIntermediarioServicoData = GetIntegracaoIntermediarioServico(capa.tipoRps, capa.serieRps, capa.numeroRps, capa.numrCnpj)
            If intermediario IsNot Nothing Then
                If intermediario.tipoDocumento = TipoDocumento.Jurídica Then
                    adicionaValor("Cnpj", dtIdentificacaoIntermediario, intermediario.numrDocumento.ToString("00000000000000"))
                    adicionaValor("Cpf", dtIdentificacaoIntermediario, DBNull.Value)
                Else
                    adicionaValor("Cpf", dtIdentificacaoIntermediario, intermediario.numrDocumento.ToString("00000000000"))
                    adicionaValor("Cnpj", dtIdentificacaoIntermediario, DBNull.Value)
                End If
                If intermediario.inscMunicipal IsNot Nothing Then
                    adicionaValor("InscricaoMunicipal", dtIdentificacaoIntermediario, Regex.Replace(intermediario.inscMunicipal, "[^0-9a-zA-Z]", ""))
                Else
                    adicionaValor("InscricaoMunicipal", dtIdentificacaoIntermediario, Nothing)
                End If
                adicionaValor("RazaoSocial", dtIntermediario, intermediario.razaoSocial)
            Else
                dtIntermediario.Rows(0).Delete()
            End If
            LogInfo("MakeXML popularXML: 21")
            If ambiente = TipoAmbiente.HOMOLOGACAO Then
                adicionaValorDbNullOnEmpty("RegimeEspecialTributacao", dtInfDeclaracaoPrestacaoServico, 1, True)
                'adicionaValor("RegimeEspecialTributacao", dtInfDeclaracaoPrestacaoServico, Nothing)
            Else
                adicionaValorDbNullOnEmpty("RegimeEspecialTributacao", dtInfDeclaracaoPrestacaoServico, capa.regimeEspecialTributacao, True)
            End If

            adicionaValor("OptanteSimplesNacional", dtInfDeclaracaoPrestacaoServico, capa.optanteSimplesNacional)
            adicionaValor("IncentivoFiscal", dtInfDeclaracaoPrestacaoServico, 1)
            LogInfo("MakeXML popularXML: 12")
        End Sub

        ''' <summary>
        ''' Metodo auxiliar que referencia os objetos datatable do dataset em copias de uso interno
        ''' </summary>
        Private Sub initDataTables()
            dsLote = documentoXML.DataSet

            dtLoteRps = dsLote.Tables("LoteRps")
            dtListaRps = dsLote.Tables("ListaRps")
            dtRps = dsLote.Tables("Rps1")
            dtInfRps = dsLote.Tables("InfRps")
            dtIdentificacaoRps = dsLote.Tables("IdentificacaoRps")
            dtInfDeclaracaoPrestacaoServico = dsLote.Tables("InfDeclaracaoPrestacaoServico")
            dtRpsSubstituido = dsLote.Tables("RpsSubstituido")
            dtServico = dsLote.Tables("Servico")
            dtValores = dsLote.Tables("Valores")
            dtPrestador = dsLote.Tables("Prestador")
            dtCpfCnpjPrestador = dsLote.Tables("CpfCnpj")
            dtTomador = dsLote.Tables("Tomador")
            dtIdentificacaoTomador = dsLote.Tables("IdentificacaoTomador")
            dtCpfCnpjTomador = dsLote.Tables("CpfCnpj1")
            dtEnderecoTomador = dsLote.Tables("Endereco")
            dtContato = dsLote.Tables("Contato")
            dtIntermediario = dsLote.Tables("Intermediario")
            dtIdentificacaoIntermediario = dsLote.Tables("CpfCnpj1")
        End Sub

        Public Overrides Function ordenarXML(ByVal xml As String, ByVal isSubstituido As Boolean) As String
            Dim xmldoc As XmlDocument = New XmlDocument()
            xmldoc.LoadXml(xml)

            'Dim xPathLoteRps = "/*[local-name()='GerarNfseEnvio']/*[local-name()='Rps']/*[local-name()='InfDeclaracaoPrestacaoServico']/*[local-name()='"
            'moveAfterTag(xmldoc, xPathLoteRps + "Rps1']", xPathLoteRps + "Competencia']")
            'moveAfterTag(xmldoc, xPathLoteRps + "Rps1']/*[local-name()='IdentificacaoRps']", xPathLoteRps + "Rps1']/*[local-name()='DataEmissao']")
            'moveAfterTag(xmldoc, xPathLoteRps + "Servico']/*[local-name()='Valores']", xPathLoteRps + "Servico']/*[local-name()='IssRetido']")
            'moveAfterTag(xmldoc, xPathLoteRps + "Prestador']/*[local-name()='CpfCnpj']", xPathLoteRps + "Prestador']/*[local-name()='InscricaoMunicipal']")
            'moveAfterTag(xmldoc, xPathLoteRps + "Tomador']/*[local-name()='IdentificacaoTomador']", xPathLoteRps + "Tomador']/*[local-name()='RazaoSocial']")
            'moveAfterTag(xmldoc, xPathLoteRps + "Tomador']/*[local-name()='IdentificacaoTomador']/*[local-name()='CpfCnpj1']", xPathLoteRps + "Tomador']/*[local-name()='IdentificacaoTomador']/*[local-name()='InscricaoMunicipal']")
            'moveBeforeTag(xmldoc, xPathLoteRps + "Competencia']", xPathLoteRps + "Rps1']")
            'moveBeforeTag(xmldoc, xPathLoteRps + "OptanteSimplesNacional']", xPathLoteRps + "Tomador']")
            'moveBeforeTag(xmldoc, xPathLoteRps + "IncentivoFiscal']", xPathLoteRps + "OptanteSimplesNacional']")
            Dim xPathLote = GetXpathFromString("GerarNfseEnvio, Rps")
            Dim xPathRps = GetXpathFromString("GerarNfseEnvio, Rps, InfDeclaracaoPrestacaoServico")

            moveAfterTag(xmldoc, xPathLote + "/*[local-name()='CpfCnpj']", xPathLote + "/*[local-name()='NumeroLote']")
            moveBeforeTag(xmldoc, xPathRps + "/*[local-name()='Rps1']", xPathRps + "/*[local-name()='Competencia']")
            moveBeforeTag(xmldoc, xPathRps + "/*[local-name()='Rps1']/*[local-name()='IdentificacaoRps']", xPathRps + "/*[local-name()='Rps1']/*[local-name()='DataEmissao']")
            moveAfterTag(xmldoc, xPathRps + "/*[local-name()='OptanteSimplesNacional']", xPathRps + "/*[local-name()='Tomador']")
            moveAfterTag(xmldoc, xPathRps + "/*[local-name()='RegimeEspecialTributacao']", xPathRps + "/*[local-name()='Tomador']")
            moveAfterTag(xmldoc, xPathRps + "/*[local-name()='IncentivoFiscal']", xPathRps + "/*[local-name()='OptanteSimplesNacional']")
            moveBeforeTag(xmldoc, xPathRps + "/*[local-name()='Servico']/*[local-name()='Valores']", xPathRps + "/*[local-name()='Servico']/*[local-name()='IssRetido']")
            moveBeforeTag(xmldoc, xPathRps + "/*[local-name()='Prestador']/*[local-name()='CpfCnpj']", xPathRps + "/*[local-name()='Prestador']/*[local-name()='InscricaoMunicipal']")
            moveBeforeTag(xmldoc, xPathRps + "/*[local-name()='Tomador']/*[local-name()='IdentificacaoTomador']/*[local-name()='CpfCnpj1']", xPathRps + "/*[local-name()='Tomador']/*[local-name()='IdentificacaoTomador']/*[local-name()='InscricaoMunicipal']")
            moveBeforeTag(xmldoc, xPathRps + "/*[local-name()='Tomador']/*[local-name()='IdentificacaoTomador']", xPathRps + "/*[local-name()='Tomador']/*[local-name()='RazaoSocial']")
            Return xmldoc.OuterXml
        End Function

    End Class
End Namespace
