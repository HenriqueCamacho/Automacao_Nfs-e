Imports CClass
Imports CClass.comum
Imports GrapeCity.ActiveReports.Document
Imports GrapeCity.ActiveReports.Extensibility
Imports System.Xml
Imports System.Text.RegularExpressions
Imports NFSeCore

Namespace cWebISS2
    Public Class ImprimirNFSe
        Inherits AImprimirNFSe

        Public Sub New(ByRef conn As IDbConnection, _
                       ByRef trans As IDbTransaction, _
                       ByRef factory As DBFactory, _
                       ByRef publicVar As PublicVar)
            MyBase.factory = factory
            MyBase.publicVar = publicVar
            MyBase.conn = conn
            MyBase.trans = trans
        End Sub

        Public Overrides Sub imprimir(ByVal codgCidade As Integer, _
                                      ByVal movimentoData As NfseMovimentoData, _
                                      ByVal tpImpressao As TipoImpressao)
            Dim tpSistema As TipoSistema = CInt(ParametroUtil.GetParametro(ParametroUtil.PARAM_GERAL, "TIPO_SISTEMA"))

            If tpSistema = TipoSistema.DELPHI Then
                imprimirDelphi(codgCidade, movimentoData, tpImpressao)
            Else
                imprimirWeb(codgCidade, movimentoData, tpImpressao)
            End If
        End Sub

        Protected Overrides Sub imprimirRps(ByVal codgCidade As Integer, ByVal movimentoData As NfseMovimentoData)
            Dim document As SectionDocument = GetDocumentRps(codgCidade, movimentoData)

            document.Printer.PrinterName = movimentoData.nomeImpressora
            GrapeCity.ActiveReports.PrintExtension.Print(document, False, False)
        End Sub

        Private Function GetDocumentRps(ByVal codgCidade As Integer, ByVal movimentoData As NfseMovimentoData) As SectionDocument
            ' obter uma instancia do datatable
            Dim dtRps As DataTable = ImpressaoUtil.criarTblImpressaoRPS()
            Dim row As DataRow = dtRps.NewRow()
            ' gerar um objeto xml com os dados do rps
            Dim rps As New XmlDocument()
            Dim prestador As XmlDocument = Nothing
            Dim tomador As XmlDocument = Nothing
            Dim servico As XmlDocument = Nothing
            Dim servicoData As NfseServicoData = Nothing
            Dim tomadorData As NfseTomadorData = Nothing
            Dim empresaData As NfseEmpresasData = Nothing
            Dim document As SectionDocument = Nothing
            Dim report As GrapeCity.ActiveReports.SectionReport = Nothing

            Try
                rps.LoadXml(movimentoData.xmlNota)
                prestador = XmlUtil.getDocByTag(rps, "Prestador")
                tomador = XmlUtil.getDocByTag(rps, "Tomador")
                servico = XmlUtil.getDocByTag(rps, "Servico")
                servicoData = GetIntegracaoServico(movimentoData.tipoRps, movimentoData.serieRps, movimentoData.numeroRps, movimentoData.numrCnpj)
                tomadorData = GetIntegracaoTomador(movimentoData.tipoRps, movimentoData.serieRps, movimentoData.numeroRps, movimentoData.numrCnpj)
                empresaData = GetIntegracaoEmpresa(CInt(movimentoData.idEmpresa))

                ' preencher o datatable com os dados necessarios
                row("NUMR_RPS") = servicoData.numeroRps
                'row("DATA_EMISSAO") = GetXmlDate(XmlUtil.getValorTag(rps, "DataEmissao")).ToString("dd/MM/yyyy HH:mm:ss")
                row("DATA_EMISSAO") = movimentoData.dataRecebimento
                row("SERIE_RPS") = servicoData.numeroRps

                If XmlUtil.getValorTag(prestador, "Cnpj") <> "" Then
                    row("CNPJ_CPF_PRESTADOR") = CLng(empresaData.numrCnpj).ToString("00\.000\.000/0000-00")
                Else
                    row("CNPJ_CPF_PRESTADOR") = CLng(XmlUtil.getValorTag(prestador, "Cpf")).ToString("000\.000\.000-00")
                End If

                row("INSC_MUNC_PRESTADOR") = XmlUtil.getValorTag(prestador, "InscricaoMunicipal")
                row("RAZAO_SOCIAL_PRESTADOR") = empresaData.nomeEmpresa
                row("ENDERECO_PRESTADOR") = empresaData.endereco
                row("NUM_END_PRESTADOR") = empresaData.numrEnd
                row("BAIRRO_END_PRESTADOR") = empresaData.bairroEnd
                row("CEP_END_PRESTADOR") = empresaData.cepEnd
                row("EMAIL_PRESTADOR") = empresaData.email.ToLower
                row("MUNICIPIO_PRESTADOR") = empresaData.municipio
                row("UF_PRESTADOR") = empresaData.uf
                row("RAZAO_SOCIAL_TOMADOR") = XmlUtil.getValorTag(tomador, "RazaoSocial")

                If XmlUtil.getValorTag(tomador, "Cnpj") <> "" Then
                    row("CNPJ_CPF_TOMADOR") = CLng(XmlUtil.getValorTag(tomador, "Cnpj")).ToString("00\.000\.000/0000-00")
                Else
                    row("CNPJ_CPF_TOMADOR") = CLng(XmlUtil.getValorTag(tomador, "Cpf")).ToString("000\.000\.000-00")
                End If
                row("INSC_EST_TOMADOR") = tomadorData.inscricaoEstadual
                row("INSC_MUNC_TOMADOR") = XmlUtil.getValorTag(tomador, "InscricaoMunicipal")
                row("ENDERECO_TOMADOR") = XmlUtil.getValorTagXPath(rps, APropriedadesXML.GetXpathFromCollection(New String() {"EnviarLoteRpsEnvio", "LoteRps", "ListaRps", "Rps", "InfRps", "Tomador", "Endereco", "Endereco"}))
                row("NUM_END_TOMADOR") = XmlUtil.getValorTag(tomador, "Numero")
                row("BAIRRO_END_TOMADOR") = XmlUtil.getValorTag(tomador, "Bairro")
                row("CEP_END_TOMADOR") = XmlUtil.getValorTag(tomador, "Cep")
                row("MUNICIPIO_TOMADOR") = tomadorData.nomeMunicipio
                row("UF_TOMADOR") = tomadorData.uf
                row("DISCRIMINACAO") = XmlUtil.getValorTag(servico, "Discriminacao").Replace("|", Chr(13) + Chr(10))

                Try
                    row("VALOR_SERVICO") = CDec(XmlUtil.getValorTag(servico, "ValorServicos").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_SERVICO") = XmlUtil.getValorTag(servico, "ValorServicos")
                End Try

                Try
                    row("VALOR_DEDUCOES") = servicoData.valrDeducoes.ToString().Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_DEDUCOES") = XmlUtil.getValorTag(servico, "ValorDeducoes")
                End Try

                Try
                    row("BASE_CALCULO") = servicoData.baseCalculo.ToString().Replace(".", ",")
                Catch ex As Exception
                    row("BASE_CALCULO") = XmlUtil.getValorTag(servico, "BaseCalculo")
                End Try

                Try
                    row("ALIQUOTA") = servicoData.aliquota.ToString().Replace(".", ",")
                Catch ex As Exception
                    row("ALIQUOTA") = XmlUtil.getValorTag(servico, "Aliquota")
                End Try
                If servicoData.issRetido = TipoIss.NAO_RETIDO Then
                    Try
                        row("VALOR_ISS") = servicoData.valrIss.ToString().Replace(".", ",")
                    Catch ex As Exception
                        row("VALOR_ISS") = XmlUtil.getValorTag(servico, "ValorIss")
                    End Try
                End If
                
                Try
                    row("VALR_DESCONTOS") = servicoData.valrDescIncondicionado.ToString().Replace(".", ",")
                Catch ex As Exception
                    row("VALR_DESCONTOS") = XmlUtil.getValorTag(servico, "DescontoIncondicionado")
                End Try

                Try
                    row("valr_liquido_nfse") = servicoData.valrLiquidoNfse.ToString().Replace(".", ",")
                Catch ex As Exception
                    row("valr_liquido_nfse") = XmlUtil.getValorTag(servico, "ValorLiquidoNfse")
                End Try

                Try
                    row("VALOR_PIS") = servicoData.valrPisRetido.ToString().Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_PIS") = XmlUtil.getValorTag(servico, "ValorPis")
                End Try

                Try
                    row("VALOR_COFINS") = servicoData.valrCofinsRetido.ToString().Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_COFINS") = XmlUtil.getValorTag(servico, "ValorCofins")
                End Try

                Try
                    row("VALOR_INSS") = servicoData.valrInss.ToString().Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_INSS") = XmlUtil.getValorTag(servico, "ValorInss")
                End Try

                Try
                    row("VALOR_IRRF") = servicoData.valrIr.ToString().Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_IRRF") = XmlUtil.getValorTag(servico, "ValorIr")
                End Try

                Try
                    row("VALOR_CSLL") = servicoData.valrCsll.ToString().Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_CSLL") = XmlUtil.getValorTag(servico, "ValorCsll")
                End Try

                Try
                    row("VALOR_OUTRAS_RETENCOES") = servicoData.outrasRetencoes.ToString().Replace(".", ",")
                Catch ex As Exception
                    row("VALOR_OUTRAS_RETENCOES") = XmlUtil.getValorTag(servico, "OutrasRetencoes")
                End Try

                If servicoData.issRetido = TipoIss.RETIDO Then
                    Try
                        row("VALOR_ISSQN") = servicoData.valrIssRetido
                    Catch ex As Exception
                        row("VALOR_ISSQN") = XmlUtil.getValorTag(servico, "ValorIss")
                    End Try
                Else
                    row("VALOR_ISSQN") = 0
                End If

                row("COD_DESC_SERVICO") = XmlUtil.getValorTag(servico, "ItemListaServico") + " - " + GetAtividadeDescricao(CInt(XmlUtil.getValorTag(servico, "ItemListaServico")))
                row("VALOR_CREDITO") = DBNull.Value
                row("OUTRAS_OBSERVACAO") = ""
                row("VALOR_TOTAL_NOTA") = servicoData.valrServicos - servicoData.valrDescIncondicionado

                SetDadosRelatorio(row, movimentoData)
                dtRps.Rows.Add(row)
                report = GetReportRps(movimentoData)
                report.DataSource = dtRps
                document = report.Document

                report.Run()

                Return document
            Finally
                rps = Nothing
                prestador = Nothing
                tomador = Nothing
                servico = Nothing
                servicoData = Nothing
                tomadorData = Nothing
                empresaData = Nothing
                report = Nothing
            End Try
        End Function

        Protected Overrides Sub imprimirNfse(ByVal codgCidade As Integer, ByVal movimentoData As NfseMovimentoData)
            Dim document As SectionDocument = GetDocumentNfse(codgCidade, movimentoData)

            document.Printer.PrinterName = movimentoData.nomeImpressora
            GrapeCity.ActiveReports.PrintExtension.Print(document, False, False)
        End Sub

        Public Overrides Function GetDocumentNfse(ByVal codgCidade As Integer, ByVal movimentoData As NfseMovimentoData) As SectionDocument
            'If codgCidade = 1721000 Then 'Imprimir sempre RPS Não remover isso a não ser que seja autorizado pela prefeitura de Palmas-TO
            'Return GetDocumentRps(codgCidade, movimentoData)
            'End If

            ' obter uma instancia do datatable
            Dim dtNfse As DataTable = ImpressaoUtil.criarTblImpressaoNFSe()
            Dim row As DataRow = dtNfse.NewRow()
            ' gerar um objeto xml com os dados do rps
            Dim nfse As New XmlDocument()
            Dim prestador As XmlDocument = Nothing
            Dim tomador As XmlDocument = Nothing
            Dim servico As XmlDocument = Nothing

            Dim endereco As XmlDocument = Nothing
            Dim servicoData As NfseServicoData = Nothing
            Dim tomadorData As NfseTomadorData = Nothing
            Dim empresaData As NfseEmpresasData = Nothing
            Dim document As SectionDocument = Nothing
            Dim report As GrapeCity.ActiveReports.SectionReport = Nothing
            Dim cidadeData As NfseCidadesData = Nothing

            Try
                LogInfo("MakeXML popularXML 1")
                nfse.LoadXml(movimentoData.xmlNota)
                prestador = XmlUtil.getDocByTag(nfse, "Prestador")
                tomador = XmlUtil.getDocByTag(nfse, "Tomador")
                servico = XmlUtil.getDocByTag(nfse, "Servico")
                endereco = XmlUtil.getDocByTag(tomador, "Endereco")

                servicoData = GetIntegracaoServico(movimentoData.tipoRps, movimentoData.serieRps, movimentoData.numeroRps, movimentoData.numrCnpj)
                tomadorData = GetIntegracaoTomador(movimentoData.tipoRps, movimentoData.serieRps, movimentoData.numeroRps, movimentoData.numrCnpj)
                empresaData = GetIntegracaoEmpresa(movimentoData.numrCnpj)

                LogInfo("MakeXML popularXML 2")
                row("NUMR_NFSE") = XmlUtil.getValorTag(nfse, "Numero")
                LogInfo("MakeXML popularXML 3")
                'row("DATA_EMISSAO") = GetXmlDate(XmlUtil.getValorTag(nfse, "DataEmissao")).ToString("dd/MM/yyyy HH:mm:ss")
                row("DATA_EMISSAO") = XmlUtil.getValorTag(nfse, "DataEmissao")
                LogInfo("MakeXML popularXML 4")
                row("CODIGO_VERIFICACAO") = XmlUtil.getValorTag(nfse, "CodigoVerificacao")
                LogInfo("MakeXML popularXML 5")

                If XmlUtil.getValorTag(prestador, "Cnpj") <> "" Then
                    row("CNPJ_CPF_PRESTADOR") = XmlUtil.getValorTag(prestador, "Cnpj")
                Else
                    row("CNPJ_CPF_PRESTADOR") = XmlUtil.getValorTag(prestador, "Cpf")
                End If
                LogInfo("MakeXML popularXML 6")
                row("INSC_MUNC_PRESTADOR") = XmlUtil.getValorTag(prestador, "InscricaoMunicipal")
                LogInfo("MakeXML popularXML 7")
                row("RAZAO_SOCIAL_PRESTADOR") = empresaData.nomeEmpresa
                LogInfo("MakeXML popularXML 8")
                'If XmlUtil.getValorTagXPath(nfse, "/*[local-name()='ConsultarLoteRpsResposta']/*[local-name()='ListaNfse']/*[local-name()='CompNfse']/*[local-name()='Nfse']/*[local-name()='InfNfse']/*[local-name()='PrestadorServico']/*[local-name()='Endereco']/*[local-name()='Endereco']").Trim.Length > 0 Then
                '    row("ENDERECO_PRESTADOR") = XmlUtil.getValorTagXPath(nfse, "/*[local-name()='ConsultarLoteRpsResposta']/*[local-name()='ListaNfse']/*[local-name()='CompNfse']/*[local-name()='Nfse']/*[local-name()='InfNfse']/*[local-name()='PrestadorServico']/*[local-name()='Endereco']/*[local-name()='Endereco']")
                'Else
                '    row("ENDERECO_PRESTADOR") = XmlUtil.getValorTagXPath(nfse, "/*[local-name()='ConsultarNfseRpsResposta']/*[local-name()='CompNfse']/*[local-name()='Nfse']/*[local-name()='InfNfse']/*[local-name()='PrestadorServico']/*[local-name()='Endereco']/*[local-name()='Endereco']")
                'End If
                row("ENDERECO_PRESTADOR") = empresaData.endereco
                LogInfo("MakeXML popularXML 9")
                row("NUM_END_PRESTADOR") = empresaData.numrEnd
                LogInfo("MakeXML popularXML 10")
                row("BAIRRO_END_PRESTADOR") = empresaData.bairroEnd
                LogInfo("MakeXML popularXML 11")
                row("CEP_END_PRESTADOR") = empresaData.cepEnd
                LogInfo("MakeXML popularXML 12")
                row("MUNICIPIO_PRESTADOR") = empresaData.municipio
                LogInfo("MakeXML popularXML 13")
                row("EMAIL_PRESTADOR") = empresaData.email.ToLower
                LogInfo("MakeXML popularXML 14")
                row("UF_PRESTADOR") = empresaData.uf
                LogInfo("MakeXML popularXML 15")
                row("RAZAO_SOCIAL_TOMADOR") = XmlUtil.getValorTag(tomador, "RazaoSocial")
                LogInfo("MakeXML popularXML 16")
                If XmlUtil.getValorTag(tomador, "Cnpj") <> "" Then
                    row("CNPJ_CPF_TOMADOR") = XmlUtil.getValorTag(tomador, "Cnpj")
                Else
                    row("CNPJ_CPF_TOMADOR") = XmlUtil.getValorTag(tomador, "Cpf")
                End If
                LogInfo("MakeXML popularXML 17")
                row("INSC_EST_TOMADOR") = tomadorData.inscricaoEstadual
                row("INSC_MUNC_TOMADOR") = tomadorData.inscMunicipal
                LogInfo("MakeXML popularXML 18")

                'If XmlUtil.getValorTagXPath(nfse, "/*[local-name()='ConsultarLoteRpsResposta']/*[local-name()='ListaNfse']/*[local-name()='CompNfse']/*[local-name()='Nfse']/*[local-name()='InfNfse']/*[local-name()='TomadorServico']/*[local-name()='Endereco']/*[local-name()='Endereco']").Trim.Length > 0 Then
                '    row("ENDERECO_TOMADOR") = XmlUtil.getValorTagXPath(nfse, "/*[local-name()='ConsultarLoteRpsResposta']/*[local-name()='ListaNfse']/*[local-name()='CompNfse']/*[local-name()='Nfse']/*[local-name()='InfNfse']/*[local-name()='TomadorServico']/*[local-name()='Endereco']/*[local-name()='Endereco']")
                'Else
                '    row("ENDERECO_TOMADOR") = XmlUtil.getValorTagXPath(nfse, "/*[local-name()='ConsultarNfseRpsResposta']/*[local-name()='CompNfse']/*[local-name()='Nfse']/*[local-name()='InfNfse']/*[local-name()='TomadorServico']/*[local-name()='Endereco']/*[local-name()='Endereco']")
                'End If
                'If XmlUtil.getValorTag(endereco, "Endereco") Then

                'End If

                Try
                    row("ENDERECO_TOMADOR") = CClass.Util.ajustar(XmlUtil.getValorTag(endereco, "Endereco"), 30, " ", 2)
                Catch ex As Exception
                    row("ENDERECO_TOMADOR") = XmlUtil.getValorTag(endereco, "Endereco")
                End Try

                LogInfo("MakeXML popularXML 19")
                row("NUM_END_TOMADOR") = XmlUtil.getValorTag(tomador, "Numero")
                LogInfo("MakeXML popularXML 20")
                row("BAIRRO_END_TOMADOR") = XmlUtil.getValorTag(tomador, "Bairro")
                LogInfo("MakeXML popularXML 21")
                row("CEP_END_TOMADOR") = XmlUtil.getValorTag(tomador, "Cep")
                LogInfo("MakeXML popularXML 22")
                row("MUNICIPIO_TOMADOR") = tomadorData.nomeMunicipio
                LogInfo("MakeXML popularXML 23")
                row("UF_TOMADOR") = XmlUtil.getValorTag(tomador, "Uf")
                LogInfo("MakeXML popularXML 24")
                row("DISCRIMINACAO") = XmlUtil.getValorTag(servico, "Discriminacao").Replace("|", Chr(13) + Chr(10)).Replace("??", "çã")
                LogInfo("MakeXML popularXML 25")
                Try
                    row("EMAIL_TOMADOR") = tomadorData.email.ToString
                Catch ex As Exception
                    row("EMAIL_TOMADOR") = ""
                End Try
                LogInfo("MakeXML popularXML 26")
                Try
                    row("VALOR_SERVICO") = CDec(XmlUtil.getValorTag(servico, "ValorServicos").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_SERVICO") = XmlUtil.getValorTag(servico, "ValorServicos")
                End Try
                LogInfo("MakeXML popularXML 27")
                Try
                    row("VALOR_DEDUCOES") = CDec(XmlUtil.getValorTag(servico, "ValorDeducoes").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_DEDUCOES") = XmlUtil.getValorTag(servico, "ValorDeducoes")
                End Try
                LogInfo("MakeXML popularXML 28")
                Try
                    row("BASE_CALCULO") = CDec(XmlUtil.getValorTag(nfse, "BaseCalculo").Replace(".", ","))
                Catch ex As Exception
                    row("BASE_CALCULO") = XmlUtil.getValorTag(servico, "BaseCalculo")
                End Try
                LogInfo("MakeXML popularXML 29")
                Try
                    row("ALIQUOTA") = CDec(XmlUtil.getValorTag(nfse, "Aliquota").Replace(".", ","))
                Catch ex As Exception
                    row("ALIQUOTA") = XmlUtil.getValorTag(servico, "Aliquota")
                End Try
                LogInfo("MakeXML popularXML 30")
                If servicoData.issRetido = TipoIss.NAO_RETIDO Then
                    Try
                        row("VALOR_ISS") = servicoData.valrIss
                    Catch ex As Exception
                        row("VALOR_ISS") = servicoData.valrIss
                    End Try
                    LogInfo("MakeXML popularXML 31")
                End If
                Try
                    row("VALOR_CREDITO") = CDec(XmlUtil.getValorTag(nfse, "ValorCredito").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_CREDITO") = XmlUtil.getValorTag(nfse, "ValorCredito")
                End Try
                LogInfo("MakeXML popularXML 32")
                Try
                    row("VALR_DESCONTOS") = CDec(XmlUtil.getValorTag(servico, "DescontoIncondicionado").Replace(".", ","))
                Catch ex As Exception
                    row("VALR_DESCONTOS") = XmlUtil.getValorTag(servico, "DescontoIncondicionado")
                End Try
                LogInfo("MakeXML popularXML 33")
                Try
                    row("valr_liquido_nfse") = CDec(XmlUtil.getValorTag(nfse, "ValorLiquidoNfse").Replace(".", ","))
                Catch ex As Exception
                    row("valr_liquido_nfse") = XmlUtil.getValorTag(servico, "ValorLiquidoNfse")
                End Try
                LogInfo("MakeXML popularXML 34")
                Try
                    row("VALOR_PIS") = CDec(XmlUtil.getValorTag(servico, "ValorPis").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_PIS") = XmlUtil.getValorTag(servico, "ValorPis")
                End Try
                LogInfo("MakeXML popularXML 35")
                Try
                    row("VALOR_COFINS") = CDec(XmlUtil.getValorTag(servico, "ValorCofins").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_COFINS") = XmlUtil.getValorTag(servico, "ValorCofins")
                End Try
                LogInfo("MakeXML popularXML 36")
                Try
                    row("VALOR_INSS") = CDec(XmlUtil.getValorTag(servico, "ValorInss").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_INSS") = XmlUtil.getValorTag(servico, "ValorInss")
                End Try
                LogInfo("MakeXML popularXML 37")
                Try
                    row("VALOR_IRRF") = CDec(XmlUtil.getValorTag(servico, "ValorIr").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_IRRF") = XmlUtil.getValorTag(servico, "ValorIr")
                End Try
                LogInfo("MakeXML popularXML 38")
                Try
                    row("VALOR_CSLL") = CDec(XmlUtil.getValorTag(servico, "ValorCsll").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_CSLL") = XmlUtil.getValorTag(servico, "ValorCsll")
                End Try
                LogInfo("MakeXML popularXML 39")
                Try
                    row("VALOR_OUTRAS_RETENCOES") = CDec(XmlUtil.getValorTag(servico, "OutrasRetencoes").Replace(".", ","))
                Catch ex As Exception
                    row("VALOR_OUTRAS_RETENCOES") = XmlUtil.getValorTag(servico, "OutrasRetencoes")
                End Try
                LogInfo("MakeXML popularXML 40")
                If servicoData.issRetido = TipoIss.RETIDO Then 'é retido
                    row("VALOR_ISSQN") = servicoData.valrIssRetido
                Else
                    row("VALOR_ISSQN") = 0
                End If
                LogInfo("MakeXML popularXML 41")
                row("COD_DESC_SERVICO") = XmlUtil.getValorTag(servico, "ItemListaServico") + " - " + GetAtividadeDescricao(CInt(XmlUtil.getValorTag(servico, "ItemListaServico")))
                LogInfo("MakeXML popularXML 42")
                row("OUTRAS_OBSERVACAO") = XmlUtil.getValorTag(nfse, "OutrasInformacoes")
                LogInfo("MakeXML popularXML 43")
                row("MSG_RPS") = "RPS nº:" + movimentoData.numeroRps.ToString + ",serie:" + movimentoData.serieRps.ToString + ", emitido em " + String.Format("{0:dd/MM/yyyy}", movimentoData.dataRecebimento)
                LogInfo("MakeXML popularXML 44")
                row("VALOR_TOTAL_NOTA") = CDec(XmlUtil.getValorTag(servico, "ValorServicos").Replace(".", ","))

                'dtNfse.Rows.Add(row)
                LogInfo("MakeXML popularXML 45")
                report = New NfseReport()

                Dim destaque As String

                Try
                    destaque = ParametroUtil.GetParametro(CInt(movimentoData.idEmpresa), "DESTAQUE_IMP_NFSE_CLIENTE")
                Catch ex As Exception
                    destaque = "N"
                End Try

                If destaque = "S" Then
                    LogInfo("MakeXML popularXML 46")
                    CType(report, NfseReport).lblTitulo1Cliente.Visible = True
                    CType(report, NfseReport).lblTitulo2Cliente.Visible = True
                    CType(report, NfseReport).lblTitulo3Cliente.Visible = True
                    CType(report, NfseReport).lblDataRecebimentoCliente.Visible = True
                    CType(report, NfseReport).lblIdentificacaoAssinaturaCliente.Visible = True
                    CType(report, NfseReport).lblNfseCliente.Visible = True
                    CType(report, NfseReport).txbNumeroNotaCliente.Visible = True
                    CType(report, NfseReport).shapeDestaqueCliente.Visible = True
                    CType(report, NfseReport).line1DestaqueCliente.Visible = True
                    CType(report, NfseReport).line2DestaqueCliente.Visible = True
                    CType(report, NfseReport).line3DestaqueCliente.Visible = True
                End If
                LogInfo("MakeXML popularXML 47")
                'Imprimir o campo valor liquido
                Dim mostrarValrLiq As String
                LogInfo("MakeXML popularXML 48")
                Try
                    mostrarValrLiq = ParametroUtil.GetParametro(CInt(movimentoData.idEmpresa), "VALR_LIQ_VISIBLE")
                Catch ex As Exception
                    mostrarValrLiq = "S"
                End Try
                LogInfo("MakeXML popularXML 49")
                If mostrarValrLiq = "N" Then
                    CType(report, NfseReport).Label37.Visible = False
                    CType(report, NfseReport).TextBox34.Visible = False
                    row("OUTRAS_OBSERVACAO") = "RETENÇÃO DE IMPOSTOS POR CONTA DO DESTINATÁRIO!"
                Else
                    CType(report, NfseReport).Label37.Visible = True
                    CType(report, NfseReport).TextBox34.Visible = True
                End If
                LogInfo("MakeXML popularXML 50")
                dtNfse.Rows.Add(row)
                LogInfo("MakeXML popularXML 51")
                report.DataSource = dtNfse
                LogInfo("MakeXML popularXML 52")
                document = report.Document
                LogInfo("MakeXML popularXML 53")
                ' nao é obrigatorio o uso de logomarca por parte da empresa... 
                ' mas nao custa nada avisar que ele nao esta usando
                Try
                    CType(report, NfseReport).LogoEmpresa.Image = ImpressaoUtil.CarregarImagem(ParametroUtil.GetParametro(movimentoData.idEmpresa, "LOGOMARCA"))
                Catch ex As Exception
                    MensagemUtil.GravarMsg(conn, trans, factory, publicVar, movimentoData.idMovimento, "AVISO: nao foi possivel usar logo marca na impressão: " + ex.Message)
                End Try
                LogInfo("MakeXML popularXML 54")
                '---Inicio logomarca prefeitura
                Dim logoPrefeitura As String = ""
                LogInfo("MakeXML popularXML 55")
                Try
                    logoPrefeitura = ParametroUtil.GetParametro(movimentoData.idEmpresa, "BRASAO_PREFEITURA")
                Catch ex As Exception
                    logoPrefeitura = ""
                End Try
                LogInfo("MakeXML popularXML 56")
                If logoPrefeitura <> "" Then
                    Try
                        CType(report, NfseReport).logoPrefeitura.Image = ImpressaoUtil.CarregarImagem(logoPrefeitura)
                    Catch ex As Exception
                        MensagemUtil.GravarMsg(conn, trans, factory, publicVar, movimentoData.idMovimento, "AVISO: nao foi possivel usar logo marca da prefeitura na impressão: " + ex.Message)
                    End Try
                End If
                '---Fim logomarca prefeitura

                ' CType(report, NfseReport).lblTituloNfse.Text = "RECIBO PROVISÓRIO DE SERVIÇO - RPS"
                LogInfo("MakeXML popularXML 57")
                'quando o status do movimento for cancelado, aparecerá uma marca d'água de "Cancelado"
                If movimentoData.status = StatusProcessamento.CANCELADA Then
                    Try
                        CType(report, NfseReport).imCancelado.Visible = True
                    Catch ex As Exception
                        MensagemUtil.GravarMsg(conn, trans, factory, publicVar, movimentoData.idMovimento, "AVISO: nao foi possivel acrescentar marca d'agua 'Cancelado': " + ex.Message)
                    End Try
                End If
                LogInfo("MakeXML popularXML 58")
                '---Inicio titulo e subtitulo
                Dim titulo As String = ""
                Dim subtitulo As String = ""
                LogInfo("MakeXML popularXML 59")
                Try
                    titulo = ParametroUtil.GetParametro(movimentoData.idEmpresa, "TITULO_IMPRESSAO")
                Catch ex As Exception
                    titulo = ""
                End Try
                LogInfo("MakeXML popularXML 60")
                Try
                    subtitulo = ParametroUtil.GetParametro(movimentoData.idEmpresa, "SUBTITULO_IMPRESSAO")
                Catch ex As Exception
                    subtitulo = ""
                End Try
                LogInfo("MakeXML popularXML 61")
                CType(report, NfseReport).lblNomePrefeitura.Text = titulo
                LogInfo("MakeXML popularXML 62")
                CType(report, NfseReport).lblSubTitulo.Text = subtitulo
                '---Fim titulo e subtitulo
                LogInfo("MakeXML popularXML63")
                report.Run()

                Return document
            Finally
                nfse = Nothing
                prestador = Nothing
                tomador = Nothing
                servico = Nothing
                servicoData = Nothing
                tomadorData = Nothing
                empresaData = Nothing
                report = Nothing
            End Try
        End Function

        Public Overrides Function GetDocument(ByVal codgCidade As Integer, ByVal movimentoData As NfseMovimentoData, ByVal tpImpressao As TipoImpressao)
            'If codgCidade = 1721000 Then 'Imprimir sempre RPS Não remover isso a não ser que seja autorizado pela prefeitura de Palmas-TO
            'Return GetDocumentRps(codgCidade, movimentoData)
            ' Else
            If tpImpressao = TipoImpressao.NFSE Then
                Return GetDocumentNfse(codgCidade, movimentoData)
            Else
                Return GetDocumentRps(codgCidade, movimentoData)
            End If
            ' End If
        End Function
    End Class
End Namespace
