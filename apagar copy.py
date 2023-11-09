

    # -----------------------------------------
    # função para atualizar exibição do gráfico
    def f_atualizaVisualizacao(self):
        global abasLidas
        
        # executando caso masterplan tenha sido selecionado
        try:
            # cor das barras
            corBarras = '#2f419e'

            # pegando valores da planilha
            nomesDosRecursos = abasLidas['Nomes dos recursos'].tolist()
            inicioPlanejado = abasLidas['Início Planejado'].tolist()
            terminoPlanejado = abasLidas['Término Planejado'].tolist()

            # percorrendo todos os recursos
            diasPlanejado = []
            for indice in range(len(nomesDosRecursos)):
                try:
                    # passando datas de início e fim da vez para datetime
                    inicioDaVez = datetime.strptime(inicioPlanejado[indice].split(' ')[1], '%d/%m/%y')
                    fimDaVez = datetime.strptime(terminoPlanejado[indice].split(' ')[1], '%d/%m/%y')
                    # calculando quantidade de dias e salvando
                    diasPlanejado.append((fimDaVez - inicioDaVez).days + 1)
                except Exception as erro:
                    print('>> atenção! erro na linha ', str(indice), ' do masterplan!\n', erro, '\n\n')

            # pegando recursos únicos
            recursosUnicos = np.unique(nomesDosRecursos)

            # verificando quantidade de dias de todos os recursos
            recursosPessoas = []
            recursosItens = []
            for recursoDaVez in (recursosUnicos):
                nDiasRecursoDaVez = 0
                iniciosRecursoDaVez = []
                finaisRecursoDaVez = []
                for indice in range(len(nomesDosRecursos)):
                    # verificando se é recurso da vez e, caso sim, somando dias desse recurso
                    if nomesDosRecursos[indice] == recursoDaVez:
                        nDiasRecursoDaVez += diasPlanejado[indice]
                        iniciosRecursoDaVez.append(datetime.strptime(inicioPlanejado[indice].split(' ')[1], '%d/%m/%y'))
                        finaisRecursoDaVez.append(datetime.strptime(terminoPlanejado[indice].split(' ')[1], '%d/%m/%y'))

                # depois que passar pelo masterplan, salvando dias de utilização do recurso
                if 'Buscar Recurso_' in recursoDaVez or recursoDaVez == 'Arquivo Técnico':
                    recursosItens.append(
                        {
                            'nomeRecurso': recursoDaVez,
                            'diasRecurso': nDiasRecursoDaVez,
                            'iniciosRecurso': iniciosRecursoDaVez,
                            'finaisRecurso': finaisRecursoDaVez,
                        }
                    )
                else:
                    recursosPessoas.append(
                        {
                            'nomeRecurso': recursoDaVez,
                            'diasRecurso': nDiasRecursoDaVez,
                            'iniciosRecurso': iniciosRecursoDaVez,
                            'finaisRecurso': finaisRecursoDaVez
                        }
                    )

            # pegando opção desejada da lista suspensa
            exibicaoEscolhida = self.listaOpcoes.currentText()

            # exibindo gráfico de acordo com modo escolhido
            if exibicaoEscolhida == 'Recursos: Itens':
                fig = px.bar(recursosItens, x = 'nomeRecurso', y = 'diasRecurso', color_discrete_sequence = [corBarras])
                fig.update_layout(
                    title = '',
                    yaxis_title = 'Quantidade de dias',
                    xaxis_title = '',
                    plot_bgcolor = 'rgba(0, 0, 0, 0)', # cor do fundo do gráfico
                    paper_bgcolor = '#f7f7f7' # cor do 'paper' em que o gráfico é plotado
                )
            elif exibicaoEscolhida == 'Recursos: Pessoas':
                fig = px.bar(recursosPessoas, x = 'nomeRecurso', y = 'diasRecurso', color_discrete_sequence = [corBarras])
                fig.update_layout(
                    title = '',
                    yaxis_title = 'Quantidade de dias',
                    xaxis_title = '',
                    plot_bgcolor = 'rgba(0, 0, 0, 0)', # cor do fundo do gráfico
                    paper_bgcolor = '#f7f7f7' # cor do 'paper' em que o gráfico é plotado
                )
            elif exibicaoEscolhida == 'Período de atividades':
                fig = go.Figure()
                for recurso in recursosPessoas:
                    fig.add_trace(
                        go.Scatter(
                            x = recurso['iniciosRecurso'],
                            y = [recurso['nomeRecurso']] * len(recurso['iniciosRecurso']),
                            mode = 'markers',
                            name = recurso['nomeRecurso']
                        )
                    )
                fig.update_layout(
                    title = '',
                    yaxis_title = '',
                    xaxis_title = '',
                    showlegend = True, #FIXME: ENTENDER PQ NÃO ESTÁ APARECENDO LEGENDA
                    plot_bgcolor = 'rgba(0, 0, 0, 0)', # cor do fundo do gráfico
                    paper_bgcolor = '#f7f7f7' # cor do 'paper' em que o gráfico é plotado
                )

            # atualizando figura
            fig.update_layout(showlegend = False)
            graficoNoCanvas = fig.to_html(full_html = False, include_plotlyjs = 'cdn')
            self.canvasParaGraficos.setHtml(graficoNoCanvas)

        # caso masterplan não tenha sido selecionado antes
        except Exception as e:
            QMessageBox.warning(self, 'Atenção!', 'Nenhum MasterPlan foi selecionado\n')
            None


