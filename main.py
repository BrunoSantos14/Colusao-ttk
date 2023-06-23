from classes import Colas, ModeloColusao, CreateToolTip
import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap.toast import ToastNotification
from ttkbootstrap.tableview import Tableview
from ttkbootstrap.validation import add_regex_validation
from datetime import datetime
import os
# Chamando explicitamente as bibliotecas abaixo para não gerar erro no pyinstaller
import jinja2
from pyarrow.vendored.version import Version


hoje = datetime.now()
app2 = None
app3 = None


class EstudoCola(ttk.Frame):
    def __init__(self, master_window):
        super().__init__(master_window, padding=(20, 10))
        self.pack(fill=BOTH, expand=YES)
        self.caminho = '//Pastor/analytics/Dados/DadosEP/Ano'
        self.nome = ttk.StringVar(value="")
        self.rodada_min = ttk.StringVar(value="")
        self.rodada_max = ttk.StringVar(value="")
        self.ano = ttk.StringVar(value="")
        self.id_modulo = ttk.DoubleVar(value='')
        self.clientes_encontrados = []
        self.grupos_encontrados = []
        self.grupos_filtro = ttk.StringVar(value="")
        self.modulo = ttk.StringVar(value="")
        self.mala_direta = pd.DataFrame()
        self.tabela = pd.DataFrame()
        self.tabela_filtrada = pd.DataFrame()
        self.colors = master_window.style.colors
        self.frame_segmentacao = ttk.Frame()
        self.frame_botao = ttk.Frame()
        self.frame_info_rodadas = ttk.Frame()
        self.frame_tabela_resumo = ttk.Frame()
        self.frame_meter = ttk.Frame()
        self.frame_tres_botoes = ttk.Frame()
        

        label = ttk.Label(self, text="Informe o ano e depois o ID do módulo para começar o estudo!")
        label.config(font=('TkDefaultFont', 10, 'bold'))
        label.pack()
        ttk.Separator(self).pack(fill=X)

        self.frame_segmentacao.pack()
        self.create_combobox(frame=self.frame_segmentacao, 
                             values=self.escolher_ano(), 
                             label="Ano: ", 
                             variable=self.ano)
        
        self.frame_botao.pack()
        self.create_buttonbox(frame=self.frame_botao,
                              action=self.label_id_modulo,
                              cancel_action=self.on_cancel,
                              tooltip_message='Clique para encontrar todas as colas do ano selecionado')    

    def ler_parquet(self):
        """Lê o arquivo .parquet de ano correspondente no diretório e filtra com dados calculados por esstatística robusta pu tradicional,
        além de manter somente onde o usuário não inseriu resposta."""
        ano = self.ano.get()
        try:
            self.tabela = pd.read_parquet(self.caminho + f'/ep_{ano}.parquet')

            # Aplicando filtros conhecidos
            self.tabela = self.tabela.loc[
                (~self.tabela['METODO_CALCULO'].isin(['Qualitativo', 'Boxplot'])) & 
                (self.tabela['RESPOSTA'].isin(['nan', 'None']))]
            return self.tabela
        except:
            raise Exception(f"'{ano}' não é um ano válido!")

    def todas_colas(self):
        self.ler_parquet()
        ano = self.ano.get()
        self.ano_usado = ano
        cola = Colas(ano)
        self.tuplas = cola.obter_todas_colas(self.tabela)
        self.modulos_com_cola = [id_mod for id_mod, qtd_exames, qtd_parts, rodada_min, rodada_max, lista_cola in self.tuplas]

    def label_id_modulo(self):
        """Criando combobox de escolha de módulo e botões"""
        self.todas_colas()
        self.excluir_views_frame(frame=self.frame_botao)

        self.create_combobox(frame=self.frame_segmentacao, 
                             values=self.modulos_com_cola, 
                             label="ID Módulo: ", 
                             variable=self.id_modulo)
        
        self.create_buttonbox(frame=self.frame_botao,
                              action=self.create_frame_resum,
                              cancel_action=self.on_cancel,
                              tooltip_message='Gerar Resumo')
    
    def create_frame_resum(self):
        self.frame_info_rodadas.pack(fill=X, padx=10, pady=10)

        if self.id_modulo.get() in self.modulos_com_cola:
            self.exame_com_cola()

        else:
            self.exame_sem_cola()
            
    def exame_sem_cola(self):
        self.excluir_views_frame(self.frame_info_rodadas)
        self.excluir_views_frame(self.frame_meter)
        self.excluir_views_frame(self.frame_tabela_resumo)
        self.excluir_views_frame(self.frame_tres_botoes)

        ttk.Separator(self.frame_info_rodadas).pack(fill=X)
        
        label = ttk.Label(self.frame_info_rodadas, text=f'Sem colas encontradas para o módulo {int(self.id_modulo.get())}')
        label.config(font=('TkDefaultFont', 10, 'bold'), background='red')
        label.pack(ipady=5)

    def exame_com_cola(self):
        """Criando a visualização na tela principal com as informações resumidas e os botões para detalhe 
        e geração do relatório final (ou mensagem para não cola, se for o caso)."""

        # Se o ano mudar, preciso atualizar o parquet
        if self.ano.get() != self.ano_usado:
            self.todas_colas()

        # Unpacking da tupla
        for id_mod, qtd_exames, qtd_parts, rodada_min, rodada_max, lista_cola in self.tuplas:
            if id_mod == self.id_modulo.get():
                self.qtd_exames = qtd_exames
                self.qtd_parts = qtd_parts
                self.rodada_min = rodada_min
                self.rodada_max = rodada_max
                self.lista_cola = lista_cola

        # Encontrando informações das rodadas
        meses = {1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'}
        self.rodada_min = f'{meses[self.rodada_min.month]}/{self.rodada_min.year}'
        self.rodada_max = f'{meses[self.rodada_max.month]}/{self.rodada_max.year}'
        print(self.rodada_min, self.rodada_max)


        # Encontrando informações das rodadas
        self.info_rodadas(self.frame_info_rodadas)

        # Montando tabelas resumo
        cola = Colas(self.ano)
        resum = cola.lista_cola_resum(self.lista_cola)
        analitos = cola.filtrar_analitos(self.lista_cola)

        self.frame_tabela_resumo.pack()
        self.grupos_a_cadastrar = self.montar_tabela(frame=self.frame_tabela_resumo,
                                                     reset=True,
                                                     df=resum,
                                                     position=0)                                
        self.montar_tabela(frame=self.frame_tabela_resumo,
                           reset=False,
                           df=analitos,
                           position=1)
        
        # Criação dos 3 botões finais
        self.frame_tres_botoes.pack()
        self.create_final_buttons(frame=self.frame_tres_botoes,
                                reset=True,
                                text='Detalhes',
                                action=self.tela_detalhes,
                                tooltip_message='Mala Direta e Lista de Cola',
                                position=0)
        self.create_final_buttons(frame=self.frame_tres_botoes,
                                reset=False,
                                text='Relatório',
                                action=self.pagina_relatorio,
                                tooltip_message='Preencha o relatório para terminar o estudo',
                                position=1)
    
    def escolher_ano(self):
        """Lista os arquivos disponíveis para o estudo de cola. São os arquivos do ano em formato .parquet
        disponíveis no diretório."""
        arquivos = os.listdir(self.caminho)
        arquivos = [arquivo[3:7] for arquivo in arquivos]
        return sorted(set(list(map(int, arquivos))), reverse=True)
    
    def excluir_views_frame(self, frame):
        """Excluir todos os views dentro de um frame para atualização."""
        [view.destroy() for view in frame.winfo_children()]

    def info_rodadas(self, frame):
        """Criando parte intermediária com as informações mais relevantes sobre as colas encontradas."""
        self.excluir_views_frame(frame)
        ttk.Separator(frame).pack(fill=X)

        self.qtd_exames_cola = len(list(self.lista_cola['Analito'].unique()))
        self.qtd_parts_cola = len(self.lista_cola['Cliente'].unique())
        self.modulo = self.lista_cola.loc[0, 'MODULO']
        
        # Escrevendo o Módulo como Título da Segunda Parte
        label = ttk.Label(frame,
                          text=f'{self.modulo}')
        label.config(font=('TkDefaultFont', 15, 'bold'))
        label.pack()   
        
        texto = f"""
        👉 Estudo feito com rodadas de {self.rodada_min} até {self.rodada_max}

        👉 Exames analisados: {self.qtd_exames}     
        Exames com Colusão: {self.qtd_exames_cola} 

        👉 Participantes Investigados: {self.qtd_parts}
        Participantes em Colusão: {self.qtd_parts_cola}
        """
        ttk.Label(frame,
                  text=texto,
                  justify='center').pack()

        self.frame_meter.pack()
        self.create_meter(frame=self.frame_meter,
                          reset=True,
                          text='Exames em Cola',
                          value=round(100*self.qtd_exames_cola/self.qtd_exames, 1),
                          position=0)
        self.create_meter(frame=self.frame_meter,
                          reset=False,
                          text='Labs em Cola',
                          value=round(100*self.qtd_parts_cola/self.qtd_parts, 1),
                          position=1)
        
        # Ler e salvar informações em atributos
        self.abrir_mala_direta()
        
    def montar_tabela(self, frame, reset, df, position):
        """Criando as tabelas que vão aparecer no app."""
        if reset:
            self.excluir_views_frame(frame)
        
        col = [{'text':col, 'stretch':True} for col in list(df.columns)]
        row = [tuple(row.values) for index, row in df.iterrows()]
        df = Tableview(frame,
                        coldata=col,
                        rowdata=row,
                        paginated=False,
                        searchable=True,
                        bootstyle=DARK,
                        height=15,
                        autoalign=True,
                        autofit=True,
                        )
        df.grid(row=0, column=position, padx=5, pady=5)
        return df

    def create_meter(self, frame, reset, text, value, position):
        """Criando os meters que serão usados para gerar informações das colas"""
        if reset:
            self.excluir_views_frame(frame)

        meter = ttk.Meter(
            bootstyle='info',
            master=self.frame_meter,
            metersize=150,
            padding=5,
            amounttotal=100,
            amountused=value,
            metertype="semi",
            subtext=text,
            interactive=False,
            textright='%',
        )
        meter.grid(row=0, column=position, padx=5, pady=5)
    
    def create_combobox(self, frame, values, label, variable):
        """Criação dos comboboxs do app."""      
        form_field_container = ttk.Frame(frame)
        form_field_container.pack(fill=X, expand=YES, pady=5)

        form_field_label = ttk.Label(master=form_field_container, text=label, width=15)
        form_field_label.pack(side=LEFT, padx=12)

        form_input = ttk.Combobox(master=form_field_container, textvariable=variable, values=values)
        form_input.pack(side=LEFT, padx=5, fill=X, expand=YES)
        return form_input

    def create_buttonbox(self, frame, action, cancel_action, tooltip_message):       
        """Criação dos botões de Submit e Cancel."""
        button_container = ttk.Frame(frame)
        button_container.pack(fill=X, expand=YES, pady=(15, 10))

        cancel_btn = ttk.Button(
            master=button_container,
            text="Cancel",
            command=cancel_action,
            bootstyle=DANGER,
            width=7,
        )

        cancel_btn.pack(side=RIGHT, padx=5)
        CreateToolTip(cancel_btn, 'Fechar')

        submit_btn = ttk.Button(
            master=button_container,
            text="Submit",
            command=action,
            bootstyle=SUCCESS,
            width=7,
        )

        submit_btn.pack(side=RIGHT, padx=5)
        CreateToolTip(submit_btn, tooltip_message)

    def create_final_buttons(self, frame, reset, text, action, position, tooltip_message):
        """Criação dos botões finais (acesso a página de detalhes e relatório)"""
        if reset:
            self.excluir_views_frame(frame)

        button = ttk.Button(
            master=frame,
            text=text,
            command=action,
            bootstyle=INFO,
            width=15,
        )
        button.grid(row=0, column=position, padx=5, pady=5)
        CreateToolTip(button, tooltip_message)
    
    def filtrar_segunda_pagina(self):
        """Método que filtra os participantes da tela de detalhes conforme o grupos selecionado pelo usuário."""
        mala_filtrada = self.mala_direta.copy()
        cola_filtrada = self.lista_cola.copy()

        grupos = self.grupos_filtro.get()

        if len(grupos) > 0:
            cola_filtrada = cola_filtrada.loc[cola_filtrada['Grupos'] == grupos]
            mala_filtrada = mala_filtrada.loc[mala_filtrada['ID'].isin(list(cola_filtrada['Cliente'].unique()))]
            mala_filtrada.sort_values('ID', inplace=True)
            self.grupos_filtro.set('')

        return mala_filtrada, cola_filtrada

    def tela_detalhes(self):
        """Criando a página em toplevel que permite visualizar mais detalhes das colas encontradas e informações dos participantes (mala direta)."""
        # Fechar o toplevel se já estiver aberto
        global app2
        if app2 and app2.winfo_exists():
            app2.withdraw()
        

        app2 = ttk.Toplevel(title='Detalhes do Estudo')
        app2.geometry('+500+50')

        mala_filtrada, cola_filtrada = self.filtrar_segunda_pagina()

        # Criar filtro para os grupos
        frame_box = ttk.Frame(app2)
        frame_box.pack()
        self.create_combobox(frame=frame_box,
                            values=self.grupos_encontrados,
                            label='Grupos: ',
                            variable=self.grupos_filtro)
    
        frame_botao = ttk.Frame(app2)
        frame_botao.pack()
        self.create_buttonbox(frame=frame_botao,
                            action=self.tela_detalhes,
                            cancel_action=app2.destroy,
                            tooltip_message='Filtrar')
        
        # Criando notebook com tabela da mala direta (pag1) e lista cola completa (pag2) com filtro de grupos para ambas
        notebook = ttk.Notebook(app2)
        notebook.pack(pady=10, expand=True)

        frame_mala = ttk.Frame(notebook)
        self.montar_tabela(frame=frame_mala,
                        reset=True,
                        df=mala_filtrada,
                        position=0)
        
        frame_cola = ttk.Frame(notebook)
        self.montar_tabela(frame=frame_cola,
                        reset=True,
                        df=cola_filtrada,
                        position=0)

        notebook.add(frame_mala, text='Mala Direta')
        notebook.add(frame_cola, text='Lista Cola')

        app2.mainloop()

    def abrir_mala_direta(self):
        """Lendo a mala direta e filtrando com os participantes encontrados no estudo."""
        # Chamando a mala direta
        mala_direta = pd.read_csv('//Pastor/analytics/MalaDireta.csv', sep=';', encoding='latin1')
        mala_direta = mala_direta[['ID', 'Nome fantasia', 'Grupo representação', 'Grupo empresarial', 'País', 'UF', 'Cidade', 'Bairro', 'Ativo Geral']]
        mala_direta.drop_duplicates('ID', keep='first', inplace=True)
        for coluna in ['Nome fantasia', 'Grupo representação', 'Grupo empresarial', 'País', 'UF', 'Cidade', 'Bairro', 'Ativo Geral']:
            mala_direta[coluna].fillna('-', inplace=True)

        self.clientes_encontrados = list(self.lista_cola['Cliente'].unique())
        self.grupos_encontrados = list(self.lista_cola['Grupos'].unique())

        self.mala_direta = mala_direta.loc[mala_direta['ID'].isin(self.clientes_encontrados)]

    def pagina_relatorio(self):
        """Criando a página em toplevel que gera o relatório final em PDF.""" 
        grupos = [i.values[0] for i in self.grupos_a_cadastrar.tablerows]

        # Fechar o toplevel se já estiver aberto
        global app3
        if app3 and app3.winfo_exists():
            app3.withdraw()
        app3 = ttk.Toplevel(title='Gerar Relatório')
        app3.geometry('+500+200')

        # Mensagem inicial
        label = ttk.Label(app3, text="Informe os dados para o relatório final")
        label.config(font=('TkDefaultFont', 10, 'bold'))
        label.pack()

        label2 = ttk.Label(app3, text='OBS.: Preencha os campos de grupos com um grupo em cada linha e com os participantes separados por um "-".')
        # label2.config(font=('TkDefaultFont', 10, 'bold'))
        label2.pack()
        
        ttk.Separator(app3).pack(fill=X, pady=5)

        # Pedindo o nome do usuário
        frame_nome = ttk.Frame(app3)
        frame_nome.pack(fill=X, expand=YES, pady=5)
        label = ttk.Label(master=frame_nome,
                          text='Responsável pelo estudo: ')
        label.pack(side=LEFT, padx=12)
        nome = ttk.Entry(frame_nome,
                         width=40,
                         textvariable=self.nome)
        nome.pack(side=LEFT, padx=5, expand=YES)
        add_regex_validation(nome, r'[a-zA-Z]')

        # Inserindo label de grupos cadastrados
        frame_cadastrados = ttk.Frame(app3)
        frame_cadastrados.pack()

        label = ttk.Label(master=frame_cadastrados,
                          text='Grupos Cadastrados: ')
        label.pack(side=LEFT, padx=12)
        self.cadastrados = ScrolledText(frame_cadastrados,
                                        padding=10,
                                        height=7,
                                        width=60)
        self.cadastrados.insert(END, '\n'.join(grupos))
        self.cadastrados.pack(side=LEFT, padx=5, expand=YES)

        # Inserindo label de grupos a retirar
        frame_retirados = ttk.Frame(app3)
        frame_retirados.pack()
        label = ttk.Label(master=frame_retirados,
                          text='Grupos Retirados: ')
        label.pack(side=LEFT, padx=12)
        self.retirados = ScrolledText(frame_retirados,
                                 padding=10,
                                 height=7,
                                 width=60)
        self.retirados.pack(side=LEFT, padx=5, expand=YES)

        # Criando botão para gerar o relatório final
        frame_botao = ttk.Frame(app3)
        frame_botao.pack(fill=X, expand=YES, pady=5)

        botao_relatorio = ttk.Button(
                          master=frame_botao,
                          text="Gerar Relatório",
                          command=self.gerar_relatorio,
                          bootstyle=SUCCESS,
                          width=20,
        )

        botao_relatorio.pack(side=BOTTOM, padx=5)
        CreateToolTip(botao_relatorio, 'Gerar o relatório final em PDF\ne salvar no diretório analytics')

        app3.mainloop()

    def gerar_relatorio(self):
        path = '//pastor/analytics/Estudos/01-Máscaras/04 Estudo de colusao/Modelo.docx'
        resp = self.nome.get().title()
        id_mod = int(self.id_modulo.get())
        ano = int(self.ano.get())
        nome_arquivo = f'cola_{ano}_{id_mod}'
        modulo = self.modulo
        selecionados = self.cadastrados.get("1.0", ttk.END)
        retirados = self.retirados.get("1.0", ttk.END)
        retirados = '-' if len(retirados)==1 else retirados

        if resp == "":
            self.notificacao_falha(message="Erro ao salvar o relatório. Informe o nome do colaborador")
        else:
            try:
                modelo = ModeloColusao(path, resp, id_mod, modulo, self.qtd_exames, self.qtd_exames_cola,
                                        self.qtd_parts, self.qtd_parts_cola, selecionados, retirados)
                
                path_salvar = f'//pastor/analytics/Estudos/estudo_cola_modulo/{ano}'
                # Verificar se a pasta não existe antes de criar
                if not os.path.exists(path_salvar):
                    os.makedirs(path_salvar)

                modelo.salvar(f'{nome_arquivo}',
                            path_salvar,
                            app='word')
                
                self.notificacao_sucesso()

                # Fechando a tela de relatório
                global app3
                app3.destroy()

            except Exception as e:
                self.notificacao_falha(message="Ops, algo deu errado. Contate a equipe Analytics ou tente novamente.")
                print(f"Unexpected {e=}, {type(e)=}")
                raise
            
    def notificacao_sucesso(self):
        toast = ToastNotification(
        title="Relatório",
        message="Relatório salvo com sucesso",
        bootstyle=SUCCESS,
        duration=3000,
        )
        toast.show_toast()

    def notificacao_falha(self, message):
        toast = ToastNotification(
        title="Relatório",
        message=message,
        bootstyle=DANGER,
        duration=3000,
        )
        toast.show_toast()

    def on_cancel(self):
        """Cancel and close the application."""
        self.quit()




if __name__ == "__main__":
    app = ttk.Window("Estudo de Colusão", "superhero", resizable=(True, True))
    # app.iconbitmap('logo.ico')
    app.geometry('+40+20')
    EstudoCola(app)
    app.mainloop()
