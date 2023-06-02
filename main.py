from classes import Colas, ModeloColusao, CreateToolTip
import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap import Style
from ttkbootstrap.constants import *
# from ttkbootstrap.toast import ToastNotification
from ttkbootstrap.tableview import Tableview
# from ttkbootstrap.validation import add_regex_validation
from datetime import datetime
import os
# from ttkbootstrap import Style


hoje = datetime.now()
app2 = None

class EstudoCola(ttk.Frame):
    def __init__(self, master_window):
        super().__init__(master_window, padding=(20, 10))
        self.pack(fill=BOTH, expand=YES)
        self.caminho = '//Pastor/analytics/Dados/DadosEP/Ano'
        self.name = ttk.StringVar(value="")
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
        self.lista_cola = pd.DataFrame()
        self.colors = master_window.style.colors
        self.frame_segmentacao = ttk.Frame()
        self.frame_botao = ttk.Frame()
        self.frame_info_rodadas = ttk.Frame()
        self.frame_tabela_resumo = ttk.Frame()
        self.frame_meter = ttk.Frame()
        self.frame_tres_botoes = ttk.Frame()
        

        label = ttk.Label(self, text="Informe o ano e o ID do módulo para começar o estudo!")
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
                              tooltip_message='Clique para informar o Módulo')
        

    def label_id_modulo(self):
        self.ler_parquet()
        self.excluir_views_frame(frame=self.frame_botao)

        self.create_combobox(frame=self.frame_segmentacao, 
                             values=self.escolher_modulo(), 
                             label="ID Módulo: ", 
                             variable=self.id_modulo)
        
        self.create_buttonbox(frame=self.frame_botao,
                              action=self.create_frame_resum,
                              cancel_action=self.on_cancel,
                              tooltip_message='Gerar Resumo')
    
    def create_frame_resum(self):
        # Se o ano mudar, preciso atualizar o parquet
        if self.ano.get() != self.ano:
            self.ler_parquet()

        self.tabela_filtrada = self.filtrar_modulo()
        cola = Colas(self.ano)
        self.lista_cola, rodada_min, rodada_max = cola.listar_colas(self.tabela_filtrada)

        # Encontrando informações das rodadas
        self.frame_info_rodadas.pack(fill=X, padx=10, pady=10)
        self.info_rodadas(self.frame_info_rodadas, rodada_min, rodada_max)
        
        # Montando tabelas resumo
        resum = cola.lista_cola_resum(self.lista_cola)
        analitos = cola.filtrar_analitos(self.lista_cola)

        self.frame_tabela_resumo.pack()
        self.montar_tabela(frame=self.frame_tabela_resumo,
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
    
    def escolher_modulo(self):
        try:
            return list(self.tabela['ID_MODULO'].unique())
        except:
            return []
    
    def escolher_ano(self):
        arquivos = os.listdir(self.caminho)
        arquivos = [arquivo[3:7] for arquivo in arquivos]
        return sorted(set(list(map(int, arquivos))), reverse=True)

    def ler_parquet(self):
        ano = self.ano.get()
        try:
            self.tabela = pd.read_parquet(self.caminho + f'/ep_{ano}.parquet')
            # self.tabela = self.tabela.loc[self.tabela['METODO_CALCULO'] != 'Qualitativo']
            return self.tabela
        except:
            raise Exception(f"'{ano}' não é um ano válido!")

    def filtrar_modulo(self):
        id_modulo = self.id_modulo.get()
        return self.tabela.loc[self.tabela['ID_MODULO'] == id_modulo]
    
    def excluir_views_frame(self, frame):
        [view.destroy() for view in frame.winfo_children()]

    def info_rodadas(self, frame, rodada_min, rodada_max):
        self.excluir_views_frame(frame)
        ttk.Separator(frame).pack(fill=X)

        # Encontrando informações gerais
        self.modulo = self.tabela_filtrada['MODULO'].unique()[0]
        qtd_exames = len(list(self.tabela_filtrada['NOME_DET'].unique()))
        qtd_parts = len(list(self.tabela_filtrada['PART'].unique()))
        
        # Encontrando informações das rodadas
        self.rodada_min = rodada_min.to_pydatetime()
        self.rodada_max = rodada_max.to_pydatetime()
        meses = {1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'}
        self.rodada_min = f'{meses[self.rodada_min.month]}/{self.rodada_min.year}'
        self.rodada_max = f'{meses[self.rodada_max.month]}/{self.rodada_max.year}'


        qtd_exames_cola = len(list(self.lista_cola['Analito'].unique()))
        qtd_parts_cola = len(self.lista_cola['Cliente'].unique())
        
        # Escrevendo o Módulo como Título da Segunda Parte
        label = ttk.Label(frame,
                          text=f'{self.modulo}')
        label.config(font=('TkDefaultFont', 15, 'bold'))
        label.pack()   
        
        texto = f"""
        👉 Estudo feito com rodadas de {self.rodada_min} até {self.rodada_max}

        👉 Exames analisados: {qtd_exames}     
        Exames com Colusão: {qtd_exames_cola} 

        👉 Participantes Investigados: {qtd_parts}
        Participantes em Colusão: {qtd_parts_cola}
        """
        ttk.Label(frame,
                  text=texto,
                  justify='center').pack()

        self.frame_meter.pack()
        self.create_meter(frame=self.frame_meter,
                          reset=True,
                          text='Exames em Cola',
                          value=round(100*qtd_exames_cola/qtd_exames, 1),
                          position=0)
        self.create_meter(frame=self.frame_meter,
                          reset=False,
                          text='Labs em Cola',
                          value=round(100*qtd_parts_cola/qtd_parts, 1),
                          position=1)
        
        # Ler e salvar informações em atributos
        self.abrir_mala_direta()
        
    def montar_tabela(self, frame, reset, df, position):
        if reset:
            self.excluir_views_frame(frame)
        
        col = [{'text':col, 'stretch':True} for col in list(df.columns)]
        row = [tuple(row.values) for index, row in df.iterrows()]
        df = Tableview(frame,
                        coldata=col,
                        rowdata=row,
                        searchable=True,
                        bootstyle=DARK,
                        height=15,
                        autoalign=True,
                        autofit=True,
                        )
        df.grid(row=0, column=position, padx=5, pady=5)

    def create_meter(self, frame, reset, text, value, position):
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
        
    def create_form_entry(self, frame, label, variable):
        form_field_container = ttk.Frame(frame)
        form_field_container.pack(fill=X, expand=YES, pady=5)

        form_field_label = ttk.Label(master=form_field_container, text=label, width=15)
        form_field_label.pack(side=LEFT, padx=12)

        form_input = ttk.Entry(master=form_field_container, textvariable=variable)
        form_input.pack(side=LEFT, padx=5, fill=X, expand=YES)
        return form_input
    
    def create_combobox(self, frame, values, label, variable):       
        form_field_container = ttk.Frame(frame)
        form_field_container.pack(fill=X, expand=YES, pady=5)

        form_field_label = ttk.Label(master=form_field_container, text=label, width=15)
        form_field_label.pack(side=LEFT, padx=12)

        form_input = ttk.Combobox(master=form_field_container, textvariable=variable, values=values)
        form_input.pack(side=LEFT, padx=5, fill=X, expand=YES)
        return form_input

    def create_buttonbox(self, frame, action, cancel_action, tooltip_message):       
        button_container = ttk.Frame(frame)
        button_container.pack(fill=X, expand=YES, pady=(15, 10))

        cancel_btn = ttk.Button(
            master=button_container,
            text="Cancel",
            command=cancel_action,
            bootstyle=DANGER,
            width=6,
        )

        cancel_btn.pack(side=RIGHT, padx=5)
        CreateToolTip(cancel_btn, 'Fechar')

        submit_btn = ttk.Button(
            master=button_container,
            text="Submit",
            command=action,
            bootstyle=SUCCESS,
            width=6,
        )

        submit_btn.pack(side=RIGHT, padx=5)
        CreateToolTip(submit_btn, tooltip_message)

    def create_final_buttons(self, frame, reset, text, action, position, tooltip_message):
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
        # Chamando a mala direta
        mala_direta = pd.read_csv('//Pastor/analytics/MalaDireta.csv', sep=';', encoding='latin1')
        mala_direta = mala_direta[['ID', 'Nome fantasia', 'Grupo representação', 'Grupo empresarial', 'País', 'UF', 'Cidade', 'Bairro', 'Ativo Geral']]
        mala_direta.drop_duplicates('ID', keep='first', inplace=True)
        for coluna in ['Nome fantasia', 'Grupo representação', 'Grupo empresarial', 'País', 'UF', 'Cidade', 'Bairro', 'Ativo Geral']:
            mala_direta[coluna].fillna('-', inplace=True)

        self.clientes_encontrados = list(self.lista_cola['Cliente'].unique())
        self.grupos_encontrados = list(self.lista_cola['Grupos'].unique())

        self.mala_direta = mala_direta.loc[mala_direta['ID'].isin(self.clientes_encontrados)]


    # def lista_cola_completa(self):
    #     pass

    def pagina_relatorio(self):
        pass


    def on_cancel(self):
        """Cancel and close the application."""
        self.quit()






if __name__ == "__main__":

    app = ttk.Window("Estudo de Colusão", "superhero", resizable=(True, True))
    app.geometry('+40+20')
    EstudoCola(app)
    app.mainloop()

