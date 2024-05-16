import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout
import pyautogui
import pandas as pd
from time import sleep

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Inserir Vendedor e Rota')
        self.setGeometry(100, 100, 400, 200)

        layout = QVBoxLayout()

        label_vendedor = QLabel('Digite o Vendedor:', self)
        layout.addWidget(label_vendedor)

        self.edit_vendedor = QLineEdit(self)
        layout.addWidget(self.edit_vendedor)

        label_rota = QLabel('Digite a Rota:', self)
        layout.addWidget(label_rota)

        self.edit_rota = QLineEdit(self)
        layout.addWidget(self.edit_rota)

        btn_inserir = QPushButton('Iniciar', self)
        btn_inserir.clicked.connect(self.inserir_proximo_cliente)
        layout.addWidget(btn_inserir)

        self.setLayout(layout)


        # Carregar o Excel
        self.arquivo_excel = 'VENDEDOR_1659.xlsx'
        self.df = pd.read_excel(self.arquivo_excel, dtype={'Cliente': str})
        self.cliente_atual = 0

        # Definindo a velocidade
        pyautogui.PAUSE = 1.5

       # Definindo as coordenadas (inserir e excluir vendedor)
        self.coordenada_1 = (1114,296) # botao localizar 
        self.coordenada_2 = (576,765) # clicar no vendedor
        self.coordenada_3 = (893,722) # Será a mesma para excluir o vendedor
        self.coordenada_7 = (505, 674) # Coordenada do vendedor, será preenchida dinamicamente 
        self.coordenada_9 = (856, 721) # Botão de incluir vendedor

        # Ir para tela de rotas
        self.coordenada_10 = (987, 325)

        # Definindo as coordenadas (inserir e excluir rota)
        self.coordenada_11 = (761, 615)
        self.coordenada_12 = (924, 398) # Botão de excluir
        self.coordenada_13 = (509, 395) # Será a mesma para excluir a rota
        self.coordenada_14 = (407, 442)
        self.coordenada_15 = (532,320) # Coordenada da Rota, será preenchida dinamicamente
        self.coordenada_16 = (282, 346) # Coordenada da Rota
        self.coordenada_17 = (803, 352) # Botão de incluir rota
        self.coordenada_18 = (293, 240)

    def inserir_proximo_cliente(self):
        while self.cliente_atual < len(self.df):
            sleep(5)
            cliente = self.df.loc[self.cliente_atual, 'Cliente']
            self.inserir_cliente(cliente)
            self.cliente_atual += 1

    def inserir_cliente(self, cliente):
        print(f'Inserindo cliente: {cliente}')

        # Passo 1
        pyautogui.click(self.coordenada_1)
        pyautogui.press('backspace')
        pyautogui.write(cliente)
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.sleep(5)

        # Passo 2

        pyautogui.doubleClick(self.coordenada_2)
        pyautogui.click(self.coordenada_3)
        pyautogui.press('enter')

        pyautogui.doubleClick(self.coordenada_2)
        pyautogui.click(self.coordenada_3)
        pyautogui.press('enter')
        '''
        pyautogui.doubleClick(self.coordenada_2)
        pyautogui.click(self.coordenada_3)
        pyautogui.press('enter')

        pyautogui.doubleClick(self.coordenada_2)
        pyautogui.click(self.coordenada_3)
        pyautogui.press('enter')

        pyautogui.doubleClick(self.coordenada_2)
        pyautogui.click(self.coordenada_3)
        pyautogui.press('enter')
        '''
        pyautogui.click(self.coordenada_7)
        pyautogui.write(self.edit_vendedor.text()) # era aqui a 8 
        pyautogui.click(self.coordenada_9)

        # Passo 3
        pyautogui.click(self.coordenada_10)
        pyautogui.doubleClick(self.coordenada_11)
        pyautogui.click(self.coordenada_12)
        pyautogui.press('enter')

        pyautogui.doubleClick(self.coordenada_11)
        pyautogui.click(self.coordenada_12)
        pyautogui.press('enter')
        '''
        pyautogui.doubleClick(self.coordenada_11)
        pyautogui.click(self.coordenada_12)
        pyautogui.press('enter')
        
        pyautogui.doubleClick(self.coordenada_11)
        pyautogui.click(self.coordenada_12)
        pyautogui.press('enter')

        pyautogui.doubleClick(self.coordenada_11)
        pyautogui.click(self.coordenada_12)
        pyautogui.press('enter')
        '''
        pyautogui.click(self.coordenada_13)
        pyautogui.write(self.edit_rota.text())
        pyautogui.press('enter')
        pyautogui.click(self.coordenada_15)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())