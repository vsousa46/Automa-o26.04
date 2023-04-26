from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import (
    presence_of_element_located,
    alert_is_present,
    visibility_of_element_located,
    invisibility_of_element_located,
    frame_to_be_available_and_switch_to_it, 
    staleness_of,
    element_to_be_clickable
)
from selenium.webdriver.common.alert import Alert

class Cadastro:
    def __init__(self, nome, sobrenome, perfil_movel, nome_completo, olos, email, cpf, ):
        self.nome = nome
        self.sobrenome = sobrenome
        self.perfil_movel = perfil_movel
        self.nome_completo = nome_completo
        self.olos = olos
        self.email = email
        self.cpf = cpf

    def zimbra(self):
        navegador = webdriver.Chrome()
            
        #Criação Email Zimbra 
        #Entrar no ADM Zimbra
        navegador.get('https://email.ferreiraechagas.com.br:7071/zimbraAdmin/')
        time.sleep(1)

        #Login
        navegador.find_element(By.ID, 'ZLoginUserName').send_keys('admin')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'ZLoginPassword').send_keys('Gsx750fRf900rGs500e@3')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'ZLoginButton').click()
        time.sleep(10)

        #Adicionar Conta
        navegador.find_element(By.ID, 'ztabv__HOMEV_output_14___container').click()
        time.sleep(2)

        #Criar
        navegador.find_element(By.ID, 'zdlgv__NEW_ACCT_name_2').send_keys(self.perfil_movel)
        time.sleep(0.5)
        navegador.find_element(By.ID, 'zdlgv__NEW_ACCT_givenName').send_keys(self.nome)
        time.sleep(0.5)
        navegador.find_element(By.ID, 'zdlgv__NEW_ACCT_sn').send_keys(self.sobrenome)
        time.sleep(0.5)
        navegador.find_element(By.ID, 'zdlgv__NEW_ACCT_password').send_keys('Mudar@123')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'zdlgv__NEW_ACCT_confirmPassword').send_keys('Mudar@123')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'zdlgv__NEW_ACCT_zimbraPasswordMustChange').click()
        time.sleep(0.5)

        #Concluir
        navegador.find_element(By.ID, 'zdlg__NEW_ACCT_button13_title').click()


        time.sleep(5)

        navegador.quit()

    def cob(self):
        navegador = webdriver.Chrome()
        navegador.get('https://app.cobmais.com.br/cob/')


        navegador.find_element(By.ID, 'txtUsername').send_keys('vitor.varelo')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'txtPasswd').send_keys('Vitor2002$')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'btnSignin').click()
        time.sleep(0.5)

        #Ir para tele de cadastro
        navegador.get('https://app.cobmais.com.br/cob/operador')
        time.sleep(3)

        #Novo Cadastro
        navegador.find_element(By.ID, 'btnNovo').click()
        time.sleep(1)
        navegador.find_element(By.ID, 'txtNome').send_keys(self.nome_completo)
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="selNivel"]/option[16]').click()
        time.sleep(0.5)
        navegador.find_element(By.ID, 'txtCPF').send_keys(self.cpf)
        time.sleep(0.5)
        navegador.find_element(By.ID, 'txtLogin').send_keys(self.perfil_movel)
        time.sleep(0.5)
        navegador.find_element(By.ID, 'txtSenha').send_keys('Mudar123')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'txtSenhaConfirmacao').send_keys('Mudar123')
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="botabs1"]/div/div/div/div[2]/div[2]/label').click()
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="botabs1"]/div/div/div/div[2]/div[1]/label').click()
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="tabOperadores"]/li[2]/a').click()
        time.sleep(0.5)
        navegador.find_element(By.ID, 'txtEmail').send_keys(self.email)
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="tabOperadores"]/li[3]/a').click()
        time.sleep(0.5)
        navegador.find_element(By.ID, 'txtRamal').send_keys(str(self.olos).replace(".0",""))
        time.sleep(0.5)
        navegador.find_element(By.ID, 'txtPin').send_keys('12345678')
        time.sleep(1)

        #Salvar
        navegador.find_element(By.ID, 'btnSalvar').click()
        
    def oolos(self):
        navegador = webdriver.Chrome()
        ############################################## OLOS #####################################################

        #Entrar no Olos

        navegador.get("http://192.168.1.21/Olos/Login.aspx?logout=true")
        wait = WebDriverWait(navegador, 2)
    #verifica se tem alert na poçilga da pagina do zimbra que nao deveria mais existir nessa empresa
        try:
            wait.until(alert_is_present())
            alert2 = Alert(navegador)
            if alert2.text != '':
                observacao = alert2.text
                alert2.accept()
            else:
                alert2.accept()
                print('O alerta não apareceu')
        except:
            print('Não tem essa poçilga aqui')


        time.sleep(5)

        #Logar
        navegador.find_element('xpath', '//*[@id="UserTxt"]').send_keys('vitor_varelo')
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="Password"]').send_keys('12345678')
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="BtnOK"]').click()

        #Ir para a parte de criação de usuario
        navegador.find_element('xpath', '//*[@id="ctl00_TopMenu_SysConfiguration"]').click()
        time.sleep(2)
        navegador.find_element('xpath', '//*[@id="sidebar"]/ul/li/h2[4]/a').click()

        time.sleep(1)
        navegador.find_element(By.ID, 'ctl00_PageMenu_ctl00__labelUserMenu_Users_NewUser').click()

        #Informações da Criação
        navegador.find_element(By.ID, 'ctl00_PageContent_TabContainer1_TabPanel1_UserName').send_keys(self.nome_completo)
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="ctl00_PageContent_TabContainer1_TabPanel1_UserLogin"]').send_keys(str(self.olos).replace(".0",""))
        time.sleep(0.5)
        navegador.find_element(By.ID, 'ctl00_PageContent_TabContainer1_TabPanel1_DropDownPermission').click()
        time.sleep(3)
        navegador.find_element('xpath', '//*[@id="ctl00_PageContent_TabContainer1_TabPanel1_DropDownPermission"]/option[3]').click()
        time.sleep(1.8)
        navegador.find_element('xpath', '//*[@id="ctl00_PageContent_TabContainer1_TabPanel1_DropDownLanguage"]/option[2]').click()
        time.sleep(1.8)
        navegador.find_element(By.ID, 'ctl00_PageContent_TabContainer1_TabPanel1_UserAgent').click()
        time.sleep(3)
        navegador.find_element(By.ID, 'ctl00_PageContent_TabContainer1_TabPanel1_chkPersonalTransfer').click()
        time.sleep(1.8)
        navegador.find_element(By.ID, 'ctl00_PageContent_TabContainer1_TabPanel1_UserActivated').click()
        time.sleep(1.8)
        navegador.find_element(By.ID, '__tab_ctl00_PageContent_TabContainer1_TabPanelLocalization').click()
        time.sleep(1)

        #Localização
        navegador.find_element('xpath', '//*[@id="ctl00_PageContent_TabContainer1_TabPanelLocalization_DropDownListCountry"]/option[2]').click()
        time.sleep(3)
        navegador.find_element('xpath', '//*[@id="ctl00_PageContent_TabContainer1_TabPanelLocalization_DropDownListState"]/option[27]').click()
        time.sleep(3)
        navegador.find_element('xpath', '//*[@id="ctl00_PageContent_TabContainer1_TabPanelLocalization_DropDownListCity"]/option[515]').click()
        time.sleep(1.5)

        #Senha
        navegador.find_element(By.ID, '__tab_ctl00_PageContent_TabContainer1_TabPanel1').click()
        time.sleep(1.8)
        navegador.find_element(By.ID, 'ctl00_PageContent_TabContainer1_TabPanel1_UserPassword').send_keys('12345678')

        #Salvar
        navegador.find_element(By.ID, 'ctl00_PageContent_ButtonCreateUser').click()

        time.sleep(3)      

    def intra(self):



        ########################################### Entrar no Intranet #############################################
        navegador = webdriver.Chrome()
        navegador.get('http://172.16.0.5/default.aspx')
        time.sleep(3)

        #Login 
        navegador.find_element(By.ID, 'nomeUsuarioTextBox').send_keys('vitor.varelo')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'senhaUsuarioTextBox').send_keys('123456')
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="acBtn"]').click()
        time.sleep(0.5)

        #Entrar no controle de funcionarios e cadastro de Colaborador
        navegador.find_element(By.ID, 'configImageButton').click()
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="menu_UpdatePanel121"]/div[1]/button').click()
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="menu_opcaoCadastro"]/a').click()
        time.sleep(0.5)
        navegador.find_element(By.ID, 'menu_cadastroUsuarios').click()
        time.sleep(1)

        #Cadastro
        navegador.find_element(By.ID, 'nomeCompletoTextBox').send_keys(self.nome_completo)
        time.sleep(0.5)
        navegador.find_element('xpath', '//*[@id="departamentoDropDownList"]/option[63]').click()
        time.sleep(0.5)
        navegador.find_element(By.ID, 'emailTextBox').send_keys(self.email)
        time.sleep(0.5)
        navegador.find_element(By.ID, 'usuarioADTextBox').send_keys(self.perfil_movel)
        time.sleep(0.5)
        navegador.find_element(By.ID, 'senhaUsuarioADTextBox').send_keys('Mudar@123')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'loginCobmaisTextBox').send_keys(self.perfil_movel)
        time.sleep(0.5)
        navegador.find_element(By.ID, 'loginOlosTextBox').send_keys(str(self.olos).replace(".0",""))
        time.sleep(0.5)
        navegador.find_element(By.ID, 'loginTextBox').send_keys(self.perfil_movel)
        time.sleep(0.5)
        navegador.find_element(By.ID, 'senhaTextBox').send_keys('123456')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'confirmarSenhaTextBox').send_keys('123456')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'senhaTextBox').send_keys('123456')
        time.sleep(2.5)

        #Salvar
        navegador.find_element(By.ID, 'cadastraUsuarioButton').click()

        time.sleep(1)

    def total(self):
    ####################################### TOTAL ##############################################

        navegador = webdriver.Chrome()
        #entrar no Intranet
        navegador.get('https://intranet.ferreiraechagas.com.br/intranet/login')
        time.sleep(1.5)


        #Logar no Intra
        navegador.find_element(By.ID, 'intranet_user_email').send_keys('vitor.varelo')
        time.sleep(0.5)
        navegador.find_element(By.ID, 'intranet_user_password').send_keys('Vtr021109!')
        time.sleep(0.5)
        navegador.find_element(By.XPATH, '//*[@id="new_intranet_user"]/div[3]/button').click()
        time.sleep(5)

        #Entar no TotalWeb\Pagina de Cadastro
        navegador.find_element(By.XPATH, '/html/body/section[5]/div/a[1]').click()
        time.sleep(4)
        aba_total = navegador.window_handles[1]
        navegador.switch_to.window(aba_total)
        navegador.get('https://totalweb.fcjur.com.br/Administracao/CadUsuario')
        time.sleep(3)


        #Iniciar Cadastro
        navegador.find_element(By.ID, 'btnNovo').click()
        time.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="s2id_Usuario-Tipo"]/a').click()
        time.sleep(1.5)
        navegador.find_element(By.ID, 's2id_autogen2_search').send_keys('Usuario')
        time.sleep(1.5)
        navegador.find_element(By.ID, 's2id_autogen2_search').send_keys(Keys.ENTER)
        time.sleep(1)
        navegador.find_element(By.ID, 'Usuario-NomeCompleto').send_keys(self.nome_completo)
        time.sleep(0.5)
        navegador.find_element(By.ID, 'Usuario-NomeCompleto').send_keys(Keys.ENTER)
        time.sleep(1)
        navegador.find_element(By.ID, 'Usuario-NumAcessoSimultaneo').send_keys('2')
        time.sleep(1)
        navegador.find_element(By.ID, 'Usuario-QtdProcessosDia').send_keys('1')
        time.sleep(1)
        navegador.find_element(By.ID, 'Usuario-Email').send_keys(self.email)
        time.sleep(1)
        navegador.find_element(By.ID, 'Usuario-Login').send_keys(self.perfil_movel)
        time.sleep(0.5)


        #salvar
        navegador.find_element(By.ID, 'btnSalvar').click()