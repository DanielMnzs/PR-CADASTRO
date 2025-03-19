from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd



caminho = r'//depo0903/gpa$/PAC/Daniel Menezes/Python/SELENIUM/PRÉ CADASTRO/CADASTRO.xlsx'
ler = pd.read_excel(caminho)
tabela = pd.DataFrame(ler)
print(tabela)

x = 0
y = 0





while x < len (tabela):
    
    plu = tabela.iloc[x, y]
    pacote = tabela.iloc[x, y + 9]
    pacote2= 1000
    comprimento_caixa = tabela.iloc[x, y + 10]
    pacote3 = "{:.0f}".format(float(pacote)).replace(".", ",")
    largura_caixa = tabela.iloc[x, y +11]
    altura_caixa = tabela.iloc[x, y + 12]
    peso_caixa = tabela.iloc[x, y + 13]
    volume_caixa = tabela.iloc[x, y +14]
    shelf = tabela.iloc[x, 15]
    shelf_str = "{:.0f}".format(float(shelf)).replace(".", ",")
    peso_caixa_str = "{:.2f}".format(float(peso_caixa)).replace(".", ",")
    comprimento_caixa1 = "{:.1f}".format(float(comprimento_caixa)).replace(".", ",")
    largura_caixa1 = "{:.1f}".format(float(largura_caixa)).replace(".", ",")
    altura_caixa1 = "{:.1f}".format(float(altura_caixa)).replace(".", ",")


    volume_caixa_str = "{:.4f}".format(float(volume_caixa)).replace(".", ",")
    temperatura = tabela.iloc[x, y + 21]
    alocacao = tabela.iloc[x,y + 17]
    alocacao_str = str(alocacao)
    base2 = tabela.iloc[x, y + 19]
    altura = tabela.iloc[x, y + 20]
    print("a base é {}".format(base2))
    print("a altura é {}".format(altura))
    print(alocacao_str)
    print(peso_caixa)
    print(temperatura)
    print("shelf aqui {}".format(shelf_str))

    

    altura_unid =  tabela.iloc[x,y + 4]
    comprimento_unid = tabela.iloc[x,y + 5]
    largura_unid =tabela.iloc[x,y + 6]
    peso_unid =tabela.iloc[x ,y + 7]
    volume_unid =tabela.iloc[x, y + 8]
    print (volume_unid)
    print (comprimento_unid)
    peso_unid1 = 0.0001
    volume_unid_str = "{:.4f}".format(float(volume_unid)).replace(".", ",")
    peso_unid_str = "{:.4f}".format(float(peso_unid)).replace(".", ",")
    min1 = shelf*0.66
    min2 = "{:.0f}".format(float(min1)).replace(".", ",")

    altura_unid1 = "{:.1f}".format(float(altura_unid)).replace(".", ",")
    comprimento_unid1 = "{:.1f}".format(float(comprimento_unid)).replace(".", ",")
    largura_unid1 = "{:.1f}".format(float(largura_unid)).replace(".", ",")
    
    
    print ("o minimo é aqui {}".format(min2))



   
    chrome_options = Options()
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])  
    service = Service(r"C:/Program Files/Google/Chrome/Application/chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)

        
    driver.get(r'https://wmswmos1942.cbd.root.gpa/manh/index.html?i=214')

        
    def macro(by, caminho):
        return WebDriverWait(driver, 30).until(EC.element_to_be_clickable((by, caminho)))

    def macro2(by, caminho):
        return WebDriverWait(driver, 30).until(EC.visibility_of_element_located((by, caminho)))

     
    id_input = macro(By.XPATH, '//*[@id="username"]')
    id_input.send_keys("5446902")

    senha_input = macro(By.XPATH, '//*[@id="password"]')
    senha_input.send_keys("@Gpafevereiro2024")
    senha_input.send_keys(Keys.RETURN)  

    time.sleep(3)  

        
    actions = ActionChains(driver)
    def coord(x, y, texto):
        actions.move_by_offset(x, y).click().perform()
        time.sleep(0.5)
        actions.send_keys(texto).perform()

    tres = coord(17, 20, "itens")

       
    itens = macro(By.XPATH, '//*[@id="mps_menusearch-1230-inputEl"]')

    for i in range(3):
        time.sleep(1)
        itens.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.5)

    itens.send_keys(Keys.ENTER)

    time.sleep(5)  


    if alocacao_str == "KGS":
        peso_unid = 1
        comprimento_unid1 = 1
        largura_unid = 1
        altura_unid = 1
        peso_unid = "{:.4f}".format(float(peso_unid1)).replace(".", ",")
        volume_unid_str = 0
        pacote = pacote * 1000
        print ("o pacote agora é {}".format(pacote))

    print(peso_unid_str)

    try:
        iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
        driver.switch_to.frame(iframe)  
        print("Entrou no iframe")
    except:
        print("Nenhum iframe encontrado, seguindo normalmente")

  

     
    try:
        input_itens = macro2(By.XPATH, '//*[contains(@id, "itemLookUpId")]')  
        actions.double_click(input_itens).perform()
        input_itens.send_keys(str(plu))
        print("Texto digitado com sucesso")
    except Exception as e:
        print("Erro ao encontrar o campo de input:", e)

    

    aplicar = macro(By.XPATH, '//*[@id="dataForm:ItemList_lv:ItemList_filterId:ItemList_filterIdapply"]')
    aplicar.click()

    select = macro(By.XPATH, '//*[@id="checkAll_c0_dataForm:ItemList_lv:dataTable"]')
    select.click()

    ver = macro(By.XPATH, '//*[@id="dataForm:ItemList_Viewbutton"]')
    ver.click()

    pacotes = macro(By.XPATH, '//*[@id="Item_Package_Tab_lnk"]')
    pacotes.click()

    editar = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsMain_Edit_button"]')
    editar.click()

    adc = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_addbutton"]')
    adc.click()

    if alocacao_str == "KGS":
        adc.click()


        input_kgs = macro(By.XPATH,'//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_2:ItemPackageListEV_inputMask_PackageQty"]')

        UOM_pacote_KGS = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_2:ItemPackageListEV_PackageUOM_Sel_Menu"]')

        pacote_padrao_kgs = macro(By.XPATH,'//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_2:ItemPackageListEV_StdPackage_Sbc"]')

        UOM_dimensao_KGS = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_2:ItemPackageListEV_DimensionUOM_Sel_Menu"]')

        comp_kgs = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_2:ItemPackageListEV_inputMask_ItemPkgLength"]')

        larg_kgs = macro(By.XPATH,'//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_2:ItemPackageListEV_inputMask_ItemPkgWidth"]')

        alt_kgs = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_2:ItemPackageListEV_inputMask_ItemPkgHeight"]')

        peso_kgs = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_2:ItemPackageListEV_inputMask_Weight"]')

        peso2_kgs = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_2:ItemPackageListEV_WeightUOM_Sel_Menu"]')

        volume_kgs = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_2:ItemPackageListEV_inputMask_ItemPkgVolume"]')

        volume2_kgs = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_2:ItemPackageListEV_VolUOM_Sel_Menu"]')

    if alocacao_str == "KGS":
        pacote3 = "{:.0f}".format(float(peso_caixa)).replace(".", ",")
        input_kgs.send_keys(str(pacote2))
        UOM_pacote_KGS.send_keys("Kg")
        UOM_dimensao_KGS.send_keys("cm")
        pacote_padrao_kgs.click()
        comp_kgs.send_keys(str(comprimento_caixa1))
        larg_kgs.send_keys(str(largura_caixa1))
        alt_kgs.send_keys(str(altura_caixa1))
        peso_kgs.send_keys("1")
        peso2_kgs.send_keys('kg')
        volume_kgs.send_keys("0")
        volume2_kgs.send_keys('m3')
        largura_unid1 = 1
        comprimento_unid1 = 1
        altura_unid1 = 1
        peso_unid_str = "0,001"



    if alocacao_str =='KGS':
        input_pacotes = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_inputMask_PackageQty"]')
        input_pacotes.send_keys(str(pacote3))
        input_pacotes.send_keys('000')
    else:
        input_pacotes = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_inputMask_PackageQty"]')
        input_pacotes.send_keys(str(pacote3))
        


    UOM_pacote = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_PackageUOM_Sel_Menu"]')
    UOM_pacote.send_keys("Caixa")
    time.sleep(0.5)
    UOM_pacote.send_keys(Keys.ENTER)    

    pacote_padrao = macro(By.XPATH,'//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_StdPackage_Sbc"]')
    pacote_padrao.click()

    UOM_dimensao = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_DimensionUOM_Sel_Menu"]')
    for i in range(2):
        UOM_dimensao.send_keys(Keys.ARROW_DOWN)

    comp = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_inputMask_ItemPkgLength"]')
    comp.send_keys(str(comprimento_caixa1))  

    larg = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_inputMask_ItemPkgWidth"]')
    larg.send_keys(str(largura_caixa1))

    alt = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_inputMask_ItemPkgHeight"]')
    alt.send_keys(str(altura_caixa1))

    peso = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_inputMask_Weight"]')
    peso.send_keys(str(peso_caixa_str))

    peso2 = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_WeightUOM_Sel_Menu"]')

    for i in range(3):
        peso2.send_keys(Keys.ARROW_DOWN)

    volume = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_inputMask_ItemPkgVolume"]')
    volume.send_keys(str(volume_caixa_str))

    volume2 = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:newRow_1:ItemPackageListEV_VolUOM_Sel_Menu"]')

    for i in range(6):
        volume2.send_keys(Keys.ARROW_DOWN)

    salvar = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsMain_Save_Detail_button"]')
    salvar.click()

    for i in range(1):
        try:
            erro = macro(By.XPATH, '//*[@id="er_d1_c1"]')
            if erro:
                print("ITEM JÁ CADASTRADO, SEGUINDO PARA O PRÓXIMO PASSO")
                erro.click()
                especificações = macro(By.XPATH, '//*[@id="Item_Detail_Tab_lnk"]')
                especificações.click()
                break
        except:
            print("PACOTES CADASTRADO COM SUCESSO")

    editar2 = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsMain_Edit_button"]')
    editar2.click()

    botao_peso_var = macro(By.XPATH,'//*[@id="dataForm:ItemDetailsEditView_BoolVariableWt"]')
    time.sleep(1)

    if alocacao_str == "KGS":
        botao_peso_var.click()


    tipo_de_prod = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsEditView_SOMProductType"]')
    tipo_de_prod.send_keys("4")

    classe = macro(By.XPATH,'//*[@id="dataForm:ItemDetailsEditView_SOMProductClass"]')
    classe.send_keys("4")
    classe.send_keys(Keys.ENTER)

    nivel_protecao = macro(By.XPATH,'//*[@id="dataForm:ItemDetailsEditView_SOMProtectionLevel"]')
    nivel_protecao.send_keys(temperatura)
    UOM_quantidade = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsEditView_SOMDisplayQtyUOM"]')
    UOM_gramas = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsEditView_SOMBaseUOM"]')

    if alocacao_str =="KGS":
        UOM_quantidade.send_keys("Kg")
        UOM_gramas.send_keys("Gramas")
    else:
        UOM_quantidade.send_keys("Caixa")

    UOM_dimensao_unid = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsEditView_SOMDimensionUOM"]')

    UOM_dimensao_unid.send_keys("cm")

    comp_unid = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsEditView_ITUnitLength"]')
    actions.double_click(comp_unid).perform()
    comp_unid.send_keys(str(comprimento_unid1))

    larg_unid = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsEditView_ITUnitWidth"]')
    actions.double_click(larg_unid).perform()
    larg_unid.send_keys(str(largura_unid1))

    altura_unid_input = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsEditView_ITUnitHeight"]')
    actions.double_click(altura_unid_input).perform()
    altura_unid_input.send_keys(str(altura_unid1))

    peso_unid_input = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsEditView_ITUnitweight"]')
    actions.double_click(peso_unid_input).perform()
    peso_unid_input.send_keys(str(peso_unid_str))

    peso_unid_input2 = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsEditView_SOMWeightUOM"]')

    for i in range(3):
        peso_unid_input2.send_keys("kg")

    volume_unid_input = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsEditView_ITUnitVolume"]')
    actions.double_click(volume_unid_input).perform()
    volume_unid_input.send_keys(volume_unid_str) 

    volume_unid_input2 = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsEditView_SOMVolumeUOM"]')
    volume_unid_input2.send_keys("m3")

    salvar_unid = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsMain_Save_Detail_button"]')
    salvar_unid.click()

    deposito = macro(By.XPATH, '//*[@id="Item_Wmos_Tab_lnk"]')
    deposito.click()

    editar_descricao = macro(By.XPATH,'//*[@id="dataForm:ItemDetailsMain_Edit_button"]')
    editar_descricao.click()

    solicitar_embalagens = macro(By.XPATH, '//*[@id="dataForm:ItemWmosEV_ITITPromptPackQty"]')
    solicitar_embalagens.send_keys('Sempre exigir')

    fabric = macro(By.XPATH, '//*[@id="dataForm:ItemWmosEV_SOMMfgDateRequired"]')
    fabric.send_keys('Sempre exigir')

    venc = macro(By.XPATH, '//*[@id="dataForm:ItemWmosEV_SOMExpirationDateRequired"]')
    venc.send_keys("Sempre exigir")

    base = macro(By.XPATH, '//*[@id="dataForm:ItemWmosEV_SOMDateBasisCode"]')
    base.send_keys('Dt de expiração')

    max = macro(By.XPATH, '//*[@id="dataForm:ItemWmosEV_ITMaxDaysBeforeExp"]')
    actions.double_click(max).perform()
    max.send_keys(str(shelf_str))

    min = macro(By.XPATH,'//*[@id="dataForm:ItemWmosEV_ITMinDaysBeforeExp"]')
    actions.double_click(min).perform()
    time.sleep(1)
    min.send_keys(str(min2))
    time.sleep(1)
    salvar_deposito = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsMain_Save_Detail_button"]')
    salvar_deposito.click( )

    instal_item = macro(By.XPATH, '//*[@id="dataForm:Itemdetails_ItemFacilityList"]')
    instal_item.click()

    caixa_select = macro(By.XPATH, '//*[@id="checkAll_c0_dataForm:ItemFacilityList_lv:dataTable"]')
    caixa_select.click()

    ver_instal = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityList_Viewbutton"]')
    ver_instal.click()

    editar_instal = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityView_Edit_button"]')
    editar_instal.click()

    onda = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityDetails_SOMExcessWaveNeedProcessType"]')
    onda.send_keys("Reabas os locais dinâmicos igualmente com o resto")

    pere = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityDetails_SOMQualityInspectionGroup"]')
    pere.send_keys("Perecível")

    onda2 = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityDetails_SOMDefaultWaveProcessType"]')

    if alocacao_str == "KGS":
        onda2.send_keys('Retirar da reserva, ativo em Unidade')
    else:
        onda2.send_keys('Retirar da reserva, ativo em Caixa')

    alocacao_deposito = macro(By.XPATH,'//*[@id="dataForm:ItemFacilityDetails_SOMAllocationType"]')

    if alocacao_str =="KGS":
        alocacao_deposito.send_keys("KGS")
    else: 
        alocacao_deposito.send_keys("CXS")


    prat = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityDetails_ITShelfDays"]')
    regra_21 = round(shelf * 0.33)
    regra_11 = round(shelf*0.33 + 1)


    #REGRA FEFO
    if shelf >= 180:
        actions.double_click(prat).perform()
        prat.send_keys(30)
    else:
        if shelf < 180 and shelf >= 21:
            actions.double_click(prat).perform()
            prat.send_keys(regra_21)
        else:
            if shelf >= 0 and shelf <=21:
                actions.double_click(prat).perform()
                prat.send_keys(regra_11)

    atribuicao = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityDetails_ITActiveMaxUnits"]')
    actions.double_click(atribuicao).perform()
    atribuicao.send_keys('99999')
    time.sleep(1)

    atrib_max = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityDetails_SOMAssignDynamicCasePick"]')
    atrib_max.send_keys("Atrib c/ base qtd máx itens")
    time.sleep(1)

    atribuicao2 = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityDetails_ITCasePickMaxCases"]')
    actions.double_click(atribuicao2).perform()
    atribuicao2.send_keys('99999')
    time.sleep(1)

    atribuicao3 = macro(By.XPATH,'//*[@id="dataForm:ItemFacilityDetails_ITMaxNbrOfDynamicActiveLocations"]')
    actions.double_click(atribuicao3).perform()
    atribuicao3.send_keys("3")

    time.sleep(1)
    atrib_max2 = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityDetails_SOMAssignDynamicActive"]')
    atrib_max2.send_keys('Atrib c/ base qtd máx itens')

    time.sleep(1)
    atribuicao4 = macro(By.XPATH,'//*[@id="dataForm:ItemFacilityDetails_ITActiveMaxCases"]')
    actions.double_click(atribuicao4).perform()
    atribuicao4.send_keys("99999")
    time.sleep(1)

    atribuicao5 = macro(By.XPATH,'//*[@id="dataForm:ItemFacilityDetails_ITCasePickMaxUnits"]')
    actions.double_click(atribuicao5).perform()
    atribuicao5.send_keys("99999")

    time.sleep(1)

    atribuicao6 = macro(By.XPATH,'//*[@id="dataForm:ItemFacilityDetails_ITMaxNbrOfDynamicCasePickLocnsPerItem"]')
    actions.double_click(atribuicao6).perform()
    atribuicao6.send_keys("3")

    time.sleep(1)

    base_input = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityDetails_ITCasesPerTier"]')
    time.sleep(0.5)
    actions.double_click(base_input).perform()
    base_input.send_keys(str(base2))

    time.sleep(1)

    altura_input = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityDetails_ITTiersPerPallet"]')
    actions.double_click(altura_input).perform()
    altura_input.send_keys(str(altura))

    data = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityDetails_ITDefaultDateFormatSOM"]')
    for i in range(10):
        data.send_keys(Keys.ARROW_DOWN)

    salvar_instal = macro(By.XPATH, '//*[@id="dataForm:ItemFacilityEdit_save_button"]')
    salvar_instal.click()

    time.sleep(1)

    voltar = macro2(By.XPATH,'//*[@id="backImage"]')
    voltar.click()

    


    driver.quit()
            


    time.sleep(2)
    x +=  1
    print(x)
    print("PRODUTO {} CADASTRADO COM SUCESSO! ".format(plu))

