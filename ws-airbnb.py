import pandas as pd

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

timeout = 10

# Carregar a planilha
file_path = "inout.xlsx"
sheet_name = "planilha"
df = pd.read_excel(file_path, sheet_name=sheet_name)
df["Localização"] = df["Localização"].astype(str)
df["Avaliações"] = df["Avaliações"].astype(str)
df["Nota"] = df["Nota"].astype(str)
df["Itens inclusos"] = df["Itens inclusos"].astype(str)
df["Valor"] = df["Valor"].astype(str)

df = df.astype('string')

# Filtrar os links que precisam ser processados
links_para_processar = df[df["Localização"] == "nan"]["Link"].tolist()

print(f"Encontrados {len(links_para_processar)} links para processar.")
# print(f"\n {links_para_processar[0]}")

# Criar uma função para atualizar os dados na planilha
def salvar_dados(df, file_path):
  with pd.ExcelWriter(file_path, engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, sheet_name=sheet_name, index=False)
  print("Dados salvos com sucesso!")

for index, link in enumerate(links_para_processar):
  op = webdriver.ChromeOptions()
  op.add_argument('--headless')
  driver = webdriver.Chrome(options=op)
  # driver = webdriver.Chrome()
  driver.get(link)
  time.sleep(5)

  Button = driver.find_elements(By.XPATH, "/html/body/div[9]/div/div/section/div/div/div[2]/div/div[1]/button/span")
  if len(Button) > 0:
    Button[0].click()
  else:
    time.sleep(1)

  Button = WebDriverWait(driver, timeout).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'button[class="l1ovpqvx atm_1he2i46_1k8pnbi_10saat9 atm_yxpdqi_1pv6nv4_10saat9 atm_1a0hdzc_w1h1e8_10saat9 atm_2bu6ew_929bqk_10saat9 atm_12oyo1u_73u7pn_10saat9 atm_fiaz40_1etamxe_10saat9 bmx2gr4 atm_9j_tlke0l atm_9s_1o8liyq atm_gi_idpfg4 atm_mk_h2mmj6 atm_r3_1h6ojuz atm_rd_glywfm atm_70_5j5alw atm_tl_1gw4zv3 atm_9j_13gfvf7_1o5j5ji c1ih3c6 atm_bx_48h72j atm_cs_10d11i2 atm_5j_t09oo2 atm_kd_glywfm atm_uc_1lizyuv atm_r2_1j28jx2 atm_jb_1fkumsa atm_3f_glywfm atm_26_18sdevw atm_7l_1v2u014 atm_8w_1t7jgwy atm_uc_glywfm__1rrf6b5 atm_kd_glywfm_1w3cfyq atm_uc_aaiy6o_1w3cfyq atm_70_1b8lkes_1w3cfyq atm_3f_glywfm_e4a3ld atm_l8_idpfg4_e4a3ld atm_gi_idpfg4_e4a3ld atm_3f_glywfm_1r4qscq atm_kd_glywfm_6y7yyg atm_uc_glywfm_1w3cfyq_1rrf6b5 atm_kd_glywfm_pfnrn2_1oszvuo atm_uc_aaiy6o_pfnrn2_1oszvuo atm_70_1b8lkes_pfnrn2_1oszvuo atm_3f_glywfm_1icshfk_1oszvuo atm_l8_idpfg4_1icshfk_1oszvuo atm_gi_idpfg4_1icshfk_1oszvuo atm_3f_glywfm_b5gff8_1oszvuo atm_kd_glywfm_2by9w9_1oszvuo atm_uc_glywfm_pfnrn2_1o31aam atm_tr_18md41p_csw3t1 atm_k4_kb7nvz_1o5j5ji atm_3f_glywfm_1nos8r_uv4tnr atm_26_wcf0q_1nos8r_uv4tnr atm_7l_1v2u014_1nos8r_uv4tnr atm_3f_glywfm_4fughm_uv4tnr atm_26_4ccpr2_4fughm_uv4tnr atm_7l_1v2u014_4fughm_uv4tnr atm_3f_glywfm_csw3t1 atm_26_wcf0q_csw3t1 atm_7l_1v2u014_csw3t1 atm_7l_1v2u014_pfnrn2 atm_3f_glywfm_1o5j5ji atm_26_4ccpr2_1o5j5ji atm_7l_1v2u014_1o5j5ji f1hzc007 atm_vy_1osqo2v b1sbs18w atm_c8_km0zk7 atm_g3_18khvle atm_fr_1m9t47k atm_l8_182pks8 dir dir-ltr"]'))
  )
  # Button = driver.find_element(By.CSS_SELECTOR, 'button[class="l1ovpqvx atm_1he2i46_1k8pnbi_10saat9 atm_yxpdqi_1pv6nv4_10saat9 atm_1a0hdzc_w1h1e8_10saat9 atm_2bu6ew_929bqk_10saat9 atm_12oyo1u_73u7pn_10saat9 atm_fiaz40_1etamxe_10saat9 bmx2gr4 atm_9j_tlke0l atm_9s_1o8liyq atm_gi_idpfg4 atm_mk_h2mmj6 atm_r3_1h6ojuz atm_rd_glywfm atm_70_5j5alw atm_tl_1gw4zv3 atm_9j_13gfvf7_1o5j5ji c1ih3c6 atm_bx_48h72j atm_cs_10d11i2 atm_5j_t09oo2 atm_kd_glywfm atm_uc_1lizyuv atm_r2_1j28jx2 atm_jb_1fkumsa atm_3f_glywfm atm_26_18sdevw atm_7l_1v2u014 atm_8w_1t7jgwy atm_uc_glywfm__1rrf6b5 atm_kd_glywfm_1w3cfyq atm_uc_aaiy6o_1w3cfyq atm_70_1b8lkes_1w3cfyq atm_3f_glywfm_e4a3ld atm_l8_idpfg4_e4a3ld atm_gi_idpfg4_e4a3ld atm_3f_glywfm_1r4qscq atm_kd_glywfm_6y7yyg atm_uc_glywfm_1w3cfyq_1rrf6b5 atm_kd_glywfm_pfnrn2_1oszvuo atm_uc_aaiy6o_pfnrn2_1oszvuo atm_70_1b8lkes_pfnrn2_1oszvuo atm_3f_glywfm_1icshfk_1oszvuo atm_l8_idpfg4_1icshfk_1oszvuo atm_gi_idpfg4_1icshfk_1oszvuo atm_3f_glywfm_b5gff8_1oszvuo atm_kd_glywfm_2by9w9_1oszvuo atm_uc_glywfm_pfnrn2_1o31aam atm_tr_18md41p_csw3t1 atm_k4_kb7nvz_1o5j5ji atm_3f_glywfm_1nos8r_uv4tnr atm_26_wcf0q_1nos8r_uv4tnr atm_7l_1v2u014_1nos8r_uv4tnr atm_3f_glywfm_4fughm_uv4tnr atm_26_4ccpr2_4fughm_uv4tnr atm_7l_1v2u014_4fughm_uv4tnr atm_3f_glywfm_csw3t1 atm_26_wcf0q_csw3t1 atm_7l_1v2u014_csw3t1 atm_7l_1v2u014_pfnrn2 atm_3f_glywfm_1o5j5ji atm_26_4ccpr2_1o5j5ji atm_7l_1v2u014_1o5j5ji f1hzc007 atm_vy_1osqo2v b1sbs18w atm_c8_km0zk7 atm_g3_18khvle atm_fr_1m9t47k atm_l8_182pks8 dir dir-ltr"]')
  Button.click()

  Nota = WebDriverWait(driver, timeout).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="site-content"]/div/div[1]/div[3]/div/div[1]/div/div[2]/div/div/div/a/div/div[6]/div[1]'))
  )
  # Nota = driver.find_element(By.XPATH, '//*[@id="site-content"]/div/div[1]/div[3]/div/div[1]/div/div[2]/div/div/div/a/div/div[6]/div[1]')
  Nota = Nota.text

  
  Valor = WebDriverWait(driver, timeout).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'span[class="_11jcbg2"]'))
  )
  # Valor = driver.find_element(By.CSS_SELECTOR, 'span[class="_11jcbg2"]')
  Valor = Valor.text
  print(Valor)

  # time.sleep(5)

  ActionChains(driver).scroll_by_amount(0, 1400).perform()

  # time.sleep(5)
  avaliacoes = WebDriverWait(driver, timeout).until(
    EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[class="r1bctolv atm_c8_1sjzizj atm_g3_1dgusqm atm_26_lfmit2_13uojos atm_5j_1y44olf_13uojos atm_l8_1s2714j_13uojos dir dir-ltr"]'))
  )
  # avaliacoes = driver.find_elements(By.CSS_SELECTOR, 'div[class="r1bctolv atm_c8_1sjzizj atm_g3_1dgusqm atm_26_lfmit2_13uojos atm_5j_1y44olf_13uojos atm_l8_1s2714j_13uojos dir dir-ltr"]')
  Avaliacoes = []
  for avaliacao in avaliacoes:
    Avaliacoes.append(avaliacao.text)
  Avaliacoes = [i for i in Avaliacoes if i != ""]
  Avaliacoes = ", ".join(f'"{item}"' for item in Avaliacoes)

  # time.sleep(5)  
  #//*[@id="site-content"]/div/div[1]/div[3]/div/div[1]/div/div[6]/div/div[2]/section/div[3]/button
  Button = WebDriverWait(driver, timeout).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'button[class="l1ovpqvx atm_1he2i46_1k8pnbi_10saat9 atm_yxpdqi_1pv6nv4_10saat9 atm_1a0hdzc_w1h1e8_10saat9 atm_2bu6ew_929bqk_10saat9 atm_12oyo1u_73u7pn_10saat9 atm_fiaz40_1etamxe_10saat9 b1sef8f2 atm_9j_tlke0l atm_9s_1o8liyq atm_gi_idpfg4 atm_mk_h2mmj6 atm_r3_1h6ojuz atm_rd_glywfm atm_3f_uuagnh atm_70_5j5alw atm_vy_1wugsn5 atm_tl_1gw4zv3 atm_9j_13gfvf7_1o5j5ji c3dg75g atm_bx_48h72j atm_c8_2x1prs atm_g3_1jbyh58 atm_fr_11a07z3 atm_cs_10d11i2 atm_5j_t09oo2 atm_6h_t94yts atm_66_nqa18y atm_kd_glywfm atm_uc_1lizyuv atm_r2_1j28jx2 atm_jb_1fkumsa atm_4b_1qnzqti atm_26_1qwqy05 atm_7l_jt7fhx atm_l8_1vkzbvs atm_uc_glywfm__1rrf6b5 atm_kd_glywfm_1w3cfyq atm_uc_aaiy6o_1w3cfyq atm_3f_glywfm_e4a3ld atm_l8_idpfg4_e4a3ld atm_gi_idpfg4_e4a3ld atm_3f_glywfm_1r4qscq atm_kd_glywfm_6y7yyg atm_uc_glywfm_1w3cfyq_1rrf6b5 atm_kd_glywfm_pfnrn2_1oszvuo atm_uc_aaiy6o_pfnrn2_1oszvuo atm_3f_glywfm_1icshfk_1oszvuo atm_l8_idpfg4_1icshfk_1oszvuo atm_gi_idpfg4_1icshfk_1oszvuo atm_3f_glywfm_b5gff8_1oszvuo atm_kd_glywfm_2by9w9_1oszvuo atm_uc_glywfm_pfnrn2_1o31aam atm_tr_18md41p_csw3t1 atm_k4_kb7nvz_1o5j5ji atm_4b_1qnzqti_1w3cfyq atm_7l_jt7fhx_1w3cfyq atm_70_1e7pbig_1w3cfyq atm_4b_1qnzqti_pfnrn2_1oszvuo atm_7l_jt7fhx_pfnrn2_1oszvuo atm_70_1e7pbig_pfnrn2_1oszvuo atm_4b_lb1gtz_1nos8r_uv4tnr atm_26_zbnr2t_1nos8r_uv4tnr atm_7l_jt7fhx_1nos8r_uv4tnr atm_4b_1k0ymf0_4fughm_uv4tnr atm_26_1qwqy05_4fughm_uv4tnr atm_7l_9vytuy_4fughm_uv4tnr atm_4b_lb1gtz_csw3t1 atm_26_zbnr2t_csw3t1 atm_7l_jt7fhx_csw3t1 atm_4b_1k0ymf0_1o5j5ji atm_26_1qwqy05_1o5j5ji atm_7l_9vytuy_1o5j5ji dir dir-ltr"]'))
  )
  # Button = driver.find_element(By.CSS_SELECTOR, 'button[class="l1ovpqvx atm_1he2i46_1k8pnbi_10saat9 atm_yxpdqi_1pv6nv4_10saat9 atm_1a0hdzc_w1h1e8_10saat9 atm_2bu6ew_929bqk_10saat9 atm_12oyo1u_73u7pn_10saat9 atm_fiaz40_1etamxe_10saat9 b1sef8f2 atm_9j_tlke0l atm_9s_1o8liyq atm_gi_idpfg4 atm_mk_h2mmj6 atm_r3_1h6ojuz atm_rd_glywfm atm_3f_uuagnh atm_70_5j5alw atm_vy_1wugsn5 atm_tl_1gw4zv3 atm_9j_13gfvf7_1o5j5ji c3dg75g atm_bx_48h72j atm_c8_2x1prs atm_g3_1jbyh58 atm_fr_11a07z3 atm_cs_10d11i2 atm_5j_t09oo2 atm_6h_t94yts atm_66_nqa18y atm_kd_glywfm atm_uc_1lizyuv atm_r2_1j28jx2 atm_jb_1fkumsa atm_4b_1qnzqti atm_26_1qwqy05 atm_7l_jt7fhx atm_l8_1vkzbvs atm_uc_glywfm__1rrf6b5 atm_kd_glywfm_1w3cfyq atm_uc_aaiy6o_1w3cfyq atm_3f_glywfm_e4a3ld atm_l8_idpfg4_e4a3ld atm_gi_idpfg4_e4a3ld atm_3f_glywfm_1r4qscq atm_kd_glywfm_6y7yyg atm_uc_glywfm_1w3cfyq_1rrf6b5 atm_kd_glywfm_pfnrn2_1oszvuo atm_uc_aaiy6o_pfnrn2_1oszvuo atm_3f_glywfm_1icshfk_1oszvuo atm_l8_idpfg4_1icshfk_1oszvuo atm_gi_idpfg4_1icshfk_1oszvuo atm_3f_glywfm_b5gff8_1oszvuo atm_kd_glywfm_2by9w9_1oszvuo atm_uc_glywfm_pfnrn2_1o31aam atm_tr_18md41p_csw3t1 atm_k4_kb7nvz_1o5j5ji atm_4b_1qnzqti_1w3cfyq atm_7l_jt7fhx_1w3cfyq atm_70_1e7pbig_1w3cfyq atm_4b_1qnzqti_pfnrn2_1oszvuo atm_7l_jt7fhx_pfnrn2_1oszvuo atm_70_1e7pbig_pfnrn2_1oszvuo atm_4b_lb1gtz_1nos8r_uv4tnr atm_26_zbnr2t_1nos8r_uv4tnr atm_7l_jt7fhx_1nos8r_uv4tnr atm_4b_1k0ymf0_4fughm_uv4tnr atm_26_1qwqy05_4fughm_uv4tnr atm_7l_9vytuy_4fughm_uv4tnr atm_4b_lb1gtz_csw3t1 atm_26_zbnr2t_csw3t1 atm_7l_jt7fhx_csw3t1 atm_4b_1k0ymf0_1o5j5ji atm_26_1qwqy05_1o5j5ji atm_7l_9vytuy_1o5j5ji dir dir-ltr"]')
  Button.click()
                                                              
  # time.sleep(5)

  itens = WebDriverWait(driver, timeout).until(
    EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[class="twad414 atm_7l_jt7fhx atm_9j_1kw7nm4 atm_bx_48h72j atm_c8_2x1prs atm_g3_1jbyh58 atm_fr_11a07z3 dir dir-ltr"]'))
  )
  # itens = driver.find_elements(By.CSS_SELECTOR, 'div[class="twad414 atm_7l_jt7fhx atm_9j_1kw7nm4 atm_bx_48h72j atm_c8_2x1prs atm_g3_1jbyh58 atm_fr_11a07z3 dir dir-ltr"]')
  itens_inclusos = []
  for item in itens:
    itens_inclusos.append(item.text)
  itens_inclusos = [i for i in itens_inclusos if i != ""]
  itens_inclusos[1:]
  itens_inclusos = ", ".join(str(item) for item in itens_inclusos)

  # time.sleep(5)
  
  Button = WebDriverWait(driver, timeout).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/div[9]/div/div/section/div/div/div[2]/div/div[1]/button'))
  )
  # Button = driver.find_element(By.XPATH, '/html/body/div[9]/div/div/section/div/div/div[2]/div/div[1]/button')
  Button.click()

  time.sleep(1)

  ActionChains(driver).scroll_by_amount(0, 2500).perform()

  time.sleep(1)

  Localizacao = WebDriverWait(driver, timeout).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'a[title="Abrir esta área no Google Maps (abre uma nova janela)"]'))
  )
  # Localizacao = driver.find_element(By.CSS_SELECTOR, 'a[title="Abrir esta área no Google Maps (abre uma nova janela)"]')
  Localizacao = Localizacao.get_attribute("href")

  # time.sleep(5)

  df.loc[df["Link"] == link, ["Localização", "Avaliações", "Nota", "Itens inclusos", "Valor"]] = [str(Localizacao), str(Avaliacoes), str(Nota), str(itens_inclusos), str(Valor)]

  driver.quit()

# Salvar a planilha atualizada
salvar_dados(df, file_path)

