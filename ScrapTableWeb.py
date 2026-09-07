from io import StringIO
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException
)

#L’URL de notre site où se trouve le tableau avec les joueurs et leur salaire
url = "https://www.legalsport.pl/zawodnicy/pilka-nozna/najwiecej-zarabiajacy-pilkarze/ligue-1/"

#Ce xpath je l’ai obtenu en faisant un click-droit sur le bouton "afficher plus" du site puis sur "Inspecter" 
xpath_bouton = (
    "/html/body/div[2]/div[2]/div/div[1]/div[2]/main/"
    "div/div[3]/div[3]/div/div[2]/button"
)

#Ici on prépare un navigateur Chrome automatisé pour Selenium
options = Options()
options.add_argument("--headless=new")
options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)

#On ouvre la page, on clique en boucle sur notre bouton "Afficher plus" tant qu’il est disponible, puis on récupère le code HTML
try:
    driver.get(url)

    nombre_clics = 0

    while True:
        try:
            bouton = wait.until(
                EC.element_to_be_clickable((By.XPATH, xpath_bouton))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});",
                bouton
            )

            driver.execute_script("arguments[0].click();", bouton)
            nombre_clics += 1

            # Attend le chargement des nouvelles lignes
            driver.implicitly_wait(2)

        except (TimeoutException, StaleElementReferenceException):
            break

    html = driver.page_source

finally:
    driver.quit()


# On extrait toutes les tables HTML trouvées dans notre page, puis on affiche pour chacune son index et son nombre de lignes ce qui nous permettra de déceler notre tableau cible 
tables = pd.read_html(StringIO(html))

for i, table in enumerate(tables):
    print(f"Tableau {i} : {len(table)} lignes")

# On sélectionne le tableau contenant les salaires, on l’enregistre sur le bureau
df = max(tables, key=len)

df.to_excel(
    "tableau_complet.xlsx",
    index=False,
    engine="openpyxl"
)

print(f"{len(df)} lignes exportées dans tableau_complet.xlsx")

chemin_sortie = r"C:\Users\User\Desktop\tableau_complet.xlsx"

df.to_excel(
    chemin_sortie,
    index=False,
    engine="openpyxl"
)
print(f"Fichier enregistré ici : {chemin_sortie}")
