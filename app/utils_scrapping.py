import configparser
from PyPDF2 import PdfReader
import time
from datetime import datetime, timedelta
from app.data_list import sites
import os
from .config import get_config_value, get_pdf_path, set_config_value
import json
from urllib.request import urlopen
import locale
import re
import requests
from bs4 import BeautifulSoup
import ssl
from urllib.request import Request, urlopen

config = configparser.ConfigParser()
config.read('../config.ini')

def extract_2360(soup, checkbox_state = False, start_date=None, end_date=None):
    """
    Extraire les données de la table Materion depuis un fichier PDF et les retourner.

    Cette fonction ouvre un fichier PDF spécifié, lit le texte et extrait les informations
    pertinentes concernant l'Alliage 360. Elle retourne la date (numéro de la semaine) et
    la valeur correspondante extraite.
    
    Parameters:
    - soup (BeautifulSoup object): Objet BeautifulSoup du contenu web (non utilisé dans cette fonction).
    - checkbox_state (bool): État de la checkbox (non utilisé dans cette fonction).
    - start_date (datetime): Date de début pour l'extraction des données (non utilisé dans cette fonction).
    - end_date (datetime): Date de fin pour l'extraction des données (non utilisé dans cette fonction).
    
    Returns:
    - tuple: Un tuple contenant la date (numéro de la semaine) et la valeur extraite, ou 'err' si une erreur se produit.
    """
    
    pdf_path = get_config_value('SETTINGS', 'pdf_path')  # Récupérer le chemin du PDF depuis les paramètres
    name_pdf = get_config_value('SETTINGS', 'name_pdf')  # Récupérer le nom du PDF depuis les paramètres
    if not pdf_path:
        pdf_path = os.getcwd()  # Utiliser le répertoire courant si aucun chemin n'est spécifié
    path = f"{pdf_path}/{name_pdf}"  # Construire le chemin complet vers le fichier PDF
    
    try:
        with open(path, 'rb') as pdf_materion:  # Ouvrir le fichier PDF en mode lecture binaire
            reader_materion = PdfReader(pdf_materion)  # Créer un lecteur PDF
            page_materion = reader_materion.pages[0]  # Lire la première page du PDF
            text_materion = page_materion.extract_text()  # Extraire le texte de la page
            print('PDF lu')
            lines = text_materion.split('\n')  # Diviser le texte en lignes
            alloy_line = None
            date = None
            # Date d'aujourd'hui
            today = datetime.today()
            # Date de la veille
            yesterday = today - timedelta(days=1)
            # Numéro de la semaine
            week_number = yesterday.isocalendar()[1]
            for line in lines:
                print(line)
                if "As of" in line.lower():
                    date = line.split("As Of")[-1].strip()
                if line.startswith('Alloy 360'):
                    alloy_line = line
                    break
            if alloy_line is not None:
                # Récupérer la valeur de la 4ème colonne
                columns = alloy_line.split()
                if len(columns) >= 4:
                    price_eur = columns[4]
                else:
                    price_eur = None
                formatted_data = price_eur.replace('.', ',')
                date = f"Semaine {week_number}"
                print(date, formatted_data)
                return date, formatted_data  # Retourner la date et les données formatées extraites

    except FileNotFoundError:  # Gérer l'exception si le fichier PDF n'est pas trouvé
        print(f"Le fichier PDF '{name_pdf}' n'a pas été trouvé. Passage à autre chose.")
        date = 'date none'
        formatted_data = 'value none'  # Retourner 'err' si une erreur se produit
        return date, formatted_data

def extract_1AG1(soup, checkbox_state = False, start_date=None, end_date=None):
    """
    Extraire les données de la table Cookson depuis un JSON ou une page HTML.

    Cette fonction récupère des données depuis une URL JSON si une plage de dates est spécifiée,
    sinon elle extrait les données directement depuis une page HTML passée en paramètre.
    
    Parameters:
    - soup (BeautifulSoup object): Objet BeautifulSoup du contenu web à analyser.
    - checkbox_state (bool): État de la checkbox, détermine si une plage de dates doit être utilisée.
    - start_date (datetime): Date de début pour l'extraction des données.
    - end_date (datetime): Date de fin pour l'extraction des données.
    
    Returns:
    - tuple or list: Retourne soit un tuple de (date, données formatées), soit une liste de tuples si une plage de dates est utilisée.
    """

    url = 'https://www.cookson-clal-industrie.com/prix-des-metaux/'

    if checkbox_state and start_date and end_date:
        
        extracted_values = [] # Liste pour stocker les valeurs extraites
        current_date = start_date
        while current_date <= end_date:
            params = {
                'coursday': current_date.strftime('%d'),
                'coursmonth': current_date.strftime('%m'),
                'coursyear': current_date.strftime('%Y'),
            }
            response = requests.post(url, data=params, headers=headers, verify=False)
            print(response.status_code)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                tables = soup.find_all('div', class_="metal-table")
                table = tables[0]
                rows = table.find_all('div', class_="metal-table-item")
                if len(rows)>3:
                    row = rows[3]
                    value = row.find('div', class_="table-col-5")
                    value = value.text.strip()
                    date = datetime.strftime(current_date, "%d/%m/%Y")
                    print(date)
                    print(value)
                    formatted_data = value.replace("€", "").replace(".", ",")
                    if formatted_data != "Non renseigné":
                        extracted_values.append((date, formatted_data))
            else:
                print(f"Failed to retrieve data for {current_date.strftime('%Y-%m-%d')}")
            
            current_date += timedelta(days=1)
                    
        return extracted_values
    
    else:
        try:
            tables = soup.find_all("table", class_="main")
            table = tables[3]
            rows = table.find_all("tr", class_="lgn1")
            row = rows[1]
            columns = row.find_all("td")
            last_column = columns[4]

            # Extraire la date de la première colonne du tbody
            tbody = soup.find('tbody')
            first_td_in_tbody = soup.find('td')
            date_day = first_td_in_tbody.text.strip()


            # Extraire le texte de la quatrième colonne
            data = last_column.text.strip()
            formatted_data = data.replace('€/kg', '').replace(".", "")
            date = date_day.replace("Cours de Londres du ", " ")
            return date, formatted_data # Retourner la date et la valeur formatées extraite
        except:
            date = 'date none'
            formatted_data = 'value none'
            return date, formatted_data

def extract_1AG2(soup, checkbox_state = False, start_date=None, end_date=None):
    """
    Extraire les données de prix de l'argent depuis une URL JSON.
    
    Cette fonction récupère et traite les données de prix de l'argent depuis une URL JSON.
    Si une plage de dates est spécifiée (et que checkbox_state est True), la fonction retourne
    les données correspondant à cette plage. Sinon, elle retourne la dernière valeur disponible.
    
    Parameters:
    - soup (BeautifulSoup object): Objet BeautifulSoup du contenu web à analyser (non utilisé dans cette fonction).
    - checkbox_state (bool): État de la checkbox, détermine si une plage de dates doit être utilisée.
    - start_date (datetime): Date de début pour l'extraction des données.
    - end_date (datetime): Date de fin pour l'extraction des données.
    
    Returns:
    - tuple or list: Retourne soit un tuple de (date, données formatées), soit une liste de tuples si une plage de dates est utilisée.
    """
    
    url = "https://prices.lbma.org.uk/json/silver.json?r=211497526"

    # On crée le contexte SSL
    context = ssl._create_unverified_context()

    # On crée une requête avec un "User-Agent" de navigateur récent
    req = Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'})

    # On ouvre la session avec la requête ET le contexte
    response = urlopen(req, context=context).read()
    
    
    # response = urlopen(url).read() # Obtenir la réponse de l'URL
    data = json.loads(response) # Charger les données JSON

    if checkbox_state and start_date and end_date:

        extracted_values = [] # Liste pour stocker les valeurs extraites

        # Parcourir chaque entrée dans les données
        for entry in data:
            entry_date_str = entry.get("d")
            entry_date = datetime.strptime(entry_date_str, "%Y-%m-%d")
            date_data_obj = entry_date.date()
            
            if start_date <= date_data_obj <= end_date:
                value = entry['v'].pop(0)
                if value:
                    extracted_values.append((date_data_obj.strftime('%d/%m/%Y'), value))

        return extracted_values # Retourner la liste des valeurs extraites
    else:
        try:
            latest_prices = data[-1]
            first_value = latest_prices['v'].pop(0)
            data_value = latest_prices['d']

            date_object = datetime.strptime(data_value, '%Y-%m-%d')
            formatted_date = date_object.strftime('%d/%m/%Y')

            data = str(first_value)
            formatted_data = data.replace('.', ',')
            
            print(formatted_data)
            return formatted_date, formatted_data # Retourner la date et la valeur formatée extraite
        except:
            date = 'date none'
            formatted_data = 'value none'
            return date, formatted_data

def extract_3AL1(soup, checkbox_state=False, start_date=None, end_date=None):
    """
    Extraire les données depuis une table HTML basée sur une plage de dates spécifiée.
    
    Cette fonction parcourt une table HTML et extrait les données de chaque ligne basée sur
    une plage de dates spécifiée. Si la plage de dates est activée (checkbox_state=True) et
    que des dates de début et de fin sont fournies, la fonction extrait les données qui 
    correspondent à cette plage. Sinon, elle extrait les données de la première ligne de la table.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup contenant le contenu HTML à analyser.
        checkbox_state (bool, optional): Indique si une plage de dates doit être utilisée. Defaults to False.
        start_date (datetime.date, optional): Date de début de la plage. Defaults to None.
        end_date (datetime.date, optional): Date de fin de la plage. Defaults to None.
    
    Returns:
        tuple or list: Un tuple contenant la date et la valeur extraites, ou une liste de tuples
                       si une plage de dates est utilisée.
    """
    
    months = {
        'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 
        'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
        'November': '11', 'December': '12'
    }

    table = soup.find("table")
    rows = soup.find_all("tr")
    
    if checkbox_state and start_date and end_date:

        extracted_values = []
        
        # Parcourir chaque ligne de la table, en ignorant l'en-tête
        for row in rows[1:]:
            columns = row.find_all("td")
            
            if len(columns) < 2:  # nombre de colonne pour ligne séparatrice
                continue
            
            date_data_raw = columns[0].text.strip()

            
            if len(columns) >= 1:
                date_data_raw = columns[0].text.strip()
                
            
                # Conversion de la date du format "05. September 2023" à "05/09/2023"
                day, month_name, year = date_data_raw.replace('.', '').split()
                month_num = months.get(month_name, '00')  # Si le mois n'est pas trouvé, '00' est utilisé par défaut
                date_data = f"{day}/{month_num}/{year}"
            
                date_data_obj = datetime.strptime(date_data, "%d/%m/%Y").date()
            
                if start_date <= date_data_obj <= end_date:
                    fourth_column = columns[1]
                    data = fourth_column.text.strip()
                    formatted_data = data.replace(',', '').replace('.', ',')
                    extracted_values.append((date_data, formatted_data))
                    
        extracted_values.reverse()
        return extracted_values # Retourner la liste des valeurs extraites
                
    else:
        try:
            # Si la checkbox n'est pas cochée ou si les dates ne sont pas fournies, récupérer la première valeur
            second_row = rows[1]
            columns = second_row.find_all("td")
            fourth_column = columns[1]
            date_data_raw = columns[0].text.strip()
            
            # Conversion de la date et extraction des données
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
            
            data = fourth_column.text.strip()
            formatted_data = data.replace(',', '').replace('.', ',')
            
            return date_data, formatted_data # Retourner la date et la valeur extraites
        except:
            date = 'date none'
            formatted_data = 'value none'
            return date, formatted_data

def extract_1AU2(soup, checkbox_state = False, start_date = None, end_date = None):
    """
    Extraire les données de prix de l'or depuis une URL JSON basée sur une plage de dates spécifiée.
    
    Cette fonction récupère les données de prix de l'or depuis une URL JSON. Si une plage de dates
    est spécifiée (checkbox_state=True) et que des dates de début et de fin sont fournies, la fonction
    extrait les données qui correspondent à cette plage. Sinon, elle extrait la dernière valeur disponible
    dans les données JSON.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup contenant le contenu HTML à analyser (non utilisé dans cette fonction).
        checkbox_state (bool, optional): Indique si une plage de dates doit être utilisée. Defaults to False.
        start_date (datetime.date, optional): Date de début de la plage. Defaults to None.
        end_date (datetime.date, optional): Date de fin de la plage. Defaults to None.
    
    Returns:
        tuple or list: Un tuple contenant la date et la valeur extraites, ou une liste de tuples
                       si une plage de dates est utilisée.
    """

    url = "https://prices.lbma.org.uk/json/gold_pm.json?r=666323974"

    # On crée le contexte SSL (que vous avez déjà ajouté)
    context = ssl._create_unverified_context()

    # On crée une requête avec un "User-Agent" de navigateur récent
    req = Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'})

    # On ouvre la session avec la requête ET le contexte
    response = urlopen(req, context=context).read()

    #response = urlopen(url).read()
    data = json.loads(response)

    if checkbox_state and start_date and end_date:

        extracted_values = []

        # Parcourir chaque entrée dans les données JSON
        for entry in data:
            entry_date_str = entry.get("d")
            entry_date = datetime.strptime(entry_date_str, "%Y-%m-%d")
            date_data_obj = entry_date.date()
            
            # Vérifier si la date de l'entrée dans la plage spécifiée
            if start_date <= date_data_obj <= end_date:
                value = entry['v'].pop(0)
                if value:
                    extracted_values.append((date_data_obj.strftime('%d/%m/%Y'), value))
                    
        return extracted_values # Retourner la liste des valeurs extraites

    else:
        try:
            # Si aucune plage de dates spécifiée, obtenir la dernière valeur
            latest_prices = data[-1]
            first_value = latest_prices['v'].pop(0)
            date_value = latest_prices['d']

            date_object = datetime.strptime(date_value, '%Y-%m-%d')
            formatted_date = date_object.strftime('%d/%m/%Y')

            data = str(first_value)
            formatted_data = data.replace('.', ',').replace(" ", "")
            
            print(formatted_data)
            return formatted_date, formatted_data # Retourner la date et la valeur extraites
        except:
            formatted_date = 'date none'
            formatted_data = 'value none'
            return formatted_date, formatted_data

def extract_1AU3(soup, checkbox_state = False, start_date=None, end_date=None):
    """
    Extraire les données de prix de l'or depuis une URL JSON ou une page HTML basée sur une plage de dates spécifiée.
    
    Cette fonction récupère les données de prix de l'or depuis une URL JSON. Si une plage de dates
    est spécifiée (checkbox_state=True) et que des dates de début et de fin sont fournies, la fonction
    extrait les données qui correspondent à cette plage. Si aucune plage de dates n'est spécifiée,
    la fonction extrait les données depuis une table HTML présente dans le contenu de la page soup.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup contenant le contenu HTML à analyser.
        checkbox_state (bool, optional): Indique si une plage de dates doit être utilisée. Defaults to False.
        start_date (datetime.date, optional): Date de début de la plage. Defaults to None.
        end_date (datetime.date, optional): Date de fin de la plage. Defaults to None.
    
    Returns:
        tuple or list: Un tuple contenant la date et la valeur extraites, ou une liste de tuples
                       si une plage de dates est utilisée.
    """

    url = 'https://www.cookson-clal-industrie.com/prix-des-metaux/'


    if checkbox_state and start_date and end_date:
        extracted_values = []
        
        current_date = start_date
        while current_date <= end_date:
            params = {
                'coursday': current_date.strftime('%d'),
                'coursmonth': current_date.strftime('%m'),
                'coursyear': current_date.strftime('%Y'),
            }
            response = requests.post(url, data=params, verify=False)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                tables = soup.find_all('div', class_="metal-table")
                table = tables[1]
                rows = table.find_all('div', class_="metal-table-item")
                row = rows[1]
                value = row.find('div', class_="table-col-5")
                value = value.text.strip()
                date = datetime.strftime(current_date, "%d/%m/%Y")
                print(date)
                print(value)
                formatted_data = value.replace("€", "")
                if formatted_data != "Non renseigné":
                    extracted_values.append((date, formatted_data))
            else:
                print(f"Failed to retrieve data for {current_date.strftime('%Y-%m-%d')}")
            
            current_date += timedelta(days=1)
            
        return extracted_values # Retourner la liste de valeurs extraites
    
    else:
        try:
            # Si aucune plage de dates n'est spécifiée, extraire les données depuis la table HTML
            tables = soup.find_all("table", class_="main")
            print(len(tables))
            table = tables[4]
            row_lgn = table.find("tr", class_='lgn1')
            print(row_lgn)
            # Extraire la date de la première colonne du tbody
            tbody = soup.find('tbody')
            first_td_in_tbody = soup.find('td')
            date_day = first_td_in_tbody.text.strip()

            # Trouver la quatrième colonne de la table dans la troisième ligne
            columns = row_lgn.find_all("td")
            data = columns[4]
            data = data.text.strip()
            print(data)

            # Extraire le texte de la quatrième colonne
            data = data.replace(" ", "").replace("€/kg", "").replace(",", ".")
            formatted_data = data.replace('.', ',').replace('€', '').replace(' ', '')
            date = date_day.replace("Cours de Londres du ", " ")
            
            return date, formatted_data # Retourner la date et la valeur extraites
        except:
            date = 'date none'
            formatted_data = 'value none'
            return date, formatted_data

def extract_2B16(soup, checkbox_state=False, start_date=None, end_date=None):
    """
    Extraire et formater les données d'une table HTML basée sur une plage de dates spécifiée.
    
    Cette fonction parcourt une table HTML pour extraire des données. Si une plage de dates est
    spécifiée (checkbox_state=True) et que des dates de début et de fin sont fournies, la fonction
    extrait les données qui correspondent à cette plage. Si aucune plage de dates n'est spécifiée,
    la fonction extrait les données de la première ligne de la table.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup contenant le contenu HTML à analyser.
        checkbox_state (bool, optional): Indique si une plage de dates doit être utilisée. Defaults to False.
        start_date (datetime.date, optional): Date de début de la plage. Defaults to None.
        end_date (datetime.date, optional): Date de fin de la plage. Defaults to None.
    
    Returns:
        tuple or list: Un tuple contenant la date et la valeur extraites, ou une liste de tuples
                       si une plage de dates est utilisée.
    """
    
    months = {
        'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05',
        'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
        'November': '11', 'December': '12'
    }

    table = soup.find("table")
    rows = soup.find_all("tr")
    
    if checkbox_state and start_date and end_date:

        extracted_values = []
        
        # Parcourir chaque ligne de la table, en ignorant l'en-tête
        for row in rows[1:]:
            columns = row.find_all("td")
            
            # Ignorer les lignes qui n'ont pas assez de colonnes
            if len(columns) < 2:
                continue
            
            date_data_raw = columns[0].text.strip()

            # Conversion de la date en format "JJ/MM/AAAA"
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
        
            date_data_obj = datetime.strptime(date_data, "%d/%m/%Y").date()

            # Vérifier si la date est dans la plage spécifiée
            if start_date <= date_data_obj <= end_date:
                fourth_column = columns[1]
                data = fourth_column.text.strip()
                formatted_data = data.replace(',', '').replace('.', ',')
                extracted_values.append((date_data, formatted_data))
                
        extracted_values.reverse()
        return extracted_values # Retourner la liste de valeurs extraites
                
    else:
        try:
            # Si aucune plage de dates n'est spécifiée, récupérer la première valeur
            second_row = rows[1]
            columns = second_row.find_all("td")
            thirth_column = columns[1]
            date_data_raw = columns[0].text.strip()
            
            # Conversion et formatage de la date et des données
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
            
            data = thirth_column.text.strip()
            formatted_data = data.replace(',', '').replace('.', ',')
            
            return date_data, formatted_data # Retourner la date et valeur extraites
        except:
            date_data = 'date none'
            formatted_data = 'value none'
            return date_data, formatted_data

def extract_3CU1(soup, checkbox_state=False, start_date=None, end_date=None):
    """
    Extraire et formater les données d'une table HTML basée sur une plage de dates spécifiée.
    
    Cette fonction parcourt une table HTML pour extraire des données. Si une plage de dates est
    spécifiée (checkbox_state=True) et que des dates de début et de fin sont fournies, la fonction
    extrait les données qui correspondent à cette plage. Si aucune plage de dates n'est spécifiée,
    la fonction extrait les données de la première ligne de la table.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup contenant le contenu HTML à analyser.
        checkbox_state (bool, optional): Indique si une plage de dates doit être utilisée. Defaults to False.
        start_date (datetime.date, optional): Date de début de la plage. Defaults to None.
        end_date (datetime.date, optional): Date de fin de la plage. Defaults to None.
    
    Returns:
        tuple or list: Un tuple contenant la date et la valeur extraites, ou une liste de tuples
                       si une plage de dates est utilisée.
    """

    months = {
        'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 
        'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
        'November': '11', 'December': '12'
    }

    table = soup.find("table")
    rows = soup.find_all("tr")
    
    if checkbox_state and start_date and end_date:

        extracted_values = []
        
        # Parcourir chaque ligne de la table, en ignorant l'en-tête
        for row in rows[1:]:
            columns = row.find_all("td")

            # Ignorer les lignes qui n'ont pas assez de colonnes
            if len(columns) < 2:
                continue
            
            date_data_raw = columns[0].text.strip()
                
            # Conversion de la date du format "05. September 2023" à "05/09/2023"
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
        
            date_data_obj = datetime.strptime(date_data, "%d/%m/%Y").date()

            # Vérifier si la date est dans la plage spécifiée
            if start_date <= date_data_obj <= end_date:
                fourth_column = columns[1]
                data = fourth_column.text.strip()
                formatted_data = data.replace(',', '').replace('.', ',')
                extracted_values.append((date_data, formatted_data))
        extracted_values.reverse()
        return extracted_values # Retourner une liste de valeurs extraites
                
    else:
        try:
            # Si aucune plage n'est spécifiée, récupérer la première valeur
            second_row = rows[1]
            columns = second_row.find_all("td")
            thirth_column = columns[1]
            date_data_raw = columns[0].text.strip()
            
            # Conversion et formatage de la date et des données
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
            
            data = thirth_column.text.strip()
            formatted_data = data.replace(',', '').replace('.', ',')
            
            return date_data, formatted_data # Retourner la date et valeur extraites
        except:
            date_data = 'date none'
            formatted_data = 'value none'
            return date_data, formatted_data

def extract_3CU3(soup, checkbox_state = False, start_date=None, end_date=None):
    """
    Extraire et formater les données d'une table HTML basée sur une plage de dates spécifiée.
    
    Cette fonction parcourt une table HTML pour extraire des données. Si une plage de dates est
    spécifiée (checkbox_state=True) et que des dates de début et de fin sont fournies, la fonction
    extrait les données qui correspondent à cette plage. Si aucune plage de dates n'est spécifiée,
    la fonction extrait les données de la première ligne de la table.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup contenant le contenu HTML à analyser.
        checkbox_state (bool, optional): Indique si une plage de dates doit être utilisée. Defaults to False.
        start_date (datetime.date, optional): Date de début de la plage. Defaults to None.
        end_date (datetime.date, optional): Date de fin de la plage. Defaults to None.
    
    Returns:
        tuple or list: Un tuple contenant la date et la valeur extraites, ou une liste de tuples
                       si une plage de dates est utilisée.
    """
    
    months = {
        'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05',
        'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
        'November': '11', 'December': '12'
    }

    table = soup.find("table")
    rows = soup.find_all("tr")
    
    if checkbox_state and start_date and end_date:

        extracted_values = []
        
        # Parcourir chaque ligne de la tablen en ignorant l'en-tête
        for row in rows[1:]:
            columns = row.find_all("td")

            # Ignorer les lignes qui n'ont pas assez de colonnes
            if len(columns) < 2:
                continue
            
            date_data_raw = columns[0].text.strip()
                
            # Conversion de la date en format "JJ/MM/AAAA"
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
        
            date_data_obj = datetime.strptime(date_data, "%d/%m/%Y").date()

            # Vérifier si la date est dans la plage spécifiée
            if start_date <= date_data_obj <= end_date:
                fourth_column = columns[1]
                data = fourth_column.text.strip()
                formatted_data = data.replace(',', '').replace('.', ',')
                extracted_values.append((date_data, formatted_data))
            
        extracted_values.reverse()
        return extracted_values # Retourner la liste des données extraites
                
    else:
        try:
            # Si aucune plage n'est spécifiée, récupérer la première valeur
            second_row = rows[1]
            columns = second_row.find_all("td")
            second_column = columns[1]
            date_data_raw = columns[0].text.strip()
            
            # Conversion et formatage de la date et des données
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
            
            data = second_column.text.strip()
            formatted_data = data.replace(',', '').replace('.', ',')
            
            return date_data, formatted_data # Retourner la date et valeur extraites
        except:
            date_data = 'date none'
            formatted_data = 'value none'
            return date_data, formatted_data

def extract_2CUB(soup, checkbox_state = False, start_date=None, end_date=None):
    """
    Extraire les données spécifiques d'un fichier PDF et les formater.
    
    Cette fonction est conçue pour lire un fichier PDF spécifié, extraire du texte à partir
    d'une page spécifique du PDF, et ensuite traiter ce texte pour récupérer des informations
    pertinentes basées sur certaines conditions et formats prédéfinis.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup, non utilisé dans cette fonction.
        start_date (datetime, optional): Date de début, non utilisée dans cette fonction. Defaults to None.
        end_date (datetime, optional): Date de fin, non utilisée dans cette fonction. Defaults to None.
    
    Returns:
        tuple: Un tuple contenant la date formatée et les données extraites et formatées.
    """

    pdf_path = get_config_value('SETTINGS', 'pdf_path')
    name_pdf = get_config_value('SETTINGS', 'name_pdf')

    if not pdf_path:
        pdf_path = os.getcwd()

    path = f"{pdf_path}/{name_pdf}"

    try:
        with open(path, 'rb') as pdf_materion:
            reader_materion = PdfReader(pdf_materion)
            page_materion = reader_materion.pages[0]
            text_materion = page_materion.extract_text()
            print('PDF lu')

            lines = text_materion.split('\n')

            alloy_line = None
            date = None
            # Date d'aujourd'hui
            today = datetime.today()

            # Date de la veille
            yesterday = today - timedelta(days=1)

            # Numéro de la semaine
            week_number = yesterday.isocalendar()[1]

            for line in lines:
                if line.startswith('Alloy 25'):
                    alloy_line = line
                    break

            if alloy_line is not None:
                # Récupérer la valeur de la 4ème colonne
                columns = alloy_line.split()
                if len(columns) >= 4:
                    price_eur = columns[4]
                else:
                    price_eur = None

                formatted_data = price_eur.replace('.', ',')
                date = f" Semaine {week_number}"
                
                return date, formatted_data # Retourner la date et valeur extraites
            
    except FileNotFoundError:
        print(f"Le fichier PDF '{name_pdf}' n'a pas été trouvé. Passage à autre chose.")
        date = 'date none'
        formatted_data = 'value none'
        return date, formatted_data

def extract_2M30(soup, checkbox_state = False, start_date=None, end_date=None):
    """
    Extraire et formater les données de prix des métaux depuis une URL spécifiée.
    
    Cette fonction récupère les données depuis une URL, les traite et les formate selon
    que certaines conditions soient remplies ou non, comme l'état d'une checkbox et des dates spécifiées.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup pour analyser le contenu HTML.
        checkbox_state (bool, optional): État de la checkbox pour déterminer le mode de traitement. Defaults to False.
        start_date (datetime, optional): Date de début pour filtrer les données. Defaults to None.
        end_date (datetime, optional): Date de fin pour filtrer les données. Defaults to None.
    
    Returns:
        tuple: Un tuple contenant des dates formatées et des données extraites et formatées.
    """
    
    url = 'https://www.wieland.com/en/ajax/metal-prices/general?refKey=2121'


    if checkbox_state and start_date and end_date:
        response = requests.get(url, verify=False)
        json_data = response.json()

        # Extraire et traiter les données si les clés nécessaires sont présentes
        if 'content' in json_data and 'chart' in json_data['content']:
            chart_data = json_data['content']['chart']

            if 'labels' in chart_data and 'data' in chart_data:

                labels = chart_data['labels']
                data = chart_data['data']

                extracted_values = []
    
                for label, value in zip(labels, data):
                    # Convertir la date du label en objet datetime
                    label_date = datetime.strptime(label, '%m/%d/%Y')
                    label_date = label_date.date()
                    formatted_date = label_date.strftime('%d/%m/%Y')
                    
                    # Vérifier si la date est entre start_date et end_date
                    if start_date <= label_date <= end_date:
                        extracted_values.append((formatted_date, value))
                
                return extracted_values # Retourner la liste de données extraites
            
            else:
                print("Les clés 'labels' et/ou 'data' ne sont pas présentes dans les données.")
        else:
            print("Les clés 'content' et/ou 'chart' ne sont pas présentes dans les données.")
                        

    else:
        try:
            table = soup.find('table', class_='metalinfo-table')
            # Trouver toutes les lignes (tr) à l'intérieur de cette table
            rows = soup.find_all('tr')
            second_row = rows[23]

            # Trouver toutes les colonnes (td) de la ligne spécifiée
            columns = second_row.find_all('td')
            second_column = columns[1]
            data = second_column.text.strip()
            formatted_data = data.replace(',', '').replace('.', ',')
            # Trouver la date dans le tag <p class="date small">
            date_tag = soup.find("p", class_="date small")
            raw_date_data = date_tag.text.strip() if date_tag else "Date not found"

            # Convertir la date au format souhaité
            locale.setlocale(locale.LC_TIME, "en_US.UTF-8")
            try:
                # Supprimer "Value from " pour obtenir seulement la date
                clean_date_data = raw_date_data.replace("Value from ", "").strip()
                print(f'clean data : "{clean_date_data}"')
                # Convertir la chaîne de date au format souhaité
                datetime_obj = datetime.strptime(clean_date_data, '%b %d, %Y')
                formatted_date = datetime_obj.strftime('%d/%m/%Y')
            except ValueError:
                formatted_date = "Invalid date format"

            return formatted_date, formatted_data # Retourner la date et valeur extraites
        except:
            formatted_date = 'date none'
            formatted_data = 'value none'
            return formatted_date, formatted_data

def extract_2M37(soup, checkbox_state = False, start_date=None, end_date=None):
    """
    Extraire et formater les données de prix des métaux depuis une page HTML.
    
    Cette fonction récupère et formate les données depuis une page HTML en se basant sur l'état
    d'une checkbox et des dates de début et de fin spécifiées.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup pour analyser le contenu HTML.
        checkbox_state (bool, optional): État de la checkbox pour déterminer le mode de traitement. Defaults to False.
        start_date (datetime, optional): Date de début pour filtrer les données. Defaults to None.
        end_date (datetime, optional): Date de fin pour filtrer les données. Defaults to None.
    
    Returns:
        list or tuple: Une liste de tuples contenant des dates et des données formatées si la checkbox est cochée,
                       sinon un tuple contenant une date et une donnée formatées.
    """

    months = {
        'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05',
        'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
        'November': '11', 'December': '12'
    }

    table = soup.find("table")
    rows = soup.find_all("tr")
    
    if checkbox_state and start_date and end_date:

        extracted_values = []
        # Parcourir chaque ligne de la table en ignorant l'en-tête
        for row in rows[1:]:
            columns = row.find_all("td")

            # Ignorer les lignes qui n'ont pas assez de colonnes
            if len(columns) < 2:
                continue
            
            date_data_raw = columns[0].text.strip()

            # Conversion de la date au format "JJ/MM/AAAA"
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
        
            date_data_obj = datetime.strptime(date_data, "%d/%m/%Y").date()

            # Vérifier si la date est dans la plage de dates
            if start_date <= date_data_obj <= end_date:
                second_column = columns[1]
                data = second_column.text.strip()
                formatted_data = data.replace(',', '').replace('.', ',')
                extracted_values.append((date_data, formatted_data))

        extracted_values.reverse()
        return extracted_values # Retourner la liste de données extraites
                
    else:
        try:
            # Si aucune plage n'est spécifiée, récupérer la première valeur
            second_row = rows[1]
            columns = second_row.find_all("td")
            second_column = columns[1]
            date_data_raw = columns[0].text.strip()
            
            # Conversion de la date et extraction des données
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
            
            data = second_column.text.strip()
            formatted_data = data.replace(',', '').replace('.', ',')
            
            return date_data, formatted_data # Retourner la date et valeur extraites
        except:
            date_data = 'date none'
            formatted_data = 'value none'
            return date_data, formatted_data

def extract_3NI1(soup, checkbox_state = False, start_date=None, end_date=None):
    """
    Extraire les données de nickel depuis une page HTML.
    
    Cette fonction récupère les données de nickel depuis une table HTML spécifique, 
    en se basant sur l'état d'une checkbox et des dates de début et de fin spécifiées.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup pour analyser le contenu HTML.
        checkbox_state (bool, optional): État de la checkbox pour déterminer le mode de traitement. Defaults to False.
        start_date (datetime, optional): Date de début pour filtrer les données. Defaults to None.
        end_date (datetime, optional): Date de fin pour filtrer les données. Defaults to None.
    
    Returns:
        list or tuple: Une liste de tuples contenant des dates et des données formatées si la checkbox est cochée,
                       sinon un tuple contenant une date et une donnée formatées.
    """
    

    if checkbox_state and start_date and end_date:
        tables = soup.find_all('table', class_='table table-condensed table-hover table-striped')
        rows = tables[1].find_all("tr")
        extracted_values = []

        # Parcourir chaque ligne de la table en ignorant l'en-tête
        for row in rows[1:]:
            columns = row.find_all("td")

            if len(columns) >= 1:
                date_data_raw = columns[0].text.strip()
                
                # Utiliser une expression régulière pour extraire une date valide
                date_match = re.search(r"(\d{2}\.\d{2}\.\d{4})", date_data_raw)
                if date_match:
                    date_data_raw = date_match.group(1)
                else:
                    continue  # Passer à la prochaine ligne si aucune date valide n'est trouvée

                try:
                    date_obj = datetime.strptime(date_data_raw, '%d.%m.%Y')
                    
                    date_data_obj = date_obj.date()
                    
                    if start_date <= date_data_obj <= end_date:
                        value_data = columns[1].text.strip().replace('.', '')
                        extracted_values.append((date_obj.strftime('%d/%m/%Y'), value_data))
                except ValueError as e:
                    print(f"Erreur lors de la conversion de la date: {e}")
                    
        return extracted_values # Retourner la liste des données extraites

    else:
        try:
            tables = soup.find_all('table', class_='table table-condensed table-hover table-striped')
            rows = tables[1].find_all("tr")
        # S'assurer qu'il y a au moins deux tables et sélectionner la deuxième
            if len(tables) > 1:
                table = tables[1]
                
                # Obtenir la table qui contient les données voulues
                first_table = table.find('table', class_='table table-condensed table-hover table-striped')
                # Obtenir la première ligne de la table (en excluant l'en-tête)
                last_row = first_table.find_all('tr')[-1] if table else None
                print(last_row)
                if last_row:
                    columns = last_row.find_all('td')
                    print(columns)
                    # S'assurer qu'il y a au moins deux colonnes
                    if len(columns) >= 2:
                        # Extraire et nettoyer la date et la valeur
                        date_str = columns[0].text.strip()
                        print(f"DATE :", date_str)
                        value_str = columns[1].text.strip().replace('.', '')
                        print(f"VALUE :", value_str)
                        
                        # Convertir la date au format d/m/Y
                        try:
                            date_obj = datetime.strptime(date_str, '%d.%m.%Y')
                            formatted_date = date_obj.strftime('%d/%m/%Y')
                        except ValueError:
                            print(f"Erreur de format de date : {date_str}")
                            formatted_date = date_str  # Garder la date telle quelle si la conversion échoue
                        print(formatted_date, value_str)
                        
                        return formatted_date, value_str # Retourner la date et valeur extraites
                    else:
                        print("Les colonnes de date et de valeur sont manquantes.")
                        return None, None
                else:
                    print("Aucune ligne de données trouvée dans la table.")
                    return None, None
            else:
                print("La deuxième table est introuvable.")
                return None, None
        except:
            formatted_date = 'date none'
            formatted_data = 'value none'
            return formatted_date, formatted_data

def extract_3SN1(soup, checkbox_state = None, start_date=None, end_date=None):
    """
    Extraire les données d'étain depuis une page HTML.
    
    Cette fonction récupère les données d'étain depuis une table HTML spécifique, 
    en se basant sur l'état d'une checkbox et des dates de début et de fin spécifiées.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup pour analyser le contenu HTML.
        checkbox_state (bool, optional): État de la checkbox pour déterminer le mode de traitement. Defaults to None.
        start_date (datetime, optional): Date de début pour filtrer les données. Defaults to None.
        end_date (datetime, optional): Date de fin pour filtrer les données. Defaults to None.
    
    Returns:
        list or tuple: Une liste de tuples contenant des dates et des données formatées si la checkbox est cochée,
                       sinon un tuple contenant une date et une donnée formatées.
    """
    
    

    if checkbox_state and start_date and end_date:
        tables = soup.find_all('table', class_='table table-condensed table-hover table-striped')
        rows = tables[1].find_all("tr")
        extracted_values = []

        # Parcourir chaque ligne de la table en ignorant l'en-tête
        for row in rows[1:]:
            columns = row.find_all("td")

            if len(columns) >= 1:
                date_data_raw = columns[0].text.strip()
                
                # Utiliser une expression régulière pour extraire une date valide
                date_match = re.search(r"(\d{2}\.\d{2}\.\d{4})", date_data_raw)
                if date_match:
                    date_data_raw = date_match.group(1)
                else:
                    continue  # Passer à la prochaine ligne si aucune date valide n'est trouvée

                try:
                    date_obj = datetime.strptime(date_data_raw, '%d.%m.%Y')
                    
                    date_data_obj = date_obj.date()
                    
                    if start_date <= date_data_obj <= end_date:
                        value_data = columns[1].text.strip().replace(',', '').replace('.', ',')
                        extracted_values.append((date_obj.strftime('%d/%m/%Y'), value_data))
                except ValueError as e:
                    print(f"Erreur lors de la conversion de la date: {e}")
        
        return extracted_values

    else:
        try:
            tables = soup.find_all('table', class_='table table-condensed table-hover table-striped')
            rows = tables[1].find_all("tr")
        # S'assurer qu'il y a au moins deux tables et sélectionner la deuxième
            if len(tables) > 1:
                table = tables[1]
                
                # Obtenir la table qui contient les données voulues
                first_table = table.find('table', class_='table table-condensed table-hover table-striped')
                # Obtenir la première ligne de la table (en excluant l'en-tête)
                last_row = first_table.find_all('tr')[-1] if table else None
                print(last_row)
                if last_row:
                    columns = last_row.find_all('td')
                    print(columns)
                    # S'assurer qu'il y a au moins deux colonnes
                    if len(columns) >= 2:
                        # Extraire et nettoyer la date et la valeur
                        date_str = columns[0].text.strip()
                        print(f"DATE :", date_str)
                        value_str = columns[1].text.strip().replace('.', '')
                        print(f"VALUE :", value_str)
                        
                        # Convertir la date au format d/m/Y
                        try:
                            date_obj = datetime.strptime(date_str, '%d.%m.%Y')
                            formatted_date = date_obj.strftime('%d/%m/%Y')
                        except ValueError:
                            print(f"Erreur de format de date : {date_str}")
                            formatted_date = date_str  # Garder la date telle quelle si la conversion échoue
                            
                        print(formatted_date, value_str)
                        return formatted_date, value_str # Retourner la date et valeur extraites
                    else:
                        print("Les colonnes de date et de valeur sont manquantes.")
                        return None, None
                else:
                    print("Aucune ligne de données trouvée dans la table.")
                    return None, None
            else:
                print("La deuxième table est introuvable.")
                return None, None
        except:
            formatted_date = 'date none'
            formatted_data = 'value none'
            return formatted_date, formatted_data

def extract_3ZN1(soup, checkbox_state=False, start_date=None, end_date=None):
    """
    Extraire les données d'une table HTML en fonction de l'état de la checkbox et d'une plage de dates.
    
    Cette fonction extrait les données d'une table trouvée dans une page HTML parsée.
    Si la checkbox est cochée et que des dates de début et de fin sont fournies, la fonction extrait
    plusieurs valeurs de la table qui correspondent à la plage de dates.
    Sinon, elle extrait une seule valeur de la table.
    
    Args:
        soup (BeautifulSoup): Objet BeautifulSoup contenant la page HTML parsée.
        checkbox_state (bool): État de la checkbox qui détermine le mode d'extraction des données.
        start_date (datetime.date, optional): Date de début de la plage pour filtrer les données.
        end_date (datetime.date, optional): Date de fin de la plage pour filtrer les données.
        
    Returns:
        list or tuple: Si la checkbox est cochée, retourne une liste de tuples (date, valeur).
                       Sinon, retourne un tuple (date, valeur).
    """

    months = {
        'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 
        'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
        'November': '11', 'December': '12'
    }

    
    
    if checkbox_state and start_date and end_date:
        table = soup.find("table")
        rows = soup.find_all("tr")
        extracted_values = []
        # Parcourir chaque ligne de la table en ignorant l'en-tête
        for row in rows[1:]:
            columns = row.find_all("td")

            # Eviter les lignes qui n'ont pas assez de colonnes
            if len(columns) < 2:
                continue
            
            date_data_raw = columns[0].text.strip()
                
            # Conversion de la date du format "05. September 2023" à "05/09/2023"
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
        
            date_data_obj = datetime.strptime(date_data, "%d/%m/%Y").date()
        
            if start_date <= date_data_obj <= end_date:
                fourth_column = columns[1]
                data = fourth_column.text.strip()
                formatted_data = data.replace(',', '').replace('.', ',')
                extracted_values.append((date_data, formatted_data))
        
        extracted_values.reverse()
        return extracted_values # Retourner la liste de données extraites
                
    else:
        try:
            table = soup.find("table")
            rows = soup.find_all("tr")
            # Si aucune plage n'est spécifiée, récupérer la première valeur
            second_row = rows[1]
            columns = second_row.find_all("td")
            second_column = columns[1]
            date_data_raw = columns[0].text.strip()
            
            # Conversion de la date et extraction des données
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
            
            data = second_column.text.strip()
            formatted_data = data.replace(',', '').replace('.', ',')
            
            return date_data, formatted_data # Retourner la date et valeur extraites
        except:
            date_data = 'date none'
            formatted_data = 'value none'
            return date_data, formatted_data

def extract_ZLME(soup, checkbox_state=False, start_date=None, end_date=None):
    
    months = {
            'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 
            'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
            'November': '11', 'December': '12'
        }
    if checkbox_state and start_date and end_date:
        table = soup.find("table")
        rows = soup.find_all("tr")
        extracted_values = []
        # Parcourir chaque ligne de la table en ignorant l'en-tête
        for row in rows[1:]:
            columns = row.find_all("td")

            # Eviter les lignes qui n'ont pas assez de colonnes
            if len(columns) < 2:
                continue
            
            date_data_raw = columns[0].text.strip()
                
            # Conversion de la date du format "05. September 2023" à "05/09/2023"
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
        
            date_data_obj = datetime.strptime(date_data, "%d/%m/%Y").date()
        
            if start_date <= date_data_obj <= end_date:
                fourth_column = columns[1]
                data = fourth_column.text.strip()
                formatted_data = data.replace('.', ',')
                extracted_values.append((date_data, formatted_data))
        
        extracted_values.reverse()
        return extracted_values # Retourner la liste de données extraites
    
    else:
        try:
            table = soup.find("table")
            rows = soup.find_all("tr")
            second_row = rows[1]
            columns = second_row.find_all("td")
            second_column = columns[1]
            date_data_raw = columns[0].text.strip()
            
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
            
            data = second_column.text.strip()
            formatted_data = data.replace('.', ',')
            return date_data, formatted_data

        
        except:
            date_data = 'date none'
            formatted_data = 'value none'
            return date_data, formatted_data
        
def extract_EURX(soup, checkbox_state=False, start_date=None, end_date=None):
    
    months = {
            'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 
            'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
            'November': '11', 'December': '12'
        }
    if checkbox_state and start_date and end_date:
        table = soup.find("table")
        rows = soup.find_all("tr")
        extracted_values = []
        # Parcourir chaque ligne de la table en ignorant l'en-tête
        for row in rows[1:]:
            columns = row.find_all("td")

            # Eviter les lignes qui n'ont pas assez de colonnes
            if len(columns) < 2:
                continue
            
            date_data_raw = columns[0].text.strip()
                
            # Conversion de la date du format "05. September 2023" à "05/09/2023"
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
        
            date_data_obj = datetime.strptime(date_data, "%d/%m/%Y").date()
        
            if start_date <= date_data_obj <= end_date:
                fourth_column = columns[1]
                data = fourth_column.text.strip()
                formatted_data = data.replace('.', ',')
                extracted_values.append((date_data, formatted_data))
        
        extracted_values.reverse()
        return extracted_values # Retourner la liste de données extraites
    
    else:
        try:
            table = soup.find("table")
            rows = soup.find_all("tr")
            second_row = rows[1]
            columns = second_row.find_all("td")
            second_column = columns[1]
            date_data_raw = columns[0].text.strip()
            
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data = f"{day}/{month_num}/{year}"
            
            data = second_column.text.strip()
            formatted_data = data.replace('.', ',')
            return date_data, formatted_data

        
        except:
            date_data = 'date none'
            formatted_data = 'value none'
            return date_data, formatted_data

# Extraction données pour 1AG3 (EL) Deja importé
# def extract_1AG3(soup):
#     """Extraire les données de la table Cookson et les ajouter au classeur Excel"""
#     table = soup.find("table")
#     rows = table.find_all("tr")
#     second_row = rows[1]

#     columns = second_row.find_all("td")
    
#     # Récupérer la date de la première colonne
#     first_column = columns[0]
#     date_data = first_column.text.strip()
    
#     # Récupérer la valeur de la quatrième colonne
#     fourth_column = columns[1]
#     data = fourth_column.text.strip()
#     formatted_data = data.replace('.', ',')

#     # Retourner la date et la valeur de la quatrième colonne
#     return date_data, formatted_data

def extract_CUSN(soup, checkbox_state = False, start_date=None, end_date=None):
    pdf_path = get_config_value('SETTINGS', 'pdf_path')  # Récupérer le chemin du PDF depuis les paramètres
    name_pdf = 'Cours_metaux_template.pdf'  # Récupérer le nom du PDF depuis les paramètres
    if not pdf_path:
        pdf_path = os.getcwd()  # Utiliser le répertoire courant si aucun chemin n'est spécifié
    path = f"{pdf_path}/{name_pdf}"  # Construire le chemin complet vers le fichier PDF
    
    try:
        # Ouvrir le fichier PDF en mode lecture binaire
        with open(path, "rb") as file:
            reader = PdfReader(file)
            page_pdf = reader.pages[0]
            text_pdf = page_pdf.extract_text()
            lines = text_pdf.split('\n')
            date = None

            

        # Rechercher la ligne contenant "CUSN6" et extraire la valeur
        for line in lines:
            print(line)
            date_match = re.search(r"COURS APPLICABLES LE : (\d{2}/\d{2}/\d{4})", line)
            if date_match:
                date = date_match.group(1) if date_match else "Date non trouvée"
                pass
            # Extraire la valeur après "CUSN6"
            value_matches = re.findall(r"CUSN6 (\d+,\d+) EUR 1TO", line)
            if value_matches:
                value = value_matches[-1] if value_matches else "Valeur non trouvée"
        
        return date, value
    except FileNotFoundError:
        print("Cours SN6 pas trouvé")
        date = 'date none'
        formatted_data = 'value none'
        return date, formatted_data