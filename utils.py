import json
import os
import logging
import pytz
from datetime import datetime


def get_today():
    time_now = pytz.timezone('Indian/Antananarivo')
    today = datetime.now(time_now)
    return today


def write_log(logs, level=None):
    log_dir = 'logs'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # log_file = os.path.join('logs', f'{get_today().strftime('%d-%m-%Y')}.log')
    log_file = os.path.join(log_dir, f"{get_today().strftime('%d-%m-%Y')}.log")


    logging.basicConfig(filename=log_file, encoding='utf-8', level=level,
                        format='%(asctime)s - %(name)s - %(levelname)s : %(message)s')

    logger = logging.getLogger(__name__)

    if level == logging.ERROR:
        logger.error(logs)
    elif level == logging.INFO:
        logger.info(logs)
    elif level == logging.CRITICAL:
        logger.critical(logs)
    else:
        logger.exception(logs)


def format_number(number):
    try:
        return '{:,.2f}'.format(float(number))
    except ValueError:
        return number  # Handle non-numeric values gracefully


def extract_values_in_json(find):
    sql_values = {}

    try:
        with open('config.json', 'r') as file:
            json_data = json.load(file)
        if find in json_data:
            sql_data = json_data[find]
            for key, value in sql_data.items():
                sql_values[key] = value
    except FileNotFoundError:
        print("'config.json' Introuvable !")
    except json.decoder.JSONDecodeError:
        print("Erreur de load de donne json")
    except Exception as e:
        write_log(str(e), level='ERROR')
    return sql_values
