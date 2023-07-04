import os
import platform
import glob
import json
import fdb
import openpyxl
from datetime import datetime
import shutil


def create_temp_cows_folder():
    appdata_temp_dir = os.path.join(os.getenv('APPDATA'), 'TempCows')
    os.makedirs(appdata_temp_dir, exist_ok=True)
    return appdata_temp_dir


def delete_temp_cows_folder(appdata_temp_dir):
    if os.path.exists(appdata_temp_dir):
        for file_name in os.listdir(appdata_temp_dir):
            file_path = os.path.join(appdata_temp_dir, file_name)
            os.remove(file_path)
        os.rmdir(appdata_temp_dir)


def find_fdb_files(folder_path):
    fdb_files = []
    pattern = os.path.join(folder_path, "**", "*.fdb")
    for file_path in glob.iglob(pattern, recursive=True):
        if os.path.isfile(file_path):
            fdb_files.append(file_path)
    return fdb_files


def create_database_list(fdb_files, config):
    database_list = []
    for file_path in fdb_files:
        database = {}
        database["hostname"] = config["hostname"]
        database["database_path"] = file_path
        database["username"] = config["username"]
        database["password"] = config["password"]
        database_list.append(database)
    return database_list


def save_to_json(database_list):
    with open("database.json", "w", encoding="utf-8") as file:
        json.dump(database_list, file, indent=4, ensure_ascii=False)


def create_config_file():
    config = {}
    config["hostname"] = input("Введите имя хоста: ")
    config["username"] = input("Введите имя пользователя: ")
    config["password"] = input("Введите пароль: ")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_file_path = os.path.join(script_dir, "config.json")

    with open(config_file_path, "w", encoding="utf-8") as config_file:
        json.dump(config, config_file, indent=4, ensure_ascii=False)

    print(f"Файл конфигурации создан: {config_file_path}")
    return config


def execute_sql_query(conn_str, sql_query):
    connection = None
    try:
        connection = fdb.connect(**conn_str)
        cursor = connection.cursor()
        cursor.execute(sql_query)
        result_set = cursor.fetchall()
        return result_set
    except fdb.Error as e:
        print(f"Ошибка при выполнении SQL-запроса: {e}")
    finally:
        if connection:
            connection.close()


def save_to_excel(result_set, file_path):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    for row in result_set:
        worksheet.append(row)
    workbook.save(file_path)


def create_query_file(query_file_path):
    default_query = "SELECT * FROM table_name"
    with open(query_file_path, "w", encoding="utf-8") as query_file:
        query_file.write(default_query)
    print(f"Файл query.sql создан: {query_file_path}")


def main():
    operating_system = platform.system() + ' ' + platform.release() + ' ' + platform.version()
    print(f"Версия операционной системы: {operating_system}")

    print(r"""
   ___ _____      _____   ___                   _      
  / __/ _ \ \    / / __| | _ \___ _ __  ___ _ _| |_ ___
 | (_| (_) \ \/\/ /\__ \ |   / -_) '_ \/ _ \ '_|  _(_-<
  \___\___/ \_/\_/ |___/ |_|_\___| .__/\___/_|  \__/__/
                                 |_|                   
    """)

    appdata_temp_dir = create_temp_cows_folder()

    folder_path = input("Введите путь к папке для поиска файлов .fdb: ")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_file_path = os.path.join(script_dir, "config.json")
    report_dir = os.path.join(script_dir, "reports")
    query_file_path = os.path.join(script_dir, "query.sql")

    config = {}
    if not os.path.isfile(config_file_path):
        print("Файл конфигурации не найден.")
        config = create_config_file()
    else:
        with open(config_file_path, "r", encoding="utf-8") as config_file:
            config = json.load(config_file)

    if not os.path.isfile(query_file_path):
        print("Файл query.sql не найден.")
        create_query_file(query_file_path)

    with open(query_file_path, "r", encoding="utf-8") as query_file:
        sql_query = query_file.read()

    fdb_files = find_fdb_files(folder_path)
    database_list = create_database_list(fdb_files, config)

    if os.path.isfile("database.json"):
        overwrite = input("Файл database.json уже существует. Хотите перезаписать его? (y/n): ")
        if overwrite.lower() == "y":
            save_to_json(database_list)
            print("Список баз данных сохранен в файле database.json.")
        else:
            print("Продолжаем работу с существующим файлом database.json.")
    else:
        save_to_json(database_list)
        print("Список баз данных сохранен в файле database.json.")

    current_date = datetime.now().strftime("%d-%m-%Y")
    os.makedirs(report_dir, exist_ok=True)

    try:
        proceed = input("Считаем коров? (y/n): ")
        if proceed.lower() == "y":
            for database in database_list:
                print("Запуск SQL-запроса...")
                database_name = os.path.splitext(os.path.basename(database["database_path"]))[0]
                file_name = f"{database_name}_{datetime.now().strftime('%H-%M-%S')}.xlsx"

                database_path = os.path.dirname(database["database_path"])
                file_path = os.path.join(appdata_temp_dir, os.path.basename(database["database_path"]))

                # Копирование файла базы данных в папку TempCows
                shutil.copy2(database["database_path"], file_path)

                conn_str = {
                    "dsn": f'{database["hostname"]}:{file_path}',
                    "user": database["username"],
                    "password": database["password"],
                    "no_db_triggers": True
                }

                result_set = execute_sql_query(conn_str, sql_query)
                save_to_excel(result_set, file_path)

                # Получение относительного пути базы данных относительно исходной папки
                relative_db_path = os.path.relpath(database_path, folder_path)

                # Построение пути для сохранения файла отчета с сохранением структуры папок
                report_subdir = os.path.join(report_dir, relative_db_path)
                os.makedirs(report_subdir, exist_ok=True)
                report_file_path = os.path.join(report_subdir, file_name)
                shutil.move(file_path, report_file_path)

                print(f"Результаты SQL-запроса сохранены в файле: {report_file_path}")
        else:
            print("Программа завершена.")
    finally:
        # Удаление папки TempCows
        delete_temp_cows_folder(appdata_temp_dir)


if __name__ == "__main__":
    main()
