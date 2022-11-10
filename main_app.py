import shutil

from add_candidates_db import candidates_add
from add_registry_db import add_to_registry
from remove_folders import move_folder

from class_constant import MAIN_FILE, CONNECT, DESTINATION


# главный модуль программы
if __name__ == "__main__":
    # Создаем резервную копию книги Exel в DESTINATION
    shutil.copy(MAIN_FILE, DESTINATION)
    # Создаем резервную копию БД в DESTINATION
    shutil.copy(CONNECT, DESTINATION)

    # вызываем функцию для добавления новых записей в таблицу проверки кандидатов БД
    candidates_add()

    # вызываем функцию для добавления новых записей в реестр кандидатов БД
    add_to_registry()

    # вызываем функцию для переноса папок
    move_folder()
