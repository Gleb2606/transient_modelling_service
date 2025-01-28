import csv
import locale
from typing import List, Union
import win32com.client
import pandas as pd
import numpy as np

rastr = win32com.client.Dispatch("Astra.Rastr")


class PowerPlantAnalyzer:
    """
    Класс, описывающий станцию
    """
    RASTR = win32com.client.Dispatch("Astra.Rastr")

    def __init__(self, power_plant, mode):
        """
        Инициализация класса.

        :param power_plant: Словарь, где ключи — названия генераторов, значения — идентификаторы.
        :param mode: Строка, определяющая режим загрузки.
        """
        self.power_plant = power_plant
        self.mode = mode

    def calculate_initial_power(self):
        """
        Метод для расчета начальной мощности всех генераторов.

        :return: Сумма начальных мощностей генераторов.
        """

        initial_power = 0

        for generator in self.power_plant.values():
            self.RASTR.Load(1, self.mode, "")
            index = 0
            table = self.RASTR.Tables("Generator")
            column_item = table.Cols("Num")

            for i in range(table.Count):
                if column_item.Z(i) == generator:
                    index = i
                    break

            column_item = table.Cols("P")
            generator_power = column_item.Z(index)
            initial_power += generator_power

        return initial_power

    def get_generator_power(self):
        """
        Метод для получения мощности единицы генерирующего оборудования.

        :return: Сумма начальных мощностей генераторов.
        """

        power = 0

        for generator in self.power_plant.values():
            self.RASTR.Load(1, self.mode, "")
            index = 0
            table = self.RASTR.Tables("Generator")
            column_item = table.Cols("Num")

            for i in range(table.Count):
                if column_item.Z(i) == generator:
                    index = i
                    break

            column_item = table.Cols("P")
            generator_power = column_item.Z(index)
            power += generator_power
            break

        return power

def load_file(file_path: str, shablon: str) -> None:
    """
    Функция загрузки файла в рабочую область
    :param file_path: Путь до файла
    :param shablon: Шаблон файла
    :return: None
    """
    rastr.Load(1, file_path, shablon)

def save_file(file_name: str, shablon: str) -> None:
    """
    Функция сохранения файла
    :param file_name: Путь до файла
    :param shablon: Шаблон файла
    :return: None
    """
    try:
        rastr.Save(file_name, shablon)
    except Exception as ex:
        raise Exception(str(ex))

def create_file(shablon: str) -> None:
    """
    Функция создания файла
    :param shablon: Шаблон файла
    :return: None
    """
    rastr.NewFile(shablon)

def get_index_by_number(table_name: str, parameter_name: str, number: int) -> int:
    """
    Функция поиска индекса из таблицы по номеру
    :param table_name: Название таблицы
    :param parameter_name: Наименование параметра
    :param number: Номер узла
    :return: Индекс
    """
    table = rastr.Tables.Item(table_name)
    column_item = table.Cols.Item(parameter_name)

    for index in range(0,table.Count,1):
        if column_item.Z(index) == number:
            return index

    raise Exception(f"Элемент с номером {number} не найден")

def get_index_by_value(table_name: str, parameter_name: str, value: str) -> int:
    """
    Функция поиск индекса из таблицы по номеру
    :param table_name: Наименование таблицы
    :param parameter_name: Наименование параметра
    :param value: Параметр узла
    :return: Индекс
    """
    table = rastr.Tables.Item(table_name)
    column_item = table.Cols.Item(parameter_name)

    for index in range(0,table.Count,1):
        if column_item.Z(index) == value:
            return index

    raise Exception(f"Элемент {value} не найден")

def set_value(table_name: str, parameter_name: str, number: int,
                   chosen_parameter: str, value: any) -> None:
    """
    Функция задания логического значения
    :param table_name: Наименование таблицы
    :param parameter_name: Наименование параметра
    :param number: Номер узла
    :param chosen_parameter: Выбранный параметр
    :param value: Новое значение
    :return: None
    """
    table = rastr.Tables.Item(table_name)
    column_item = table.Cols.Item(chosen_parameter)

    index = get_index_by_number(table_name, parameter_name, number)
    column_item.SetZ(index, value)

def change_branch_state(ip_number: int, iq_number: int, np_number: int,
                        state: bool) -> None:
    """
    Функция коммутации ветви
    :param ip_number: Номер начала ветви
    :param iq_number: Номер конца ветви
    :param np_number: Номер параллельности ветви
    :param state: Состояние ветви
    :return: None
    """
    table = rastr.Tables.Item("vetv")

    ip_column_item = table.Cols.Item("ip")
    iq_column_item = table.Cols.Item("iq")
    np_column_item = table.Cols.Item("np")
    sta_column_item = table.Cols.Item("sta")

    for index in range(0,table.Count,1):
        if (ip_column_item.get_ZN(index) == ip_number and
            iq_column_item.get_ZN(index) == iq_number):

            if np_number == 0:
                sta_column_item.set_ZN(index, state)
                break

            elif np_column_item.get_ZN(index) == np_number:
                sta_column_item.set_ZN(index, state)
                break

def get_generator_list(is_plant_researched: bool) -> list:
    """
    Функция получения списка генераторов станции
    :param is_plant_researched: Исследуемая ли станция
    :return: Список генераторов станции
    """
    table = rastr.Tables.Item("ut_node")

    ny_column_item = table.Cols.Item("ny")
    pg_column_item = table.Cols.Item("pg")

    generator_list = []

    for index in range(0,table.Count,1):
        if (pg_column_item.get_ZN(index) > 0 and
            is_plant_researched):

            generator_list.append(ny_column_item.get_ZN(index))

        elif (pg_column_item.get_ZN(index) < 0 and
              is_plant_researched == False):

            generator_list.append(ny_column_item.get_ZN(index))

    return generator_list

def get_value (table_name: str, parameter_name: str, number: int,
               chosen_parameter: str) -> any:
    """
    Функция получения значения
    :param table_name: Наименование таблицы
    :param parameter_name: Наименование параметра
    :param number: Номер узла
    :param chosen_parameter: Выбранный параметр
    :return: Искомый параметр
    """
    table = rastr.Tables(table_name)
    column_item = table.Cols(chosen_parameter)

    index = get_index_by_number(table_name, parameter_name, number)
    return column_item.get_ZN(index)

def is_regime_ok () -> bool:
    """
    Функция проверки существования установившегося режима
    :return: Статус режима
    """
    status = rastr.rgm("")

    if status == 0:
        return True

    else:
        return False

def regime() -> None:
    """
    Функция расчета установившегося режима
    :return: None
    """
    rastr.rgm("")

def fill_numbers_list(table_name: str, parameter_name: str) -> list:
    """
    Функция заполнения списка номерами узлов
    :param table_name: Наименование таблицы
    :param parameter_name: Наименование параметра
    :return: Список значений
    """
    list_of_numbers = []

    try:
        table = rastr.Tables.Item(table_name)
        column_item = table.Cols.Item(parameter_name)

        for index in range(0,table.Count,1):
            list_of_numbers.append(column_item.get_ZN(index))

        return list_of_numbers

    except Exception as ex:
        raise Exception(str(ex))

def step_back() -> None:
    """
    Функция для выполнения шага назад по траектории утяжеления
    :return: None
    """
    table = rastr.Tables.Item("ut_common")
    column_item = table.Cols.Item("kfc")

    index = table.Count - 1
    step = column_item.get_ZN(index)

    column_item.set_ZN(index, -step)

    kd = rastr.step_ut("z")

    if ((kd == 0 and
       ((rastr.ut_Param[rastr.ParamUt.UT_ADD_P] == 0) or
       (rastr.ut_Param[rastr.ParamUt.UT_ADD_P] == 1)))):

        rastr.AddControl(-1, "")

    column_item.SetZN(index, step)

def add_row(table_name: str, index: int, number: int) -> None:
    """
    Функция добавления новой строки
    :param table_name: Наименование таблицы
    :param index: Идентификатор в таблице
    :param number: Номер в таблице
    :return: None
    """
    table = rastr.Tables(table_name)
    table.AddRow()
    rastr.Tables(table_name).Cols("id").SetZ(index, number)
    #column_item = table.Cols("id")
    #column_item.setZ(index, number)

def add_kpr(table_name: str, index: int, number: int) -> None:
    """
    Функция добавления контролируемых параметров
    :param table_name: Наименование таблицы
    :param index: Идентификатор в таблице
    :param number: Номер в таблице
    :return: None
    """
    table = rastr.Tables.Item(table_name)
    table.AddRow()
    rastr.Tables(table_name).Cols("Num").SetZ(index, number)
    #column_item = table.Cols.Item("Num")
    #column_item.setZ(index, number)

def add_action_row(table: str, index: int, parent_id: int, action_type: any,
                   formula: any, object_key: any, output_mode=0, runs_count=1) -> None:
    """
    Функция формирования действий сценария
    :param table: Таблица
    :param index: Индекс
    :param parent_id: Родительский объект
    :param action_type: Тип действия
    :param formula: Формула
    :param object_key: Ключ объекта
    :param output_mode: Выходное состояние
    :param runs_count: Количество пробегов
    :return: None
    """
    add_row(table, index - 1, index)
    set_value(table, "Id", index, "ParentId", parent_id)
    set_value(table, "Id", index, "Type", action_type)
    set_value(table, "Id", index, "Formula", formula)
    set_value(table, "Id", index, "ObjectKey", object_key)
    set_value(table, "Id", index, "OutputMode", output_mode)
    set_value(table, "Id", index, "RunsCount", runs_count)

def add_logic_row(table: str, index: int, actions: any, delay: float,
                  formula=1, action_type=1, output_mode=0) -> None:
    """
    Функция формирования логики сценария
    :param table: Таблица
    :param index: Идентификатор
    :param actions: Действия
    :param delay: Выдержка времени
    :param formula: Формула
    :param action_type: Тип действия
    :param output_mode: Выходное состояние
    :return: None
    """
    add_row(table, index - 1, index)
    set_value(table, "Id", index, "Formula", formula)
    set_value(table, "Id", index, "Type", action_type)
    set_value(table, "Id", index, "Actions", actions)
    set_value(table, "Id", index, "Delay", delay)
    set_value(table, "Id", index, "OutputMode", output_mode)

def make_scn_1 (shunt: float, shunt_recloser: float, protection_delay: float,
                recloser_delay: float, fault_node: int, new_fault_node: int,
                line_1: str, line_2: str, line_3: str, line_4: str) -> None:
    """
    Формирование сценариев первой и второй группы
    :param shunt: Шунт короткого замыкания
    :param shunt_recloser: Шунт после АПВ
    :param protection_delay: Выдержка времени релейной защиты
    :param recloser_delay: Выдержка времени АПВ
    :param fault_node: Узел короткого замыкания
    :param new_fault_node: Узел после АПВ
    :param line_1: Линия 1
    :param line_2: Линия 2
    :param line_3: Линия 3
    :param line_4: Линия 4
    :return: None
    """

    # Формирование действий сценария
    add_action_row("DFWAutoActionScn", 1, 1, 6, shunt, fault_node)
    add_action_row("DFWAutoActionScn", 2, 2, 3, 0, line_1)
    add_action_row("DFWAutoActionScn", 3, 2, 3, 0, line_2)
    add_action_row("DFWAutoActionScn", 4, 3, 5, 0, new_fault_node)
    add_action_row("DFWAutoActionScn", 5, 4, 6, shunt_recloser, new_fault_node)
    add_action_row("DFWAutoActionScn", 6, 5, 3, 0, line_3)
    add_action_row("DFWAutoActionScn", 7, 5, 3, 0, line_4)
    add_action_row("DFWAutoActionScn", 8, 6, 3, 1, line_3)
    add_action_row("DFWAutoActionScn", 9, 6, 3, 1, line_4)
    add_action_row("DFWAutoActionScn", 10, 7, 3, 0, line_3)
    add_action_row("DFWAutoActionScn", 11, 7, 3, 0, line_4)

    # Формирование логики сценария
    add_logic_row("DFWAutoLogicScn", 1, "A1", 0)
    add_logic_row("DFWAutoLogicScn", 2, "A2", protection_delay)
    add_logic_row("DFWAutoLogicScn", 3, "A3", protection_delay)
    add_logic_row("DFWAutoLogicScn", 4, "A4", protection_delay)
    add_logic_row("DFWAutoLogicScn", 5, "A5",
                  protection_delay + 0.05)
    add_logic_row("DFWAutoLogicScn", 6, "A6",
                  protection_delay + 0.05 + recloser_delay)
    add_logic_row("DFWAutoLogicScn", 7, "A7",
                  2 * protection_delay + 0.05 + recloser_delay)

def add_logic_row_3(table: str, index: int, formula: int, logic_type: any,
                    actions: any, delay: float, output_mode=0) -> None:
    """
    Формирование сценариев 3 группы
    :param table: Таблицы
    :param index: Индекс
    :param formula: Формула
    :param logic_type: Тип логики
    :param actions: Действия
    :param delay: Выдержка времени
    :param output_mode: Выходное состояние
    :return: None
    """
    add_row(table, index - 1, index)
    set_value(table, "Id", index, "Formula", formula)
    set_value(table, "Id", index, "Type", logic_type)
    set_value(table, "Id", index, "Actions", actions)
    set_value(table, "Id", index, "Delay", delay)
    set_value(table, "Id", index, "OutputMode", output_mode)

def make_scn_3(shunt: float, protection_delay: float,
               cbfp_delay: float, fault_node: int, line_1: str,
               line_2: str, line_3: str, line_4: str, line_5: str) -> None:
    """
    Функция формирования сценариев третьей группы
    :param shunt: Шунт короткого замыкания
    :param protection_delay: Выдержка времени релейной защиты
    :param cbfp_delay: Выдержка времени УРОВ
    :param fault_node: Узел короткого замыкания
    :param line_1: Линия 1
    :param line_2: Линия 2
    :param line_3: Линия 3
    :param line_4: Линия 4
    :param line_5: Линия 5
    :return: None
    """

    # Формирование действий сценария
    add_action_row("DFWAutoActionScn", 1, 1, 6, shunt, fault_node)
    add_action_row("DFWAutoActionScn", 2, 2, 3, 0, line_1)
    add_action_row("DFWAutoActionScn", 3, 3, 3, 0, line_2)
    add_action_row("DFWAutoActionScn", 4, 3, 3, 0, line_3)
    add_action_row("DFWAutoActionScn", 5, 4, 3, 0, line_4)
    add_action_row("DFWAutoActionScn", 6, 4, 3, 0, line_5)

    # Формирование логики сценария
    add_logic_row_3("DFWAutoLogicScn", 1, 1, 1, "A1", 0)
    add_logic_row_3("DFWAutoLogicScn", 2, 1, 1, "A2",
                    protection_delay)
    add_logic_row_3("DFWAutoLogicScn", 3, 1, 1, "A3",
                    protection_delay + 0.05)
    add_logic_row_3("DFWAutoLogicScn", 4, 1, 1, "A4",
                    protection_delay + 0.05 + cbfp_delay)

def calculate_dynamic(input_width: float, start_step: float, min_step: float,
                      max_step: float, out_step: float) -> None:
    """
    Функция запуска расчета переходного процесса
    :param input_width: Величина окна наблюдения
    :param start_step: Начальный шаг интегрирования
    :param min_step: Минимальный шаг интегрирования
    :param max_step: Максимальный шаг интегрирования
    :param out_step: Шаг печати
    :return: None
    """
    try:
        dynamic = rastr.FWDynamic()

        table = rastr.Tables.Item("com_dynamics")

        column_item_tras = table.Cols.Item("Tras")
        column_item_tras.SetZ(0, input_width)

        column_item_hint = table.Cols.Item("Hint")
        column_item_hint.SetZ(0, start_step)

        column_item_hmin = table.Cols.Item("Hmin")
        column_item_hmin.SetZ(0, min_step)

        column_item_hmax = table.Cols.Item("Hmax")
        column_item_hmax.SetZ(0, max_step)

        column_item_hout = table.Cols.Item("Hout")
        column_item_hout.SetZ(0, out_step)

        column_item_period_angle = table.Cols.Item("PeriodAngle")
        column_item_period_angle.SetZ(0, 0)

        dynamic.Run()

    except Exception as ex:
        raise Exception(str(ex))

def file_prepare(rst_file: str, scn_file: str, dfw_file: str, shablon_scn: str,
                 kpr_file: str, name: str, data_set: str,
                 name_support: str, set_support: str) -> None:
    """
    Функция подготовки файлов для расчета динамики
    :param rst_file: Путь к файлу установившегося режима
    :param scn_file: Путь к файлу сценария
    :param dfw_file: Путь к файлу автоматики
    :param shablon_scn: Шаблон сценария
    :param kpr_file: Путь к файлу контролируемых величин
    :param name: Наименование исследуемого генератора
    :param data_set: Выборка для исследуемого генератора
    :param name_support: Наименование опорного генератора
    :param set_support: Выборка для опорного генератора
    :return: None
    """
    rastr.Load(1, rst_file, "")
    rastr.Load(1, scn_file, shablon_scn)
    rastr.Load(1, dfw_file, dfw_file)
    rastr.Load(1, kpr_file, kpr_file)
    create_kpr(name, data_set, name_support, set_support)

def configure_kpr_entry(num: int, entry_name: str, entry_set: str) -> None:
    """
    Функция конфигурации файлов контролируемых величин
    :param num: Номер величины
    :param entry_name: Наименование величины
    :param entry_set: Выборка для величины
    :return: None
    """
    set_value("ots_val", "Num", num, "name", entry_name)
    set_value("ots_val", "Num", num, "tip", 0)
    set_value("ots_val", "Num", num, "tabl", "Generator")
    set_value("ots_val", "Num", num, "vibork", entry_set)
    set_value("ots_val", "Num", num, "formula", "Delta")
    set_value("ots_val", "Num", num, "prec", 2)
    set_value("ots_val", "Num", num, "mash", 57)

def create_kpr(name: str, data_set: str, name_support: str, set_support: str) -> None:
    """
    Функция создания файлов контролируемых величин
    :param name: Наименование исследуемого генератора
    :param data_set: Выборка для исследуемого генератора
    :param name_support: Наименование опорного генератора
    :param set_support: Выборка для опорного генератора
    :return: None
    """
    add_kpr("ots_val", 0,1)
    configure_kpr_entry(1, name, data_set)

    add_kpr("ots_val", 1, 2)
    configure_kpr_entry(2, name_support, set_support)

def get_transient(name: str, parameter: str, key: int, plant: str, index: int) -> str:
    """
    Функция сохранения результатов в CSV
    :param name: Наименование таблицы
    :param parameter: Наименование параметра
    :param key: Номер генератора
    :param plant: Наименование станции
    :param index: Порядковый номер результата
    :return: Наименование файла с результатом
    """
    gen_id = get_index_by_number(name, "Num", key)
    plot = rastr.GetChainedGraphSnapshot(name, parameter, gen_id, 0)

    save_path = f"C:\\Users\\Umaro\\OneDrive\\Рабочий стол\\res\\Python_{plant}_{index}.csv"

    #Установка региональных настроек для форматирования чисел
    locale.setlocale(locale.LC_NUMERIC, "ru_RU")

    # Запись CSV
    with open(save_path, "w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file, delimiter=';')

        # Запись заголовков
        writer.writerow(["delta", "t"])

        # Запись данных
        for row in plot:
            formatted_row = [f"{value:.6f}".replace('.', ',') for value in row]
            writer.writerow(formatted_row)

    return save_path

def preprocessing(input_path: str, output_path: str, generators: dict, rst_file: str) -> None:
    """
    Функция предобработки данных
    :param input_path: Исходный файл
    :param output_path: Выходной файл
    :param generators: Словарь генераторов
    :param rst_file: Файл режима
    :return: None
    """
    locale.setlocale(locale.LC_NUMERIC, "ru_RU")

    # Считываем данные
    data = pd.read_csv(input_path, delimiter=';', encoding='utf-8')

    # Заменяем запятую на точку и преобразуем в float
    data['delta'] = data['delta'].str.replace(',', '.').astype(float)

    # Вычисляем первую производную
    data['w'] = np.gradient(data['delta'].values)

    # Вычисляем вторую производную
    data['a'] = np.gradient(data['w'].values)

    plant = PowerPlantAnalyzer(generators, rst_file)
    power = plant.calculate_initial_power()
    gen_power = plant.get_generator_power()

    # Ввод управляющего воздействия
    n = len(data)
    data['p'] = 2700
    data.loc[36:n, 'p'] = 2000

    # Удаляем столбец 't'
    if 't' in data.columns:
        data.drop(columns=['t'], inplace=True)

    # Нормализация столбцов
    for column in ['delta', 'w', 'a', 'p']:
        data[column] = (data[column] - data[column].min()) / (data[column].max() - data[column].min())

    # Сохранение нормализованных данных
    data.to_csv(output_path, sep=';', encoding='utf-8', index=False, float_format='%.6f')

def parse_csv_to_array(file_path: str) -> List[List[Union[float, int]]]:
    """
    Функция преобразования файла csv в массив переменных
    :param file_path: Путь к файлу переходного процесса
    :return: Многомерный массив переходного процесса
    """
    data = []

    try:
        with open(file_path, "r", encoding="utf-8") as file:
            reader = csv.reader(file, delimiter=';')

            # Пропускаем заголовок
            next(reader, None)

            for row in reader:
                data.append([float(value.replace(',', '.')) for value in row])

    except FileNotFoundError:
        print(f"Файла {file_path} не существует")

    except ValueError as ve:
        print(f"Ошибка преобразование {ve}")

    except Exception as e:
        print(f"Ошибка {e}")

    return data

def add_array_to_csv(array: list, file_path: str) -> None:
    """
    Функция объединения массива данных окна наблюдения и окна прогноза
    :param array: Данные в окне прогноза
    :param file_path: Данные в окне наблюдения
    :return: None
    """
    try:
        # Открываем файл в режиме "a" (append) для добавления данных
        with open(file_path, mode='a', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file, delimiter=';')

            # Пишем каждую строку из массива в файл
            writer.writerows(array)

        print(f"Данные успешно добавлены в файл: {file_path}")

    except Exception as e:
        print(f"Произошла ошибка при добавлении данных в файл: {e}")