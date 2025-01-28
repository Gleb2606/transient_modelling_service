import win32com
import rustab_interaction

if __name__ == '__main__':
    rst_file = "C:\\Users\\Umaro\\OneDrive\\Рабочий стол\\rst\\БоГЭС_Потребление БоАЗ_277.rst"

    boges_generators = {
        "Богучанская ГЭС - Г1": 60533008,
        "Богучанская ГЭС - Г2": 60533009,
        "Богучанская ГЭС - Г3": 60533010,
        "Богучанская ГЭС - Г4": 60533011,
        "Богучанская ГЭС - Г5": 60533012,
        "Богучанская ГЭС - Г6": 60533013,
        "Богучанская ГЭС - Г7": 60533014,
        "Богучанская ГЭС - Г8": 60533015,
        "Богучанская ГЭС - Г9": 60533016,
    }

    boges = rustab_interaction.PowerPlantAnalyzer(boges_generators, rst_file)
    power = boges.calculate_initial_power()
    gen_power = boges.get_generator_power()

    print(power)
    print(gen_power)
    # Получение файла режима
    #rst_file = "C:\\Users\\Umaro\\OneDrive\\Рабочий стол\\rst\\БоГЭС_Потребление БоАЗ_277.rst"

    # Получение файла сценария
    #scn_file = "C:\\Users\\Umaro\\OneDrive\\Рабочий стол\\scn\\сценарий_УРОВ_БоГЭС - Озёрная_3.scn"

    # Шаблоны файлов
    #shablon_scn = "C:\\Users\\Umaro\\OneDrive\\Документы\\RastrWin3\\SHABLON\\сценарий.scn"
    #shablon_dfw = "C:\\Users\\Umaro\\OneDrive\\Документы\\RastrWin3\\SHABLON\\автоматика.dfw"
    #shablon_kpr = "C:\\Users\\Umaro\\OneDrive\\Документы\\RastrWin3\\SHABLON\\контр-е величины.kpr"

    # Хардкод
    #name = "Богучанская ГЭС"
    #node = 60533014
    #data_set = "Num=60533014"
    #name_support = "Красноярская ГЭС"
    #data_set_support = "Num=60522003"

    #try:
        #rustab_interaction.file_prepare(rst_file, scn_file, shablon_dfw, shablon_scn, shablon_kpr,
                                        #name, data_set, name_support, data_set_support)

        #print("Начался расчёт динамики")
        #rustab_interaction.calculate_dynamic(1.41, 0.01, 0.001, 0.5, 0.01)
        #print("Расчёт динамики завершен")
        #print("Запись результата ...")
        #rustab_interaction.get_transient("Generator", "Delta", node, "БоГЭС", 123)
        #print("предобработка данных...")
        #rustab_interaction.preprocessing("C:\\Users\\Umaro\\OneDrive\\Рабочий стол\\res\\Python_БоГЭС_123.csv",
                                         #"C:\\Users\\Umaro\\OneDrive\\Рабочий стол\\res\\Python_БоГЭС_123.csv")

        #path_to_csv = "C:\\Users\\Umaro\\OneDrive\\Рабочий стол\\res\\Python_БоГЭС_1.csv"

        #array = rustab_interaction.parse_csv_to_array(path_to_csv)

        #print(array)

    #except Exception as ex:
        #raise Exception(str(ex))