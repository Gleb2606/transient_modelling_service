from flask import Flask, request, jsonify
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import numpy as np
import tensorflow as tf
import pythoncom
import rustab_interaction
from config import Config
from flask_cors import CORS

app = Flask(__name__)
CORS(app)
app.config.from_object(Config)

# Инициализация объекта SQLAlchemy
db = SQLAlchemy()

# Связываем объект SQLAlchemy с приложением
db.init_app(app)

# Модель SQLAlchemy для таблицы Models
class Model(db.Model):
    __tablename__ = 'model'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    input_width = db.Column(db.Integer, nullable=False)
    data_in_frame = db.Column(db.Integer)

# Загрузка модели
model = tf.keras.models.load_model('model.h5')

@app.route('/get-transient', methods=['POST'])
def get_transient():
    """
    Функция представление для старта расчета динамики
    :return: JSON типа {'message': 'Model trained successfully',
                        'path': Путь к файлу результата}
    """
    try:
        pythoncom.CoInitialize()

        # Получение файла режима
        rst_file = request.json['rst_file']

        # Получение файла сценария
        scn_file = request.json['scn_file']

        # Получение времени моделирования
        input_width = request.json['input_width']

        # Шаблоны файлов
        shablon_scn = "C:\\Users\\Umaro\\OneDrive\\Документы\\RastrWin3\\SHABLON\\сценарий.scn"
        shablon_dfw = "C:\\Users\\Umaro\\OneDrive\\Документы\\RastrWin3\\SHABLON\\автоматика.dfw"
        shablon_kpr = "C:\\Users\\Umaro\\OneDrive\\Документы\\RastrWin3\\SHABLON\\контр-е величины.kpr"

        # Хардкод
        name = "Богучанская ГЭС"
        node = 60533014
        data_set = "Num=60533014"
        name_support = "Красноярская ГЭС"
        data_set_support = "Num=60522003"

        try:
            rustab_interaction.file_prepare(rst_file, scn_file, shablon_dfw, shablon_scn, shablon_kpr,
                                            name, data_set, name_support, data_set_support)

            rustab_interaction.calculate_dynamic(input_width, 0.01, 0.001,
                                                 0.5, 0.01)

            save_path = rustab_interaction.get_transient("Generator", "Delta",
                                                         node, "БоГЭС", 1)

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

            rustab_interaction.preprocessing(save_path, save_path, boges_generators, rst_file)

        except Exception as ex:
            raise Exception(str(ex))

        return jsonify({'message': 'Model trained successfully',
                        'path': save_path}), 200

    finally:
        pythoncom.CoUninitialize()

@app.route('/predict', methods=['POST'])
def predict():
    """
    Функция представление для прогноза оставшейся части переходного процесса
    :return: JSON типа {"result": результат прогноза, "path": путь к файлу результата})
    """
    # Получение данных из запроса
    path = request.json['path']
    name = request.json['name']

    data = rustab_interaction.parse_csv_to_array(path)

    # Получение значения из БД
    model_params = Model.query.filter_by(name=name).first()

    # Если первичный ключ не найден
    if not model_params:
        return jsonify({'error': 'No model with this id exists'}), 400

    # Количество точек в датафрейме
    t = model_params.data_in_frame

    # Входное окно
    input_t = model_params.input_width

    # Преобразование списка в датафрейм
    df = pd.DataFrame(data)

    # Формируем слайс из первых input_t точек
    test_csv_pred = df.iloc[0:input_t].to_numpy()
    test_csv_pred = np.expand_dims(test_csv_pred, axis=0)

    # Выполняем прогноз
    predictions = model.predict(test_csv_pred)

    # Преобразуем данные прогноза
    predictions_df = predictions.reshape(t - input_t, 4)

    # Создаем датафрейм с предсказанными данными
    predictions_df_ar = pd.DataFrame(data=predictions_df, columns=['delta', 'w', 'a', 'p'])

    # Денормализация данных
    for column in predictions_df_ar.columns:
        min_value = predictions_df_ar[column].min()
        max_value = predictions_df_ar[column].max()
        predictions_df_ar[column] = predictions_df_ar[column] * (max_value - min_value) + min_value

    # Преобразуем датафрейм в список списков
    result = predictions_df_ar.to_numpy().tolist()

    # Объединение прогноза и входных данных
    #rustab_interaction.add_array_to_csv(result, path)

    # Возвращаем результат как JSON
    return jsonify({"result": result, "path": path}), 200

if __name__ == '__main__':
    app.run(debug=True, port=5007)