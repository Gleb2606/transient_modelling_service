from flask import Flask, request, jsonify
import pandas as pd
import numpy as np
import tensorflow as tf

app = Flask(__name__)

# Загрузка модели
model = tf.keras.models.load_model('model.h5')


@app.route('/predict', methods=['POST'])
def predict():
    # Получение данных из запроса
    data = request.json['data']

    # Преобразование списка в датафрейм
    df = pd.DataFrame(data)

    # Количество точек в датафрейме
    t = 144
    # Входное окно
    input_t = 40

    # Формируем слайс из первых input_t точек
    test_csv_pred = df.iloc[0:input_t].to_numpy()
    test_csv_pred = np.expand_dims(test_csv_pred, axis=0)

    # Выполняем прогноз
    predictions = model.predict(test_csv_pred)

    # Преобразуем данные прогноза
    predictions_df = predictions.reshape(t - input_t, 4)

    # Создаем датафрейм с предсказанными данными
    predictions_DF_ar = pd.DataFrame(data=predictions_df, columns=['delta', 'w', 'a', 'p'])

    # Денормализация данных
    for column in predictions_DF_ar.columns:

        min_value = predictions_DF_ar[column].min()
        max_value = predictions_DF_ar[column].max()

        predictions_DF_ar[column] = predictions_DF_ar[column] * (max_value - min_value) + min_value

    # Возвращаем результат как JSON
    return jsonify(predictions_DF_ar.to_dict(orient='records'))


if __name__ == '__main__':
    app.run(debug=True)