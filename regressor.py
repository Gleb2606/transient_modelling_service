# Импорт необходимых библиотек
import pandas as pd
import numpy as np
import tensorflow as tf

# Загрузка модели
model = tf.keras.models.load_model('model.h5')

data = pd.read_csv('test.csv', sep=';')

# Количество точек в датафрейме
t = 144

# Входное окно
input_t = 40

# Формируем слайс из первых 500 мс переходного процесса
test_csv_pred = data[0:input_t]

test_csv_pred = test_csv_pred.to_numpy()
test_csv_pred = np.expand_dims(test_csv_pred, axis=0)

print(test_csv_pred)

# Выполняем прогноз с исользованием нашей обученной нейронной сети
predictions = model.predict(test_csv_pred)

# Данные прогноза преобразуем в формату, для дальнейшего создания датафрейма
predictions_df = predictions.reshape(t - input_t, 4)

# Создаем датафрейм с предсказанными данными
predictions_DF_ar = pd.DataFrame(data=predictions_df, columns=['delta', 'w', 'a', 'p'])

print(predictions_DF_ar)