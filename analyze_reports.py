import pandas as pd
import os, re

from datetime import date
from datetime import timedelta

# Настройки путей и дат
reports_path = os.path.join(os.path.abspath(os.getcwd()), 'reports')                # путь до общей папки с отчётами

# Контроль периода
first_date = date.today() - timedelta(days=date.today().weekday())                  # начало недели
yesterday_date = date.today() - timedelta(days=1)                                   # вчера
today_date = date.today()                                                           # сегодня
# Если сегодня понедельник, то берем всю прошлую неделю
if date.today() == (date.today() - timedelta(days=date.today().weekday())):
    first_date = date.today() - timedelta(days=date.today().weekday()) - timedelta(days=7) # начало прошлой недели

# Функция для сохранения датафрейма в Excel с автоподбором ширины столбца
def save_to_excel(dframe: pd.DataFrame, path, index_arg=False):
    with pd.ExcelWriter(path, mode='w', engine='openpyxl') as writer:
         dframe.to_excel(writer, index=index_arg)
         for column in dframe:
            column_width = max(dframe[column].astype(str).map(len).max(), len(column))
            col_idx = dframe.columns.get_loc(column)
            writer.sheets['Sheet1'].column_dimensions[chr(65+col_idx)].width = column_width + 5

#  1. Выгружаем отчет http://bi.mz.mosreg.ru/#form/disp_tmk за контролируемый период (с начала текущей недели или для понедельника - вся прошлая неделя).
df_disp_tmk = pd.read_excel(reports_path + '\\Количество карт ДВН и УДВН закрытых через ТМК.xlsx', skiprows=1, header=0)
# Выбираем ПОКБ
df_disp_tmk = df_disp_tmk.loc[df_disp_tmk['ОГРН медицинской организации'] == 1215000036305]
# Сконвертировать время закрытия карты в дату
df_disp_tmk['Закрытие диспансеризации через телемедицинские консультации'] = pd.to_datetime(df_disp_tmk['Закрытие диспансеризации через телемедицинские консультации'], format='%d.%m.%Y %H:%M:%S').dt.date
# Оставляем одну фамилию
df_disp_tmk['ФИО пациента'] = df_disp_tmk['ФИО пациента'].apply(lambda x: x.split(' ')[0])

# 2. Выгружаем отчет http://bi.mz.mosreg.ru/#form/pass_dvn за период "Текущая дата минус месяц"
df_pass_dvn = pd.read_excel(reports_path + '\\Прохождение пациентами ДВН или ПМО.xlsx', skiprows=1, header=0)

# Сконвертировать время закрытия карты в дату
df_pass_dvn['Дата закрытия карты диспансеризации'] = pd.to_datetime(df_pass_dvn['Дата закрытия карты диспансеризации'], format='%d.%m.%Y %H:%M:%S').dt.date

# 2.1 Причина закрытия - Обследование пройдено
# 2.2 Дата закрытия - контролируемый период
# 2.3 Вид обследования - 404 Диспансеризация и 404 Профилактические медицинские осмотры
df_pass_dvn = df_pass_dvn[(df_pass_dvn['Причина закрытия'] == 'Обследование пройдено') & \
                          (df_pass_dvn['Дата закрытия карты диспансеризации'] >= first_date) & \
                          (df_pass_dvn['Дата закрытия карты диспансеризации'] <= today_date) & \
                          (df_pass_dvn['Вид обследования'] == '404н Диспансеризация')]

# 3. Из (2) убираем записи, где "Результат обращения" содержит "Направлен на II этап"
# df_pass_dvn = df_pass_dvn.loc[~df_pass_dvn['Результат обращения'].str.contains('Направлен на II этап', na=False)]

df_disp_tmk = df_disp_tmk.rename(columns={'Закрытие диспансеризации через телемедицинские консультации':'Дата закрытия диспансеризации'})
df_pass_dvn = df_pass_dvn.rename(columns={'Дата закрытия карты диспансеризации':'Дата закрытия диспансеризации', 'Врач подписывающий заключение диспансеризации':'Врач'})

# 4. К (2) подгружаем (1) по ключу "Фамилия + Дата закрытия карты"
df_final = df_pass_dvn \
    .merge(df_disp_tmk,  
           left_on=['ФИО пациента', 'Дата закрытия диспансеризации'],  
           right_on=['ФИО пациента', 'Дата закрытия диспансеризации'], 
           how='left',
           indicator=True) \
    .query('_merge == "left_only"') \
    .drop(['_merge', '#_x', '#_y', 'Медицинская организация диспансеризации', 'ОГРН', 'ID подразделения_x', 'ID подразделения_y',
           'Причина закрытия', 'Процент прохождения', 'Вид обследования', 'Статус актуальный', 'Дата обновления статуса',
           'Текст сообщения', 'Группа здоровья', 'Результат обращения', 'Период', 'Наименование медицинской организации',
           'ОГРН медицинской организации', 'Дата последнего мероприятия 1 этапа диспансеризации', 'Дата рождения пациента',
            'Дата создания карты диспансеризации'], axis=1)

df_final['Подразделение'] = df_final['Структурное подразделение'].apply(lambda x: re.search('ОСП \d', x)[0] if re.match('^ОСП \d.*$', x) else 'Ленинградская 9')
df_final = df_final.rename(columns={'Структурное подразделение': 'Отделение'})

for department in df_final['Подразделение'].unique():
    df_temp = df_final[df_final['Подразделение'] == department].drop(['Подразделение'], axis=1).sort_values('Врач')
    # Фильтрация датафрейма по уникальному значению в колонке
    save_to_excel(df_temp, reports_path + '\\Показатель 24\\' + department + '.xlsx')