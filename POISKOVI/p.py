import folium
from folium import plugins
from folium.plugins import Search
import openpyxl
import json
import os

#Создание карты
m = folium.Map(location=[45.0448, 38.976], zoom_start = 12)

#Загрузка данных
wb = openpyxl.reader.excel.load_workbook(filename="inf.xlsx")
AL = openpyxl.reader.excel.load_workbook(filename="all.xlsx")


#Global tooltip
tooltip = 'Открой меня'

#Иконки для маркеров
logoIcon = folium.features.CustomIcon('11.png', icon_size=(40, 40))

#Geojson Data (границы Краснодара)
border = os.path.join('new file.json')

#Расстановка маркеров вет. клиник
wb.active = 0
sheet_zero = wb.active
for i in range(2,49):
    name = sheet_zero['A'+str(i)].value
    adds = sheet_zero['B'+str(i)].value
    tel = sheet_zero['D'+str(i)].value
    folium.Marker(location=[sheet_zero['E'+str(i)].value, sheet_zero['F'+str(i)].value],
    popup = '<strong>'+name+'\nАдрес</strong>: '+adds+'\nТелефон:</strong>'+tel,
                  
    tooltip=tooltip,
    icon = folium.Icon(icon = 'cloud',color = 'red')).add_to(m)

#Расстановка маркеров зоомагазинов
wb.active = 1
sheet_one = wb.active
for i in range(2,7):
    folium.Marker(location=[sheet_one['E'+str(i)].value, sheet_one['F'+str(i)].value],
    popup ='<strong>'+ sheet_zero['A'+str(i)].value + '\nАдрес</strong>: ' + sheet_zero['B'+str(i)].value + '\nТелефон:</strong>' + sheet_zero['D'+str(i)].value,
    icon = folium.Icon(icon = 'cloud',color = 'green')).add_to(m)

#Расстановка маркеров приютов и гостиниц
wb.active = 2
sheet_two = wb.active
for i in range(2,8):
    name = sheet_zero['A'+str(i)].value
    adds = sheet_zero['B'+str(i)].value
    tel = sheet_zero['D'+str(i)].value
    folium.Marker(location=[sheet_zero['E'+str(i)].value, sheet_zero['F'+str(i)].value],
    popup = '<strong>'+name+'\nАдрес</strong>: '+adds+'\nТелефон:</strong>'+tel,
    tooltip=tooltip,
    icon = folium.Icon(icon = 'cloud',color = 'blue')).add_to(m)


#Проба маркера с загруженной иконкой
folium.Marker(location=[45.0348, 38.975],
popup = '<strong>l',
tooltip=tooltip,
icon = logoIcon).add_to(m)

folium.GeoJson(border, name='Krasnodar').add_to(m)
FC={
"type": "FeatureCollection",
"features": [
{"type": "Feature","properties": {"name": 'Ветеринарная клиника "Здоровье"' },"geometry": {"type": "Point","coordinates": [ 45.113834 , 38.937515 ] }},  

{"type": "Feature","properties": {"name": 'Ветеринарная клиника "ВетПомощь"' },"geometry": {"type": "Point","coordinates": [ 45.038740 , 38.963058 ] }},  

{"type": "Feature","properties": {"name": 'Ветеринарная клиника "Лапа Помощи"' },"geometry": {"type": "Point","coordinates": [ 45.067518 , 38.958300 ] }},  

{"type": "Feature","properties": {"name": 'Ветеринарная клиника "Доверие"' },"geometry": {"type": "Point","coordinates": [ 45.080308 , 38.976254 ] }},  

{"type": "Feature","properties": {"name": 'Ветеринарная клиника "Друг"' },"geometry": {"type": "Point","coordinates": [ 45.072789 , 38.976773 ] }},  

{"type": "Feature","properties": {"name": 'Ветеринарная клиника "ВитВет"' },"geometry": {"type": "Point","coordinates": [ 45.072235 , 38.982009 ] }},  

{"type": "Feature","properties": {"name": 'Ветеринарная клиника "ДЖиМ"' },"geometry": {"type": "Point","coordinates": [ 45.088839 , 38.979097 ] }},  

{"type": "Feature","properties": {"name": 'Ветеринарная клиника'  },"geometry": {"type": "Point","coordinates": [ 45.044368 , 38.931976 ] }},  


]
}

geojson_obj = folium.GeoJson(FC).add_to(m)

plugins.Search(geojson_obj,position='topleft',
                           search_zoom=17,
                           search_label='name',
                           geom_type='Point').add_to(m)
m.save("m.html")
