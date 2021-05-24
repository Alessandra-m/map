import folium
from folium import plugins
from folium.plugins import Search
from numpy import True_
import openpyxl
import json
import os
import webbrowser

#Создание карты 
m = folium.Map(location=[45.0448, 38.976], zoom_start = 12)
folium.TileLayer('cartodbpositron').add_to(m)

#Загрузка данных 
wb = openpyxl.reader.excel.load_workbook(filename="data.xlsx")

#Global tooltip
tooltip = 'Информация'

#Расстановка маркеров вет. клиник 
wb.active = 0
sheet_zero = wb.active
vet = os.path.join('Icon','vet.png')
for i in range(2,49):
    vetIcon = folium.features.CustomIcon(vet, icon_size=(25, 25))
    name = sheet_zero['A'+str(i)].value
    adds = sheet_zero['B'+str(i)].value
    tel = sheet_zero['D'+str(i)].value
    folium.Marker(location=[sheet_zero['E'+str(i)].value, sheet_zero['F'+str(i)].value],
                  popup = '<strong>'+name+'\nАдрес</strong>: '+adds+'\nТелефон:</strong>'+tel,
                  tooltip = tooltip,
                  icon = vetIcon).add_to(m)

#Расстановка маркеров зоомагазинов
zoo = os.path.join('Icon','zoo.png')
wb.active = 1
sheet_one = wb.active
for i in range(2,8):
    zooIcon = folium.features.CustomIcon(zoo, icon_size=(30, 30))
    folium.Marker(location=[sheet_one['E'+str(i)].value, sheet_one['F'+str(i)].value],
    popup ='<strong>'+ sheet_one['A'+str(i)].value + '\nАдрес</strong>: ' + sheet_one['B'+str(i)].value + '\nТелефон:</strong>' + sheet_zero['D'+str(i)].value,
    icon = zooIcon).add_to(m)

#Расстановка маркеров приютов и гостиниц
priut = os.path.join('Icon','priut.png')
wb.active = 2
sheet_two = wb.active   
for i in range(2,8):
    priutIcon = folium.features.CustomIcon(priut, icon_size=(32, 32))
    name = sheet_two['A'+str(i)].value
    adds = sheet_two['B'+str(i)].value
    tel = sheet_two['D'+str(i)].value
    folium.Marker(location = [sheet_two['E'+str(i)].value, sheet_two['F'+str(i)].value],
                  popup = '<strong>'+name+'\nАдрес</strong>: '+adds+'\nТелефон:</strong>'+tel,
                  tooltip = tooltip,
                  icon = priutIcon,).add_to(m)

#Расстановка маркеров кафе, пиццерии и бары
cafe = os.path.join('Icon','Cafe.png')
wb.active = 3
sheet_three = wb.active   
for i in range(2,26):
    cafeIcon = folium.features.CustomIcon(cafe, icon_size=(35, 35))
    name = sheet_three['A'+str(i)].value
    adds = sheet_three['B'+str(i)].value
    tel = sheet_three['C'+str(i)].value
    folium.Marker(location = [sheet_three['D'+str(i)].value, sheet_three['E'+str(i)].value],
                  popup = '<strong>'+name+'\nАдрес</strong>: '+adds+'\nТелефон:</strong>'+tel,
                  tooltip = tooltip,
                  icon = cafeIcon,).add_to(m)

#Парки
park = os.path.join('Icon','park.png')
wb.active = 4
sheet_four = wb.active   
for i in range(2,6):
    parkIcon = folium.features.CustomIcon(park, icon_size=(30, 30))
    name = sheet_four['A'+str(i)].value
    adds = sheet_four['B'+str(i)].value
    tel = sheet_four['C'+str(i)].value
    folium.Marker(location = [sheet_four['D'+str(i)].value, sheet_four['E'+str(i)].value],
                  popup = '<strong>'+name+'\nАдрес</strong>: '+adds+'\nТелефон:</strong>'+tel,
                  tooltip = tooltip,
                  icon = parkIcon,).add_to(m)
    

#Формирование границы Краснодара
border = os.path.join('Информация','border.json')            
folium.GeoJson(border, name = 'Krasnodar', 
               control = False,
               style_function = lambda feature: {
                                        "fillColor": "#ffff00",
                                        "color": "black",
                                        "weight": 1.5,
                                        "dashArray": "4, 4"}).add_to(m)

#Поисковик
with open('searcher.json', 'r', encoding = 'utf-8') as f:
    FC = json.load(f)

geojson_obj = folium.GeoJson(FC, name = "Poisk", show = False).add_to(m)

plugins.Search(geojson_obj,position = 'topleft',
                           search_zoom = 17,
                           search_label = 'name',
                           geom_type = 'Point',).add_to(m)

#Контроль уровней
folium.LayerControl().add_to(m)

#Геолокация
plugins.LocateControl(auto_start=True).add_to(m)

m.save("m.html")
webbrowser.open("m.html")

