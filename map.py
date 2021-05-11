import folium
import openpyxl
import json
import os

#Создание карты 
m = folium.Map(location=[45.0448, 38.976], zoom_start = 12)

#Загрузка данных 
wb = openpyxl.reader.excel.load_workbook(filename="data.xlsx")

#Global tooltip
tooltip = 'Информация'


#Geojson Data (границы Краснодара)
border = os.path.join('border.json')

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
    Sh = 'Shelter' + str(i) + '.png'
    ShelterIcon = folium.features.CustomIcon(Sh, icon_size=(40, 40))
    folium.Marker(location=[sheet_one['E'+str(i)].value, sheet_one['F'+str(i)].value],
    popup ='<strong>'+ sheet_one['A'+str(i)].value + '\nАдрес</strong>: ' + sheet_one['B'+str(i)].value + '\nТелефон:</strong>' + sheet_zero['D'+str(i)].value,
    icon = ShelterIcon).add_to(m)

#Расстановка маркеров приютов и гостиниц
wb.active = 2
sheet_two = wb.active   
for i in range(2,8):
    name = sheet_two['A'+str(i)].value
    adds = sheet_two['B'+str(i)].value
    tel = sheet_two['D'+str(i)].value
    folium.Marker(location=[sheet_two['E'+str(i)].value, sheet_two['F'+str(i)].value],
                  popup = '<strong>'+name+'\nАдрес</strong>: '+adds+'\nТелефон:</strong>'+tel,
                  tooltip=tooltip,
                  icon = folium.Icon(icon = 'cloud',color = 'blue')).add_to(m)

    
#Проба маркера с загруженной иконкой
folium.Marker(location=[45.0348, 38.975],
              popup = '<strong>l',
              tooltip=tooltip,
              icon = ShelterIcon).add_to(m)


#Формирование границы             
folium.GeoJson(border, name = 'Krasnodar').add_to(m)

print('ILYT')

m.save("m.html")

