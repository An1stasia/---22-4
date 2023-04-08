import pandas as pd
import xlwings as xw
import random
import csv

'''Задания для совместного разбора'''

'''Задания 1, 2'''
b0 = xw.Book('себестоимостьА_в1.xlsx')
sh0 = b0.sheets['Рецептура']
fml0 = sh0.range('T7').formula = '=G7*G$14+H7*H$14+I7*I$14+J7*J$14+K7*K$14+L7*L$14+M7*M$14+O7*O$14'
sh0.range('T8:T10').formula = fml0
sh0.range('T4:T6').merge()
sh0.range('T4:T6').value = 'Себестоимость'

'''Задания 3, 4'''
sh0.range('T7:T13').color = (255, 255, 153)
sh0.range('T3').color = (255, 255, 0)
sh0.range('T4').color = (255, 192, 44)
sh0.range('T14:T16').merge()
sh0.range('S3:T3').merge()
sh0.range('T4').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
sh0.range('T4').font.bold = True
sh0.range('T4').font.color = "#ffffff"
sh0.range('T4:T16').api.Borders.Weight = 2

b0.save()


'''Задание 1'''
recipes = pd.read_csv('recipes_sample.csv', delimiter=',')
reviews = pd.read_csv('reviews_sample.csv', delimiter=',')
recipes_new = recipes[['id', 'name', 'minutes', 'submitted', 'description', 'n_ingredients']]
#print(recipes_new)

'''Задание 2'''
b = xw.Book()
b.save('recipes.xlsx')
b.sheets.add('Отзывы')
b.sheets.add('Рецепты')
tab1 = recipes_new.sample(round(len(recipes) * 0.05))
tab2 = reviews.sample(round(len(reviews) * 0.05))
sh1 = b.sheets['Рецепты']
sh2 = b.sheets['Отзывы']
sh1.range('A1').value = tab1
sh2.range('A1').value = tab2
sh1.range('A:A').api.Delete()
sh2.range('A:A').api.Delete()

'''Задание 3'''
tab1['seconds_assign'] = tab1['minutes']*60
sh1.range('G1').value = tab1['seconds_assign']
sh1.range('G:G').api.Delete()

'''Задание 4'''
sh1.range('H1').value = 'seconds_formula'
fml = sh1.range('H2').formula = '=C2*60'
sh1.range('H2:H1501').formula = fml

'''Задание 5'''
sh1.range('A1:E1').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
sh2.range('A1:F1').api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
sh1.range('A1:H1').font.bold = True
sh2.range('A1:F1').font.bold = True

'''Задание 6'''
k = 1
for i in tab1['minutes']:
    k+=1
    if i<5:
        sh1.range(f'C{k}').color = (0, 255, 0)
    elif i>=5 and i<=10:
        sh1.range(f'C{k}').color = (255, 255, 0)
    elif i>10:
        sh1.range(f'C{k}').color = (255, 0, 0)

'''Задание 7'''
'''col_rev = {}
for i in tab2['recipe_id']:
    if i not in col_rev.keys():
        col_rev[i] = 1
    else:
        col_rev[i] += 1
'''
sh1.range('I1').value = 'n_reviews'
fml7 = sh1.range('I2').formula = '=COUNTIF(Отзывы!$B$2:Отзывы!$B$6336, "="&Рецепты!A2)'
sh1.range('I2:I1501').formula = fml7

'''Задание 8'''
def validate(k):
    if (sh2.range(f'E{k}').value >= 0) and (sh2.range(f'E{k}').value <= 5):
        if sh2.range(f'C{k}').value in sh1.range('C2:C1501').value:
            pass
    else:
        sh2.range(f'A{k}:F{k}').color = (255, 0, 0)

for i in (2, len(tab2)+1):
    validate(i)

b.save()

'''Задание 9'''
'''with open('recipes_model.csv', 'r') as csvfile:
    recipes_model = csv.writer(csvfile, delimiter='\t')'''
recipes_model = pd.read_csv('recipes_model.csv', delimiter='\t')
b2 = xw.Book()
b2.save('recipes_model.xlsx')
b2.sheets.add('Модель')
sh2_1 = b2.sheets['Модель']
sh2_1.range('A2').value = recipes_model
sh2_1.range('A:A').api.Delete()

'''Задание 10'''
sh2_1.range('G2').value = 'Ключ'
sh2_1.range('D2').value = 'Обязательно к заполнению'
sh2_1.range('J2').value = 'Формула'
'''for i in sh2_1.range('G3:G18').value:
    if i == 'PK':
        fml10 = sh2_1.range('J3').formula = '=B3&" "&C3&" "&"PRIMARY KEY"'
        sh1.range('J3:J18').formula = fml10'''
for i in range(3,19):
    if sh2_1.range(f'G{i}').value == 'PK':
        fml10 = sh2_1.range(f'J{i}').formula = f'=B{i}&" "&C{i}&" "&"PRIMARY KEY"'
    elif sh2_1.range(f'G{i}').value == 'FK':
        fml101 = sh2_1.range(f'J{i}').formula = f'=B{i}&" "&C{i}&" "&"REFERENCES"&" "&H{i}&"("&I{i}&")"'
    else:
        if sh2_1.range(f'D{i}').value == 'Y':
            sh2_1.range(f'J{i}').value = 'NOT NULL'

'''Задание 11'''
sh2_1.range('A2:J2').color = (0, 204, 255)
sh2_1.range('A2:J2').font.bold = True
for ws in b2.sheets:
    ws.autofit(axis="columns")
    ws.used_range.api.AutoFilter(Field:=1)
b2.save()