import xlwings as xw

from calc import calc_strait_fire
from calc import calc_tvs_explosion

# работа с активной книгой
wb = xw.books.active
ws = wb.sheets['Сценарии']

for i in range(1, 95):
    # Идем пока не пустые строки
    if ws.range(f"A{i}").value != None:
        # Пропускаем шапку
        if i == 1:
            continue
        elif i in [k for k in range(2,100,9)]:
            # print('считаем пожар пролива')
            S_spill = float(ws.range(f"K{i}").value)
            arr = calc_strait_fire.Strait_fire().termal_class_zone(S_spill, m_sg=0.06, mol_mass=100, t_boiling=2, wind_velocity=1)
            ws.range(f"O{i}").value = arr[0]
            ws.range(f"P{i}").value = arr[1]
            ws.range(f"Q{i}").value = arr[2]
            ws.range(f"R{i}").value = arr[3]

        elif i in [k for k in range(3,100,9)]:
            # print('считаем взрыв')
            mass = float(ws.range(f"J{i}").value)*1000 # перевести в кг
            arr = calc_tvs_explosion.Explosion().explosion_class_zone(class_substance=3, view_space=3, mass=mass,
                             heat_of_combustion=46000, sigma=7, energy_level=2)
            ws.range(f"S{i}").value = arr[-4]
            ws.range(f"T{i}").value = arr[-3]
            ws.range(f"U{i}").value = arr[-2]
            ws.range(f"V{i}").value = arr[-1]


    else:
        print(i)
        break
