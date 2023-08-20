import xlwings as xw

from calc import calc_strait_fire
from calc import calc_tvs_explosion
from calc import calc_fireball
from calc import calc_lower_concentration
from calc import calc_torch

# работа с активной книгой
wb = xw.books.active
ws = wb.sheets['Сценарии']
NUM_ROW = 300

PIPE_GAS_1 = [i for i in range(2, 50)]


# Трубопроводы с газом
for i in PIPE_GAS_1:
    # Идем пока не пустые строки
    if ws.range(f"A{i}").value != None:
        # Пропускаем шапку
        if i == 1:
            continue

        elif i in [k for k in range(2, max(PIPE_GAS_1), 3)]:
            consumption = float(ws.range(f"K{i}").value)  # кг/c
            arr = calc_torch.Torch().liquid_torch(consumption)
            ws.range(f"Y{i}").value = arr[0]
            ws.range(f"Z{i}").value = arr[1]

        elif i in [k for k in range(3, max(PIPE_GAS_1), 6)]:
            # print('считаем взрыв')
            mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
            arr = calc_tvs_explosion.Explosion().explosion_class_zone(class_substance=2, view_space=2, mass=mass,

                                                                      heat_of_combustion=46000, sigma=7, energy_level=2)
            ws.range(f"S{i}").value = arr[-6]
            ws.range(f"T{i}").value = arr[-5]
            ws.range(f"U{i}").value = arr[-4]
            ws.range(f"V{i}").value = arr[-3]
            ws.range(f"W{i}").value = arr[-2]
            ws.range(f"X{i}").value = arr[-1]

        elif i in [k for k in range(6, max(PIPE_GAS_1), 6)]:
            # print('считаем вспышку')
            mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
            arr = calc_lower_concentration.LCLP().lower_concentration_limit(mass=mass, mol_mass=100, t_boiling=2,
                                                                            lower_concentration=3)
            ws.range(f"AA{i}").value = arr[0]
            ws.range(f"AB{i}").value = arr[1]

    else:
        print(i)
        break
