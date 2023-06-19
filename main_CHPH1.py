import xlwings as xw

from calc import calc_strait_fire
from calc import calc_tvs_explosion
from calc import calc_lower_concentration
from calc import calc_torch

# работа с активной книгой
wb = xw.books.active
ws = wb.sheets['Сценарии']
NUM_ROW = 300
ROW_START_PUMP = 63
ROW_START_TANK = 2




# Емкости
for i in range(ROW_START_TANK, ROW_START_TANK + NUM_ROW):
    # Идем пока не пустые строки
    if ws.range(f"A{i}").value != None:

        if i in [k for k in range(ROW_START_TANK+1, ROW_START_TANK + NUM_ROW, 6)]:
            # print('считаем взрыв')
            mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
            arr = calc_tvs_explosion.Explosion().explosion_class_zone(class_substance=3, view_space=3, mass=mass,
                                                                      heat_of_combustion=46000, sigma=7, energy_level=2)
            ws.range(f"S{i}").value = arr[-4]
            ws.range(f"T{i}").value = arr[-3]
            ws.range(f"U{i}").value = arr[-2]
            ws.range(f"V{i}").value = arr[-1]

        elif i in [k for k in range(ROW_START_TANK, ROW_START_TANK + NUM_ROW, 3)]:
            # print('считаем пожар пролива')
            S_spill = float(ws.range(f"K{i}").value)
            arr = calc_strait_fire.Strait_fire().termal_class_zone(S_spill, m_sg=0.06, mol_mass=100, t_boiling=2,
                                                                   wind_velocity=1)
            ws.range(f"O{i}").value = arr[0]
            ws.range(f"P{i}").value = arr[1]
            ws.range(f"Q{i}").value = arr[2]
            ws.range(f"R{i}").value = arr[3]

        elif i in [k for k in range(ROW_START_TANK+4, ROW_START_TANK + NUM_ROW, 6)]:
            # print('считаем вспышку')
            mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
            arr = calc_lower_concentration.LCLP().lower_concentration_limit(mass=mass, mol_mass=100, t_boiling=2,
                                                                            lower_concentration=3)
            ws.range(f"Y{i}").value = arr[0]
            ws.range(f"Z{i}").value = arr[1]

    else:
        print(i)
        break
#
# Насосы
for i in range(ROW_START_PUMP, ROW_START_PUMP + NUM_ROW):
    # Идем пока не пустые строки
    if ws.range(f"A{i}").value != None:
        if i in [k for k in range(ROW_START_PUMP, ROW_START_PUMP + NUM_ROW, 6)]:
            consumption = float(ws.range(f"J{i}").value)/30  # кг/c
            arr = calc_torch.Torch().liquid_torch(consumption)
            ws.range(f"W{i}").value = arr[0]
            ws.range(f"X{i}").value = arr[1]

        elif i in [k for k in range(ROW_START_PUMP+1, ROW_START_PUMP + NUM_ROW, 6)]:
            # print('считаем взрыв')
            mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
            arr = calc_tvs_explosion.Explosion().explosion_class_zone(class_substance=3, view_space=3, mass=mass,
                                                                      heat_of_combustion=46000, sigma=7, energy_level=2)
            ws.range(f"S{i}").value = arr[-4]
            ws.range(f"T{i}").value = arr[-3]
            ws.range(f"U{i}").value = arr[-2]
            ws.range(f"V{i}").value = arr[-1]

        elif i in [k for k in range(ROW_START_PUMP+3, ROW_START_PUMP + NUM_ROW, 6)]:
            # print('считаем пожар пролива')
            S_spill = float(ws.range(f"K{i}").value)
            arr = calc_strait_fire.Strait_fire().termal_class_zone(S_spill, m_sg=0.06, mol_mass=100, t_boiling=2,
                                                                   wind_velocity=1)
            ws.range(f"O{i}").value = arr[0]
            ws.range(f"P{i}").value = arr[1]
            ws.range(f"Q{i}").value = arr[2]
            ws.range(f"R{i}").value = arr[3]

        elif i in [k for k in range(ROW_START_PUMP+4, ROW_START_PUMP + NUM_ROW, 6)]:
            # print('считаем вспышку')
            mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
            arr = calc_lower_concentration.LCLP().lower_concentration_limit(mass=mass, mol_mass=100, t_boiling=2,
                                                                            lower_concentration=3)
            ws.range(f"Y{i}").value = arr[0]
            ws.range(f"Z{i}").value = arr[1]

    else:
        print(i)
        break
