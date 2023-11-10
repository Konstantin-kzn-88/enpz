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
ROW_START_PRESSURE = 1
ROW_START_TANK = 219


# Оборудование
for i in range(1, 122):
    # Идем пока не пустые строки
    if i in [k for k in range(10, 83, 9)]:
        # print('считаем шар')
        mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
        arr = calc_fireball.Fireball().termal_class_zone(mass, ef=450)
        ws.range(f"AE{i}").value = arr[0]
        ws.range(f"AF{i}").value = arr[1]
        ws.range(f"AG{i}").value = arr[2]
        ws.range(f"AH{i}").value = arr[3]

    elif i in [k for k in range(3, 83, 9)]+[k for k in range(85, 102, 6)]+[k for k in range(104, 121, 6)]:
        # print('считаем взрыв')
        mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
        arr = calc_tvs_explosion.Explosion().explosion_class_zone(class_substance=3, view_space=3, mass=mass,
                                                                  heat_of_combustion=46000, sigma=4, energy_level=1)
        ws.range(f"U{i}").value = arr[-4]
        ws.range(f"V{i}").value = arr[-3]
        ws.range(f"W{i}").value = arr[-2]
        ws.range(f"X{i}").value = arr[-1]
    #
    elif i in [k for k in range(8, 83, 9)]+[k for k in range(88, 102, 6)]+[k for k in range(107, 121, 6)]:
        # print('считаем вспышку')
        mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
        arr = calc_lower_concentration.LCLP().lower_concentration_limit(mass=mass, mol_mass=100, t_boiling=2,
                                                                        lower_concentration=3)
        ws.range(f"AA{i}").value = arr[0]
        ws.range(f"AB{i}").value = arr[1]

    elif i in [k for k in range(5, 83, 9)]+[k for k in range(84, 102, 6)]:

        consumption = float(ws.range(f"K{i}").value)  # кг/c
        arr = calc_torch.Torch().liquid_torch(consumption)
        ws.range(f"Y{i}").value = arr[0]
        ws.range(f"Z{i}").value = arr[1]
    #
    elif i in [k for k in range(7, 83, 9)]:

        consumption = float(ws.range(f"K{i}").value)  # кг/c
        arr = calc_torch.Torch().gas_torch(consumption)
        ws.range(f"Y{i}").value = arr[0]
        ws.range(f"Z{i}").value = arr[1]
    #
    elif i in [k for k in range(2, 83, 9)]+[k for k in range(87, 102, 6)]+[k for k in range(103, 121, 6)]+[k for k in range(106, 121, 6)]:
        # print('считаем пожар пролива')
        S_spill = float(ws.range(f"K{i}").value)
        arr = calc_strait_fire.Strait_fire().termal_class_zone(S_spill, m_sg=0.06, mol_mass=100, t_boiling=2,
                                                               wind_velocity=1)
        ws.range(f"O{i}").value = arr[0]
        ws.range(f"P{i}").value = arr[1]
        ws.range(f"Q{i}").value = arr[2]
        ws.range(f"R{i}").value = arr[3]


