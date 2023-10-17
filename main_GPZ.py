import xlwings as xw

from calc import calc_strait_fire
from calc import calc_tvs_explosion
from calc import calc_fireball
from calc import calc_lower_concentration
from calc import calc_torch

# работа с активной книгой
wb = xw.books.active
ws = wb.sheets['Сценарии']

PIPE_GAS_1 = [i for i in range(2, 50)]
PIPE_GAS_2 = [i for i in range(859, 883)]
PIPE_GAS_3 = [i for i in range(916, 922)]

BUL_GAS_1 = [i for i in range(51, 63)]
BUL_GAS_2 = [i for i in range(884, 896)]
BUL_GAS_3 = [i for i in range(897, 915)]
BUL_GAS_4 = [i for i in range(564, 583)]

BUL_LPG_1 = [i for i in range(64, 245)]
BUL_LPG_2 = [i for i in range(427, 527)]
BUL_LPG_3 = [i for i in range(634, 689)]
BUL_LPG_4 = [i for i in range(763, 836)]

PUMP_LGP_1 = [i for i in range(245, 288)]
PUMP_LGP_2 = [i for i in range(744, 763)]

PIPE_LGP_1 = [i for i in range(288, 427)]
PIPE_LGP_2 = [i for i in range(527, 564)]
PIPE_LGP_3 = [i for i in range(608, 627)]
PIPE_LGP_4 = [i for i in range(689, 744)]
PIPE_LGP_5 = [i for i in range(836, 849)]

BUL_LIQ_1 = [i for i in range(583, 608)]

# # Трубопроводы с газом
for i in PIPE_GAS_1 + PIPE_GAS_2 + PIPE_GAS_3:
    # Идем пока не пустые строки
    if ws.range(f"A{i}").value != None:
        # Пропускаем шапку
        if i == 1:
            continue

        elif i in [k for k in range(2, max(PIPE_GAS_1), 3)] + [k for k in range(859, max(PIPE_GAS_2), 3)] + [k for k in
                                                                                                             range(916,
                                                                                                                   max(PIPE_GAS_3),
                                                                                                                   3)]:
            consumption = float(ws.range(f"K{i}").value)  # кг/c
            arr = calc_torch.Torch().gas_torch(consumption)
            ws.range(f"Y{i}").value = arr[0]
            ws.range(f"Z{i}").value = arr[1]

        elif i in [k for k in range(3, max(PIPE_GAS_1), 6)] + [k for k in range(860, max(PIPE_GAS_2), 6)] + [k for k in
                                                                                                             range(917,
                                                                                                                   max(PIPE_GAS_3),
                                                                                                                   6)]:
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

        elif i in [k for k in range(6, max(PIPE_GAS_1), 6)] + [k for k in range(845, max(PIPE_GAS_2), 6)] + [k for k in
                                                                                                             range(920,
                                                                                                                   max(PIPE_GAS_3),
                                                                                                                   6)]:
            # print('считаем вспышку')
            mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
            arr = calc_lower_concentration.LCLP().lower_concentration_limit(mass=mass, mol_mass=100, t_boiling=2,
                                                                            lower_concentration=3)
            ws.range(f"AA{i}").value = arr[0]
            ws.range(f"AB{i}").value = arr[1]

    else:
        print(i)
        break
print()

# # Оборудование с газом
# for i in BUL_GAS_1 + BUL_GAS_2 + BUL_GAS_3 + BUL_GAS_4:
#     # Идем пока не пустые строки
#     if ws.range(f"A{i}").value != None:
#         # Пропускаем шапку
#         if i == 1:
#             continue
#
#         elif i in [k for k in range(54, max(BUL_GAS_1), 6)] + \
#                 [k for k in range(887, max(BUL_GAS_2), 6)] + \
#                 [k for k in range(900, max(BUL_GAS_3), 6)] + \
#                 [k for k in range(567, max(BUL_GAS_4), 6)]:
#
#             consumption = float(ws.range(f"K{i}").value)  # кг/c
#             arr = calc_torch.Torch().gas_torch(consumption)
#             ws.range(f"Y{i}").value = arr[0]
#             ws.range(f"Z{i}").value = arr[1]
#
#         elif i in [k for k in range(52, max(BUL_GAS_1), 6)] + \
#                 [k for k in range(885, max(BUL_GAS_2), 6)] + \
#                 [k for k in range(898, max(BUL_GAS_3), 6)] + \
#                 [k for k in range(565, max(BUL_GAS_4), 6)]:
#             # print('считаем взрыв')
#             mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
#             arr = calc_tvs_explosion.Explosion().explosion_class_zone(class_substance=2, view_space=2, mass=mass,
#
#                                                                       heat_of_combustion=46000, sigma=7, energy_level=2)
#             ws.range(f"S{i}").value = arr[-6]
#             ws.range(f"T{i}").value = arr[-5]
#             ws.range(f"U{i}").value = arr[-4]
#             ws.range(f"V{i}").value = arr[-3]
#             ws.range(f"W{i}").value = arr[-2]
#             ws.range(f"X{i}").value = arr[-1]
#
#         elif i in [k for k in range(51, max(BUL_GAS_1), 6)] + \
#                 [k for k in range(884, max(BUL_GAS_2), 6)] + \
#                 [k for k in range(55, max(BUL_GAS_1), 6)] + \
#                 [k for k in range(888, max(BUL_GAS_2), 6)] + \
#                 [k for k in range(897, max(BUL_GAS_3), 6)] + \
#                 [k for k in range(901, max(BUL_GAS_3), 6)]+ \
#                 [k for k in range(564, max(BUL_GAS_4), 6)] + \
#                 [k for k in range(568, max(BUL_GAS_4), 6)]:
#             # print('считаем вспышку')
#             mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
#             arr = calc_lower_concentration.LCLP().lower_concentration_limit(mass=mass, mol_mass=100, t_boiling=2,
#                                                                             lower_concentration=3)
#             ws.range(f"AA{i}").value = arr[0]
#             ws.range(f"AB{i}").value = arr[1]
#
#     else:
#         print(i)
#         break
#
# print()
# # Оборудование с СУГ
# for i in BUL_LPG_1+BUL_LPG_2+BUL_LPG_3+BUL_LPG_4:
#
#     if i in [k for k in range(64, max(BUL_LPG_1), 9)]+\
#             [k for k in range(427, max(BUL_LPG_2), 9)]+\
#             [k for k in range(634, max(BUL_LPG_3), 9)]+\
#             [k for k in range(763, max(BUL_LPG_4), 9)]:
#         # print('считаем пожар пролива')
#         S_spill = float(ws.range(f"K{i}").value)
#         arr = calc_strait_fire.Strait_fire().termal_class_zone(S_spill, m_sg=0.1, mol_mass=100, t_boiling=2,
#                                                                wind_velocity=1)
#         ws.range(f"O{i}").value = arr[0]
#         ws.range(f"P{i}").value = arr[1]
#         ws.range(f"Q{i}").value = arr[2]
#         ws.range(f"R{i}").value = arr[3]
#
#     elif i in [k for k in range(65, max(BUL_LPG_1), 9)]+\
#             [k for k in range(428, max(BUL_LPG_2), 9)]+\
#             [k for k in range(635, max(BUL_LPG_3), 9)]+\
#             [k for k in range(764, max(BUL_LPG_4), 9)]:
#         # print('считаем взрыв')
#         mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
#         arr = calc_tvs_explosion.Explosion().explosion_class_zone(class_substance=2, view_space=2, mass=mass,
#
#                                                                   heat_of_combustion=46000, sigma=7, energy_level=2)
#         ws.range(f"S{i}").value = arr[-6]
#         ws.range(f"T{i}").value = arr[-5]
#         ws.range(f"U{i}").value = arr[-4]
#         ws.range(f"V{i}").value = arr[-3]
#         ws.range(f"W{i}").value = arr[-2]
#         ws.range(f"X{i}").value = arr[-1]
#
#     elif i in [k for k in range(67, max(BUL_LPG_1), 9)]+\
#             [k for k in range(430, max(BUL_LPG_2), 9)]+\
#             [k for k in range(637, max(BUL_LPG_3), 9)]+\
#             [k for k in range(766, max(BUL_LPG_4), 9)]:
#
#         consumption = float(ws.range(f"K{i}").value)  # кг/c
#         arr = calc_torch.Torch().liquid_torch(consumption)
#         ws.range(f"Y{i}").value = arr[0]
#         ws.range(f"Z{i}").value = arr[1]
#
#     elif i in [k for k in range(69, max(BUL_LPG_1), 9)]+\
#             [k for k in range(432, max(BUL_LPG_2), 9)]+\
#             [k for k in range(639, max(BUL_LPG_3), 9)]+\
#             [k for k in range(768, max(BUL_LPG_4), 9)]:
#
#         consumption = float(ws.range(f"K{i}").value)  # кг/c
#         arr = calc_torch.Torch().gas_torch(consumption)
#         ws.range(f"Y{i}").value = arr[0]
#         ws.range(f"Z{i}").value = arr[1]
#
#     elif i in [k for k in range(70, max(BUL_LPG_1), 9)]+\
#             [k for k in range(433, max(BUL_LPG_2), 9)]+\
#             [k for k in range(640, max(BUL_LPG_3), 9)]+\
#             [k for k in range(769, max(BUL_LPG_4), 9)]:
#         # print('считаем вспышку')
#         mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
#         arr = calc_lower_concentration.LCLP().lower_concentration_limit(mass=mass, mol_mass=100, t_boiling=2,
#                                                                         lower_concentration=3)
#         ws.range(f"AA{i}").value = arr[0]
#         ws.range(f"AB{i}").value = arr[1]
#
#     elif i in [k for k in range(72, max(BUL_LPG_1), 9)]+\
#             [k for k in range(435, max(BUL_LPG_2), 9)]+\
#             [k for k in range(630, max(BUL_LPG_3), 9)]+\
#             [k for k in range(771, max(BUL_LPG_4), 9)]:
#         # print('считаем ОШ')
#         mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
#         arr = calc_fireball.Fireball().termal_class_zone(mass=mass, ef=450)
#         ws.range(f"AE{i}").value = arr[0]
#         ws.range(f"AF{i}").value = arr[1]
#         ws.range(f"AG{i}").value = arr[2]
#         ws.range(f"AH{i}").value = arr[3]
#
# print()
#
# # Насосы с СУГ
# for i in PUMP_LGP_1+PUMP_LGP_2:
#     if i in [k for k in range(245, max(PUMP_LGP_1), 6)]+[k for k in range(744, max(PUMP_LGP_2), 6)]:
#         consumption = float(ws.range(f"K{i}").value)  # кг/c
#         arr = calc_torch.Torch().liquid_torch(consumption)
#         ws.range(f"Y{i}").value = arr[0]
#         ws.range(f"Z{i}").value = arr[1]
#
#     elif i in [k for k in range(246, max(PUMP_LGP_1), 6)]+[k for k in range(745, max(PUMP_LGP_2), 6)]:
#         # print('считаем взрыв')
#         mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
#         arr = calc_tvs_explosion.Explosion().explosion_class_zone(class_substance=2, view_space=2, mass=mass,
#
#                                                                   heat_of_combustion=46000, sigma=7, energy_level=2)
#         ws.range(f"S{i}").value = arr[-6]
#         ws.range(f"T{i}").value = arr[-5]
#         ws.range(f"U{i}").value = arr[-4]
#         ws.range(f"V{i}").value = arr[-3]
#         ws.range(f"W{i}").value = arr[-2]
#         ws.range(f"X{i}").value = arr[-1]
#
#     elif i in [k for k in range(248, max(PUMP_LGP_1), 6)]+[k for k in range(747, max(PUMP_LGP_2), 6)]:
#         # print('считаем пожар пролива')
#         S_spill = float(ws.range(f"K{i}").value)
#         arr = calc_strait_fire.Strait_fire().termal_class_zone(S_spill, m_sg=0.1, mol_mass=100, t_boiling=2,
#                                                                wind_velocity=1)
#         ws.range(f"O{i}").value = arr[0]
#         ws.range(f"P{i}").value = arr[1]
#         ws.range(f"Q{i}").value = arr[2]
#         ws.range(f"R{i}").value = arr[3]
#
#     elif i in [k for k in range(249, max(PUMP_LGP_1), 6)]+[k for k in range(748, max(PUMP_LGP_2), 6)]:
#         # print('считаем вспышку')
#         mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
#         arr = calc_lower_concentration.LCLP().lower_concentration_limit(mass=mass, mol_mass=100, t_boiling=2,
#                                                                         lower_concentration=3)
#         ws.range(f"AA{i}").value = arr[0]
#         ws.range(f"AB{i}").value = arr[1]
#
# print()
# # Трубы с СУГ
# for i in PIPE_LGP_1 + PIPE_LGP_2 + PIPE_LGP_3 + PIPE_LGP_4 + PIPE_LGP_5:
#     if i in [k for k in range(288, max(PIPE_LGP_1), 3)] + \
#             [k for k in range(527, max(PIPE_LGP_2), 3)] + \
#             [k for k in range(608, max(PIPE_LGP_3), 3)] + \
#             [k for k in range(689, max(PIPE_LGP_4), 3)]+ \
#             [k for k in range(836, max(PIPE_LGP_5), 3)]:
#         consumption = float(ws.range(f"K{i}").value)  # кг/c
#         arr = calc_torch.Torch().liquid_torch(consumption)
#         ws.range(f"Y{i}").value = arr[0]
#         ws.range(f"Z{i}").value = arr[1]
#
#     elif i in [k for k in range(289, max(PIPE_LGP_1), 6)] + \
#             [k for k in range(528, max(PIPE_LGP_2), 6)] + \
#             [k for k in range(609, max(PIPE_LGP_3), 6)] + \
#             [k for k in range(690, max(PIPE_LGP_4), 6)] + \
#             [k for k in range(837, max(PIPE_LGP_5), 6)]:
#         # print('считаем взрыв')
#         mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
#         arr = calc_tvs_explosion.Explosion().explosion_class_zone(class_substance=2, view_space=2, mass=mass,
#
#                                                                   heat_of_combustion=46000, sigma=7, energy_level=2)
#         ws.range(f"S{i}").value = arr[-6]
#         ws.range(f"T{i}").value = arr[-5]
#         ws.range(f"U{i}").value = arr[-4]
#         ws.range(f"V{i}").value = arr[-3]
#         ws.range(f"W{i}").value = arr[-2]
#         ws.range(f"X{i}").value = arr[-1]
#
#     elif i in [k for k in range(292, max(PIPE_LGP_1), 6)] + \
#             [k for k in range(531, max(PIPE_LGP_2), 6)] + \
#             [k for k in range(612, max(PIPE_LGP_3), 6)] + \
#             [k for k in range(693, max(PIPE_LGP_4), 6)] + \
#             [k for k in range(840, max(PIPE_LGP_5), 6)]:
#         # print('считаем вспышку')
#         mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
#         arr = calc_lower_concentration.LCLP().lower_concentration_limit(mass=mass, mol_mass=100, t_boiling=2,
#                                                                         lower_concentration=3)
#         ws.range(f"AA{i}").value = arr[0]
#         ws.range(f"AB{i}").value = arr[1]
# print()
#
# # Оборудование с ЛВЖ
# for i in BUL_LIQ_1:
#     if i in [k for k in range(583, max(BUL_LIQ_1), 3)]:
#         # print('считаем пожар пролива')
#         S_spill = float(ws.range(f"K{i}").value)
#         arr = calc_strait_fire.Strait_fire().termal_class_zone(S_spill, m_sg=0.1, mol_mass=100, t_boiling=2,
#                                                                wind_velocity=1)
#         ws.range(f"O{i}").value = arr[0]
#         ws.range(f"P{i}").value = arr[1]
#         ws.range(f"Q{i}").value = arr[2]
#         ws.range(f"R{i}").value = arr[3]
#
#     elif i in [k for k in range(584, max(BUL_LIQ_1), 6)]:
#         # print('считаем взрыв')
#         mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
#         arr = calc_tvs_explosion.Explosion().explosion_class_zone(class_substance=2, view_space=2, mass=mass,
#
#                                                                   heat_of_combustion=46000, sigma=7, energy_level=2)
#         ws.range(f"S{i}").value = arr[-6]
#         ws.range(f"T{i}").value = arr[-5]
#         ws.range(f"U{i}").value = arr[-4]
#         ws.range(f"V{i}").value = arr[-3]
#         ws.range(f"W{i}").value = arr[-2]
#         ws.range(f"X{i}").value = arr[-1]
#
#     elif i in [k for k in range(587, max(BUL_LIQ_1), 6)]:
#         # print('считаем вспышку')
#         mass = float(ws.range(f"J{i}").value) * 1000  # перевести в кг
#         arr = calc_lower_concentration.LCLP().lower_concentration_limit(mass=mass, mol_mass=100, t_boiling=2,
#                                                                         lower_concentration=3)
#         ws.range(f"AA{i}").value = arr[0]
#         ws.range(f"AB{i}").value = arr[1]
