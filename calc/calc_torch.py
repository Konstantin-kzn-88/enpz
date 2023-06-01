# -----------------------------------------------------------
# Класс предназначен для расчета факельного горения
#
# Приказ МЧС № 404 от 10.07.2009
# (C) 2023 Kuznetsov Konstantin, Kazan , Russian Federation
# email kuznetsovkm@yandex.ru
# -----------------------------------------------------------
import math

class Torch:

    def liquid_torch(self, consumption: float) -> tuple:
        """
        Расчет зон факельного горения для жидкостного факела

        Parametrs:
        :@param consumption расход, кг/с;

        Return:
        :@return tuple(Lf, Df)
        """
        # Проверки
        if 0 in (consumption,):
            raise ValueError(f'Фукнция не может принимать нулевые параметры')

        Lf = int(15 * math.pow(consumption, 0.4))
        Df = math.ceil(0.15 * Lf)

        return (Lf, Df)

    def gas_torch(self, consumption: float) -> tuple:
        """
        Расчет зон факельного горения для газового факела

        Parametrs:
        :@param consumption расход, кг/с;

        Return:
        :@return tuple(Lf, Df)
        """
        # Проверки
        if 0 in (consumption,):
            raise ValueError(f'Фукнция не может принимать нулевые параметры')

        Lf = int(12.5 * math.pow(consumption, 0.4))
        Df = math.ceil(0.15 * Lf)

        return (Lf, Df)


if __name__ == '__main__':
    pass
