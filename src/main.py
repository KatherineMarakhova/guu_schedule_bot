from openpyxl import *
from openpyxl.utils import range_boundaries
from pathlib import Path
import selenium_fcs as sf
from Direct import *


# obj = Direct()
# obj.set_course(2)
# obj.first_start()
#
# obj.get_list_inst()
# print(f'obj.list_insts: {obj.list_insts}')
# obj.set_inst('ИМ 2 курс')
#
# obj.get_list_napr()
# print(f'obj.list_napr: {obj.list_napr}')
# obj.set_napr('МЕНЕДЖМЕНТ')
#
# obj.get_list_edup()
# print(f'obj.list_edup: {obj.list_edup}')
# obj.set_edup('Бренд-менеджмент')
#
# obj.get_list_group()
# print(f'obj.list_groups: {obj.list_groups}')
# obj.set_edup('1')

obj = Direct()
obj.set_course(3)
obj.first_start()

obj.get_list_inst()
print(f'obj.list_insts: {obj.list_insts}')
obj.set_inst('ИОМ 3 курс')

obj.get_list_napr()
print(f'obj.list_napr: {obj.list_napr}')
obj.set_napr('МЕНЕДЖМЕНТ')

obj.get_list_edup()
print(f'obj.list_edup: {obj.list_edup}')
# obj.set_edup('ПРИКЛАДНАЯ МАТЕМАТИКА И ИНФОРМАТИКА')
#
# obj.get_list_group()
# print(f'obj.list_groups: {obj.list_groups}')
# obj.set_group(1)
#
# # obj.get_scd_even('НЕЧЁТ.')
# obj.get_scd_weekday()




















