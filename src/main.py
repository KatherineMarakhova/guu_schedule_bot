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
obj.set_course(1)
obj.first_start()

obj.get_list_inst()
print(f'Институты : {obj.list_insts}')

# for inst in obj.list_insts:
#     sheet = obj.wb[inst]
#     print(obj.get_indexes(sheet, 'направление'))
#     obj.set_inst(inst)
#     obj.get_list_napr()
#     print(f'Направления: {obj.list_napr}')


obj.set_inst('ИУПСиБК 1 курс')
sheet = obj.wb['ИУПСиБК 1 курс']
obj.get_list_napr()
print(f'naprs: {obj.list_napr}')
obj.set_napr('ГОСТИНИЧНОЕ ДЕЛО')
obj.get_list_edup()
print(f'edups: {obj.list_edup}')
obj.set_edup('Гостиничный и ресторанный бизнес')
obj.get_list_group()
print(f'groups {obj.list_groups}')
obj.set_group(1)
print(obj.get_scd_full())
# print(f'get_indexes {obj.get_indexes(sheet, "направление")}')
# print(obj.next_idx(sheet, 'политология'))
# print(f'{obj.get_indexes_cat(sheet, "СОЦИОЛОГИЯ", "направление")}')
# print(obj.next_idx_cat(sheet, "СОЦИОЛОГИЯ", 'направление'))

# print(f'day {sheet.cell(10, 3).value}')
# obj.set_inst('ИОМ 2 курс')
# sheet = obj.wb['ИОМ 2 курс']
# obj.get_list_napr()
# print(f'naprs: {obj.list_napr}')
# obj.set_napr('МЕНЕДЖМЕНТ')
# obj.get_list_edup()
# print(f'edups: {obj.list_edup}')

# print(f'get_indexes {obj.get_indexes(sheet, "ПРИКЛАДНАЯ ИНФОРМАТИКА")}')
# # print(obj.next_idx(sheet, 'политология'))
# print(obj.next_idx_cat(sheet, 'ПРИКЛАДНАЯ ИНФОРМАТИКА', 'направление'))




















