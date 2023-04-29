import psycopg2
from psycopg2 import OperationalError
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

#Функция для подключения к базе данных
def create_connection(db_name, db_user, db_password, db_host, db_port):
    connection = None
    try:
        connection = psycopg2.connect(
            database=db_name,
            user=db_user,
            password=db_password,
            host=db_host,
            port=db_port,
        )
        print("Connection to PostgreSQL DB successful")
    except OperationalError as e:
        print(f"The error '{e}' occurred")
    return connection

#Функция для запроса
def execute_read_query(connection, query):
    cursor = connection.cursor()
    result = None
    try:
        cursor.execute(query)
        result = cursor.fetchall()
        return result
    except OperationalError as e:
        print(f"The error '{e}' occurred")

connection = create_connection("pleasebelast", "postgres", "yngaunga", "127.0.0.1", "5432")

# select_users = "SELECT week_day, week_num, lesson_id, subject.name, teacher.name FROM schedule " \
#                "LEFT JOIN subject on subject.subject_id = schedule.subject_id " \
#                "LEFT JOIN teacher on schedule.teacher_id = teacher.teacher_id " \
#                "WHERE group_id = 'ПМИ4-1'"
# users = execute_read_query(connection, select_users)
#
# for user in users:
#     print(user)

# a = process.extract("Шаманин", teachers, limit=2)
# print(a)

# b = process.extract("Крама", teachers, limit=2)
# print(a)

select_teachers = "SELECT name FROM teacher"
teachers = execute_read_query(connection, select_teachers)

for name in teachers:
    res = fuzz.WRatio('НА шананин', name)
    if res>=80:
        print(name, res)
