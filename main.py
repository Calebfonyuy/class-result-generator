import sys
import getopt
from faker import Faker
from random import randint
import xlsxwriter


def generate_students(class_size=20):
    students = []
    faker = Faker()

    for _ in range(class_size):
        students.append(faker.name())

    return students


def generate_marks(students, subject_count=1):
    marks = dict()
    faker = Faker()

    for _ in range(subject_count):
        name = faker.word().upper()
        mark = dict()
        for student in students:
            mark[student] = [randint(0, 20), randint(0, 20)]

        marks[name] = mark

    return marks


def generate_class_marks(nb_students=20, nb_classes=2, subject_count=1):
    classes = []

    for _ in range(nb_classes):
        students = generate_students(nb_students)
        class_marks = generate_marks(students, subject_count)
        classes.append(class_marks)

    return classes


def write_to_files(data=[]):
    count = 1
    for entry in data:
        name = "output/CLASS_" + str(count) + ".xlsx"
        book = xlsxwriter.Workbook(name)
        for subject in entry.keys():
            sheet = book.add_worksheet(subject)
            sheet.merge_range(0, 0, 0, 5, 'MARKS FOR '+subject)
            sheet.merge_range(1, 0, 1, 5, 'SUBJECT COEFFICIENT : '+ str(randint(1, 5)))

            sheet.write(3, 0, "S/N")
            sheet.write(3, 1, "STUDENT")
            sheet.write(3, 2, "MARK 1")
            sheet.write(3, 3, "MARK 2")
            sheet.write(3, 4, "TOTAL")
            sheet.write(3, 5, "AVERAGE")
            sheet.write(3, 6, "POSITION")
            sheet.write(3, 7, "REMARK")

            row = 4
            for mark in entry[subject]:
                sheet.write(row, 0, str(row-3))
                sheet.write(row, 1, mark)
                sheet.write(row, 2, entry[subject][mark][0])
                sheet.write(row, 3, entry[subject][mark][1])
                row += 1

        book.close()
        count += 1


if __name__ == '__main__':
    opts, _ = getopt.getopt(sys.argv[1:], 'n:c:s:')
    n = 20
    c = 2
    s = 1
    for val in opts:
        if val[0] == '-n':
            n = int(val[1])
        if val[0] == '-c':
            c = int(val[1])
        if val[0] == '-s':
            s = int(val[1])

    classes = generate_class_marks(n, c, s)
    # print(classes)
    write_to_files(classes)
