# name = "thanh"
# for i in name:
#    print(i)
# fruits = ["chuối","táo","xoài"]
# for index in range(len(fruits)):
#    print("Em thích ăn:", fruits[index])
# print(list(range(len(fruits))))
import xlwings as xw
#wb = xw.Book()
app = xw.App(visible=True,add_book=False)
wb = app.books.add()
path = r"D:\EXERCISE_PY\B3\using app\\"
for i in range(1,11):
    wb.save(path + str(i) + ".xlsx")
#for i in range(1,11):
#    wb.save(path + str(i) + ".xlsx")
wb.close()