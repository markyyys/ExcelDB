from tkinter import * # Библиотека для создания окна
from tkinter.ttk import Combobox, Entry
from openpyxl import * # Библиотека для Экселя

window = Tk() # Создание окна
window.title('EXCEl') # Название окна
window.geometry('600x400')
window.configure(padx=50, pady=50) # Оступы от краев окна

wb = load_workbook(filename = 'data/data.xlsx') # Загружаем файл экселя

def OpenNewPersonWindow(): # Откратие окна нового сотрудника
    windowPerson = Tk()
    windowPerson.title('Новый сотрудник')
    windowPerson.geometry('600x400') # Создание окна нового сотрудника

    lblFIO = Label(windowPerson, text='Фамилия И.О.', font=("Arial Bold", 14)).grid(column=1, row=2, padx=20, pady=10)
    lblPost = Label(windowPerson, text='Должность', font=("Arial Bold", 14)).grid(column=1, row=3, padx=10, pady=10)
    lblOffice = Label(windowPerson, text='Отдел', font=("Arial Bold", 14)).grid(column=1, row=4, padx=20, pady=10)
    lblNumber = Label(windowPerson, text='Телефон', font=("Arial Bold", 14)).grid(column=1, row=5, padx=20, pady=10) # Названия полей расположенные слева

    FIO = StringVar(windowPerson) # Переменная куда вводится значение из поля для ввода StringVar
    enFIO = Entry(windowPerson, width=20, font=("Arial Bold", 14), textvariable=FIO).grid(column=2, row=2, padx=20, pady=10) # Поле для ввода Entry
    Post = StringVar(windowPerson)
    enPost = Entry(windowPerson, width=20, font=("Arial Bold", 14), textvariable=Post).grid(column=2, row=3, padx=20, pady=10)
    Office = StringVar(windowPerson)
    enOffice = Entry(windowPerson, width=20, font=("Arial Bold", 14), textvariable=Office).grid(column=2, row=4, padx=20, pady=10)
    Number = StringVar(windowPerson)
    enNumber = Entry(windowPerson, width=20, font=("Arial Bold", 14), textvariable=Number).grid(column=2, row=5, padx=20, pady=10)

    def Ok(): # Кнопка окей
        sheet = wb['Сотрудники'] # Открываем страницу Сотрудники в ексель документе
        sheet.append([FIO.get(), Post.get(), Office.get(), Number.get()]) # Вводим в новую строку значения с полей для ввода на экране
        wb.save('data/data.xlsx') # Сохраняем новую таблицу


    btnOk = Button(windowPerson, text='Ок', width=20, command=Ok).grid(column=1, row=8, padx=20, pady=20) # Кнопка Окей
    btnCancel = Button(windowPerson, text='Отмена', width=20, command=windowPerson.destroy).grid(column=2, row=8, padx=20, pady=20) # Кнопка отмена


def OpenNewOrganisationWindow():
    windowOrganisation = Tk()
    windowOrganisation.title('Новая организация')
    windowOrganisation.geometry('600x400')

    lblName = Label(windowOrganisation, text='Название', font=("Arial Bold", 14)).grid(column=1, row=1, padx=20, pady=10)
    lblAdress = Label(windowOrganisation, text='Адресс', font=("Arial Bold", 14)).grid(column=1, row=2, padx=20, pady=10)
    lblNumber = Label(windowOrganisation, text='Номер телефона', font=("Arial Bold", 14)).grid(column=1, row=3, padx=20, pady=10)

    Name = StringVar(windowOrganisation)
    enName = Entry(windowOrganisation, width=20, font=("Arial Bold", 14), textvariable=Name).grid(column=2, row=1, padx=20, pady=10)
    Adress = StringVar(windowOrganisation)
    enAdress = Entry(windowOrganisation, width=20, font=("Arial Bold", 14), textvariable=Adress).grid(column=2, row=2, padx=20, pady=10)
    Number = StringVar(windowOrganisation)
    enNumber = Entry(windowOrganisation, width=20, font=("Arial Bold", 14), textvariable=Number).grid(column=2, row=3, padx=20, pady=10)

    def Ok():
        sheet = wb['Организации']
        sheet.append([Name.get(), Adress.get(), Number.get()])
        wb.save('data/data.xlsx')

    btnOk = Button(windowOrganisation, text='Ок', width=20, command=Ok).grid(column=1, row=8, padx=20, pady=20)
    btnCancel = Button(windowOrganisation, text='Отмена', width=20, command=windowOrganisation.destroy).grid(column=2, row=8, padx=20, pady=20)

def OpenNewTravelWindow():
    windowTravel = Tk()
    windowTravel.title('Новая командировка')
    windowTravel.geometry('600x400')

    lblDate = Label(windowTravel, text='Дата выезда', font=("Arial Bold", 14)).grid(column=1, row=1, padx=20, pady=10)
    lblFIO = Label(windowTravel, text='Фамилия И.О.', font=("Arial Bold", 14)).grid(column=1, row=2, padx=20, pady=10)
    lblOrganisation = Label(windowTravel, text='Организация', font=("Arial Bold", 14)).grid(column=1, row=3, padx=10, pady=10)
    lblDaysNumber = Label(windowTravel, text='Количество дней', font=("Arial Bold", 14)).grid(column=1, row=4, padx=20, pady=10)
    lblDailies = Label(windowTravel, text='Дневные', font=("Arial Bold", 14)).grid(column=1, row=5, padx=20, pady=10)
    lblTicketsPrice = Label(windowTravel, text='Цена билета', font=("Arial Bold", 14)).grid(column=1, row=6, padx=10, pady=10)
    lblSumm = Label(windowTravel, text='Сумма', font=("Arial Bold", 14)).grid(column=1, row=7, padx=20, pady=10)

    Date = StringVar(windowTravel)
    enDate = Entry(windowTravel, width=20, font=("Arial Bold", 14), textvariable=Date).grid(column=2, row=1, padx=20, pady=10)

    sheet = wb['Сотрудники'] # Переходим на страницу сотрудников
    PersonList = [] # Массив в котором будут лежать ФИО сотрудников
    for cell in sheet['A']: # Цикл по строкам первого столбца страницы сотрудников
        if cell.value != 'Фамилия И.О.': # Если значение ячейки равно ее названию то мы его не сохраняем в массив
            PersonList.append(cell.value) # Доваляем ФИО стотрудника из первго столбца

    enFIO = Combobox(windowTravel, width=19, font=("Arial Bold", 14)) # Создаем список в котором можно выбрать сотрудника
    enFIO.grid(column=2, row=2, padx=20, pady=10)
    enFIO['values'] = PersonList # Говорим что во всплывающем цикле должны быть значения из этого массива

    sheet = wb['Организации']
    OrganisationList = []
    for cell in sheet['A']:
        if cell.value != 'Наименование организации':
            OrganisationList.append(cell.value)

    enOrganisation = Combobox(windowTravel, width=19, font=("Arial Bold", 14))
    enOrganisation.grid(column=2, row=3, padx=20, pady=10)
    enOrganisation['values'] = OrganisationList

    DaysNumber = StringVar(windowTravel)
    enDaysNumber = Entry(windowTravel, width=20, font=("Arial Bold", 14), textvariable=DaysNumber).grid(column=2, row=4, padx=20, pady=10)
    Dailies = StringVar(windowTravel)
    enDailies = Entry(windowTravel, width=20, font=("Arial Bold", 14), textvariable=Dailies).grid(column=2, row=5, padx=20, pady=10)
    TicketsPrice = StringVar(windowTravel)
    enTicketsPrice = Entry(windowTravel, width=20, font=("Arial Bold", 14), textvariable=TicketsPrice).grid(column=2, row=6, padx=20, pady=10)

    Summ = StringVar(windowTravel)
    enSumm = Label(windowTravel, width=20, font=("Arial Bold", 14), textvariable=Summ).grid(column=2, row=7, padx=20, pady=10)

    def callback(*args): # Функция изменения суммы
        if Dailies.get().isdigit() and DaysNumber.get().isdigit() and TicketsPrice.get().isdigit(): # Проверка на то чтобы значение в поле для ввода Дневных, Количества дней и Цены на билеты было цифрой
            Summ.set(value=str(int(Dailies.get()) * int(DaysNumber.get()) + (int(TicketsPrice.get()) * 2)) + ' руб.') # Переводим введенный текст в цифровой вид, вычисляем сумму и переводим обратно в строковый вид

    DaysNumber.trace_add('write', callback) # Говорим что функция Callback ызывается при изменении поля для ввода
    Dailies.trace_add('write', callback)
    TicketsPrice.trace_add('write', callback)

    def Ok():
        sheet = wb['Командировки']
        sheet.append([Date.get(), enFIO.get(), enOrganisation.get(), DaysNumber.get(), Dailies.get(), TicketsPrice.get(), Summ.get()])
        wb.save('data/data.xlsx')

    btnOk = Button(windowTravel, text='Ок', width=20, command=Ok).grid(column=1, row=8, padx=20, pady=20)
    btnCancel = Button(windowTravel, text='Отмена', width=20, command=windowTravel.destroy).grid(column=2, row=8, padx=20, pady=20)




btnNewPerson = Button(window, width=20, text='Новый сотрудник', command=OpenNewPersonWindow, font=("Arial Bold", 14)) # Кнопка нового сотрудника
btnNewPerson.grid(column=1, row=1, padx=20, pady=20) # Настраиваем положение кнопки в левой верхней части

btnNewOrganisation = Button(window, width=20, text='Новая организация', command=OpenNewOrganisationWindow, font=("Arial Bold", 14)) # Кнопка оргпнизации
btnNewOrganisation.grid(column=2, row=1, padx=20, pady=20) # В правой верхней

btnNewTravel = Button(window, width=20, text='Новая коммандировка', command=OpenNewTravelWindow, font=("Arial Bold", 14)) # Кнопка командировок
btnNewTravel.grid(column=1, row=2, padx=20, pady=20) # В левой нижней

btnExit = Button(window, width=20, text='Выход', command=window.quit, font=("Arial Bold", 14)) # Кнопка выход
btnExit.grid(column=2, row=2, padx=20, pady=20) # В правой нижней





window.mainloop() # Конец работы кадра приложения