import datetime
import openpyxl
import tkinter as tk

# открываем файл и лист
workbook = openpyxl.load_workbook('feedback.xlsx')
sheet = workbook['feedback']

# функция для сохранения данных в таблицу
def save_feedback():
    # получаем данные
    gender = gender_var.get()
    age = age_entry.get()
    mood = mood_entry.get()
    now = datetime.datetime.now()
    
    # добавляем новую строку с данными, сохраняем, очищаем
    sheet.append([now, gender, age, mood])
    workbook.save('feedback.xlsx')
    age_entry.delete(0, 'end')
    mood_entry.delete(0, 'end')
    
    # подтверждаем сохранение
    confirmation_label.config(text='Saved successfully!')

# гуи
root = tk.Tk()
root.title('Feedback')

# выпадающий список для выбора пола
gender_label = tk.Label(root, text='Gender')
gender_label.grid(row=0, column=0, padx=5, pady=5)
gender_var = tk.StringVar()
gender_choices = ['Male', 'Female']
gender_dropdown = tk.OptionMenu(root, gender_var, *gender_choices)
gender_dropdown.grid(row=0, column=1, padx=5, pady=5)

# текстовое поле для ввода возраста
age_label = tk.Label(root, text='Age')
age_label.grid(row=1, column=0, padx=5, pady=5)
age_entry = tk.Entry(root)
age_entry.grid(row=1, column=1, padx=5, pady=5)

# текстовое поле для ввода настроения
mood_label = tk.Label(root, text='Mood')
mood_label.grid(row=2, column=0, padx=5, pady=5)
mood_entry = tk.Entry(root)
mood_entry.grid(row=2, column=1, padx=5, pady=5)

# кнопка
save_button = tk.Button(root, text='Confirm', command=save_feedback)
save_button.grid(row=3, column=0, padx=5, pady=5)

# метку для подтверждения сохранения
confirmation_label = tk.Label(root, text='')
confirmation_label.grid(row=3, column=1, padx=5, pady=5)

root.mainloop()
