import os
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from docxtpl import DocxTemplate, RichText

# Настройка внешнего вида приложения
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор документов по шаблонам")
        self._template = None

        self.template_button = ctk.CTkButton(root, text="Выбрать шаблон", command=self.choose_template)
        self.template_button.pack(pady=10)

        # Основной контейнер для полей ввода с прокруткой
        self.scrollable_frame = ctk.CTkScrollableFrame(root, width=400, height=300)
        self.scrollable_frame.pack(pady=(0, 10), fill="both", expand=True)
        no_data_label = ctk.CTkLabel(self.scrollable_frame, text="Нет данных",
                                     font=ctk.CTkFont(size=30),
                                     text_color="gray",
                                     anchor="center")
        no_data_label.pack(pady=130, padx=10, fill="both", expand=True)

        self.entries = {}
        self.save_frame = ctk.CTkFrame(root)
        self.save_frame.pack(fill="both", expand=True, padx=10)

        self.auto_open_var = ctk.BooleanVar(value=True)
        self.auto_open_checkbox = ctk.CTkCheckBox(self.save_frame,
                                                  text="Автоматически открыть документ",
                                                  variable=self.auto_open_var)
        self.auto_open_checkbox.grid(row=0, column=0, columnspan=1, sticky="we", pady=5, padx=5)

        self.highlight_var = ctk.BooleanVar(value=True)
        self.highlight_checkbox = ctk.CTkCheckBox(self.save_frame,
                                                  text="Выделить подставленный текст",
                                                  variable=self.highlight_var)
        self.highlight_checkbox.grid(row=1, column=0, columnspan=1, sticky="we", pady=5, padx=5)

        self.save_path_entry = ctk.CTkEntry(self.save_frame,
                                            placeholder_text="Путь для сохранения",
                                            width=250,
                                            height=30)
        self.save_path_entry.grid(row=2, column=0, padx=(5, 10), pady=5, sticky="we")

        self.save_path_button = ctk.CTkButton(self.save_frame,
                                              text="Выбрать путь",
                                              command=self.choose_save_path)
        self.save_path_button.grid(row=2, column=1, pady=5, padx=(0, 5))

        # Контейнер для быстрого заполнения всех полей
        self.fill_frame = ctk.CTkFrame(root, corner_radius=10, fg_color="transparent")
        self.fill_frame.pack(fill="x", padx=10)

        self.fill_all_entry = ctk.CTkEntry(self.fill_frame,
                                           placeholder_text="Введите текст для всех полей",
                                           width=250,
                                           height=30)
        self.fill_all_entry.grid(row=0, column=0, padx=(0, 10), pady=5, sticky="we")

        self.fill_all_button = ctk.CTkButton(self.fill_frame,
                                             text="Заполнить все поля",
                                             command=self.fill_all_fields)
        self.fill_all_button.grid(row=0, column=1, pady=5)

        self.export_button = ctk.CTkButton(root,
                                           text="Создать документ",
                                           command=self.export)
        self.export_button.pack(fill="x", pady=(5, 10), padx=10)

    def choose_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if file_path:
            self._template = DocxTemplate(file_path)
            self.load_template_variables()

    def load_template_variables(self):
        # Очистка предыдущих полей перед загрузкой новых
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        if self._template:
            variables = self._template.get_undeclared_template_variables()
            for variable in variables:
                label = ctk.CTkLabel(self.scrollable_frame, text=variable)
                label.pack(anchor="w", padx=10, pady=(5, 0))
                entry = ctk.CTkEntry(self.scrollable_frame)
                entry.pack(fill="x", padx=10, pady=5)
                self.entries[variable] = entry

    def fill_all_fields(self):
        text = self.fill_all_entry.get().strip()
        for entry in self.entries.values():
            entry.delete(0, "end")
            entry.insert(0, text)

    def choose_save_path(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
        if file_path:
            self.save_path_entry.delete(0, "end")
            self.save_path_entry.insert(0, file_path)

    def export(self):
        if not self._template:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите шаблон")
            return

        # Проверка заполнения всех полей и сбор данных
        context = {}
        missing_fields = False
        for variable, entry in self.entries.items():
            value = entry.get().strip()
            if not value:
                entry.configure(border_color="red")
                missing_fields = True
            else:
                entry.configure(border_color="")
                context[variable] = value

        if missing_fields:
            messagebox.showwarning("Внимание", "Пожалуйста, заполните все поля")
            return

        save_directory = self.save_path_entry.get().strip()
        if not save_directory:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите путь для сохранения")
            return

        save_path = Path(save_directory)

        # Применение форматирования к тексту и генерация документа
        highlight = self.highlight_var.get()
        for key, value in context.items():
            rich_text = RichText()
            if highlight:
                rich_text.add(text=value, highlight='#0018f9', underline=True)
            else:
                rich_text.add(text=value)
            context[key] = rich_text

        self._template.render(context)
        self._template.save(save_path)
        messagebox.showinfo("Готово", f"Документ создан: {save_path}")

        if self.auto_open_var.get():
            os.system(f'\"{save_path}\"')


if __name__ == '__main__':
    ctk_root = ctk.CTk()
    app = MainWindow(ctk_root)
    ctk_root.mainloop()
