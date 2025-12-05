import os
import sys
import shutil #файлы по папкам
import subprocess
from datetime import date
from pathlib import Path
from typing import Dict, Any, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import yaml
from docxtpl import DocxTemplate

DEFAULT_FIELDS = [
    ("project_name", "Название проекта"),
    ("goal", "Цель проекта"),
    ("scope", "Область применения"),
    ("feature_list", "Ключевые функции"),
    ("ui_requirements", "Требования к интерфейсу"),
    ("other_requirements", "Прочие нефункциональные требования"),
    ("start", "Дата начала работ"),
    ("deadline", "Срок окончания работ"),
    ("customer", "Заказчик"),
    ("performers", "Исполнители"),
]


def render_docx(template_path: Path, output_docx: Path, context: Dict[str, Any]) -> None:
    tpl = DocxTemplate(str(template_path)) #объект шаблона
    ctx = dict(context)
    ctx.setdefault("today", date.today().strftime("%d.%m.%Y")) #чтобы вручную не вводить (может, переделать?)
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    tpl.render(ctx)
    tpl.save(str(output_docx))


def try_convert_to_pdf(input_docx: Path, output_pdf: Path) -> bool:
    #docx2pdf
    try:
        from docx2pdf import convert as docx2pdf_convert #только конвертер
        output_pdf.parent.mkdir(parents=True, exist_ok=True) #вроде в докс проверяем, но НЕ ТРОГАТЬ
        temp_out_dir = output_pdf.parent / f"__tmp_pdf_{output_pdf.stem}" #временная папка, при втором запуске ломается
        temp_out_dir.mkdir(exist_ok=True) #UPD: больше не ломается
        docx2pdf_convert(str(input_docx), str(temp_out_dir))
        candidate = temp_out_dir / (input_docx.stem + ".pdf")
        if candidate.exists():
            shutil.move(str(candidate), str(output_pdf))
            shutil.rmtree(temp_out_dir, ignore_errors=True)
            return True
        shutil.rmtree(temp_out_dir, ignore_errors=True)
    except Exception:
        pass

    #LibreOffice
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            output_pdf.parent.mkdir(parents=True, exist_ok=True)
            cmd = [
                soffice,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(output_pdf.parent),
                str(input_docx),
            ]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            produced = input_docx.with_suffix(".pdf")
            if produced.exists():
                produced.rename(output_pdf)
                return True
        except Exception:
            pass

    return False


#Вспомогательные функции для OS
def open_file(path: Path) -> None:
    if not path.exists():
        messagebox.showerror("Ошибка", f"Файл не найден:\n{path}")
        return
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)])
    else:
        subprocess.run(["xdg-open", str(path)])


def open_folder(path: Path) -> None:
    if not path.exists():
        messagebox.showerror("Ошибка", f"Папка не найдена:\n{path}")
        return
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)])
    else:
        subprocess.run(["xdg-open", str(path)])

#основа для Tkinter
class TzApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DEAL Генератор технических заданий")
        self.geometry("900x600")
        self.minsize(900, 600)
        style = ttk.Style()
        style.configure("Big.TButton", font=("Arial", 16))

        #состояние приложения
        self.template_path: Optional[Path] = None
        self.output_dir: Path = Path("output")
        self.base_name: str = "tz"
        self.form_data: Dict[str, str] = {k: "" for k, _ in DEFAULT_FIELDS}
        self.generated_docx: Optional[Path] = None
        self.generated_pdf: Optional[Path] = None

        #контейнер для экранов
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        #создание экранов
        self.frames = {}
        for FrameCls in (MainMenuFrame, FormFrame, PreviewFrame, ResultFrame):
            frame = FrameCls(parent=container, controller=self)
            frame.grid(row=0, column=0, sticky="nsew")
            self.frames[FrameCls.__name__] = frame

        self.show_frame("MainMenuFrame")

    def show_frame(self, name: str):
        frame = self.frames[name]
        frame.tkraise()
        if hasattr(frame, "on_show"):
            frame.on_show()  # type: ignore[call-arg]

    def reset_state(self):
        self.form_data = {k: "" for k, _ in DEFAULT_FIELDS}
        self.generated_docx = None
        self.generated_pdf = None


#фреймы
class MainMenuFrame(ttk.Frame):
    def __init__(self, parent, controller: TzApp):
        super().__init__(parent)
        self.controller = controller

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        #Заголовок
        title = ttk.Label(
            self,
            text="DEAL Генератор технических заданий",
            font=("Arial", 24, "bold"),
            anchor="center",
        )
        title.grid(row=0, column=0, pady=(40, 20))

        # центральная панель с кнопками
        center = ttk.Frame(self)
        center.grid(row=1, column=0)
        for i in range(3):
            center.rowconfigure(i, pad=10)

        btn_new = ttk.Button(center, text="Создать новое ТЗ", style="Big.TButton", command=self.on_new_tz)

        btn_new.grid(row=0, column=0, ipady=10, ipadx=73, pady=10)

        btn_template = ttk.Button(
            center, text="Загрузить шаблон DOCX", style="Big.TButton", command=self.on_load_template
        )
        btn_template.grid(row=1, column=0, ipady=10, ipadx=40, pady=10)

        #подпись о текущем шаблоне
        self.template_label = ttk.Label(
            self,
            text="Текущий шаблон: [по умолчанию: templates/ts_template.docx]",
            font=("Arial", 15),
        )
        self.template_label.grid(row=2, column=0, pady=(40, 20))

        #кнопка Выход в правом нижнем углу
        bottom = ttk.Frame(self)
        bottom.grid(row=3, column=0, sticky="se", padx=20, pady=20)
        btn_exit = ttk.Button(bottom, text="Выход", style="Big.TButton", command=self.controller.destroy)
        btn_exit.pack()

    def on_show(self):
        #обновить надпись о шаблоне
        if self.controller.template_path:
            text = f"Текущий шаблон: {self.controller.template_path}"
        else:
            text = "Текущий шаблон: [по умолчанию: templates/ts_template.docx]"
        self.template_label.config(text=text)

    def on_load_template(self):
        file_path = filedialog.askopenfilename(
            title="Выберите шаблон DOCX",
            filetypes=[("DOCX файлы", "*.docx")],
        )
        if file_path:
            self.controller.template_path = Path(file_path)
            self.on_show()

    @staticmethod
    def load_prefill(path: Path) -> Dict[str, Any]:
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        return data or {}

    def on_new_tz(self):
        self.controller.reset_state()

        prefill_path = Path("prefill.yaml")
        if prefill_path.exists():
            try:
                self.controller.form_data.update(self.load_prefill(prefill_path))
            except Exception as e:
                messagebox.showerror("Ошибка префилла", str(e))

        self.controller.show_frame("FormFrame")


class FormFrame(ttk.Frame):
    def __init__(self, parent, controller: TzApp):
        super().__init__(parent)
        self.controller = controller

        self.columnconfigure(0, weight=1)

        title = ttk.Label(
            self,
            text="Введите данные технического задания",
            font=("Arial", 20, "bold"),
            anchor="center",
        )
        title.grid(row=0, column=0, pady=(30, 20))

        #серый блок формы
        form_container = ttk.Frame(self)
        form_container.grid(row=1, column=0, pady=10)
        form_container.columnconfigure(1, weight=1)

        self.entries: Dict[str, tk.Text] = {}

        for row, (key, label_text) in enumerate(DEFAULT_FIELDS):
            lbl = ttk.Label(form_container, text=f"{label_text}:")
            lbl.grid(row=row, column=0, sticky="w", pady=3, padx=(20, 10))

            #многострочный Text вместо Entry, чтобы было ближе к макету
            txt = tk.Text(form_container, height=1, width=60)
            txt.grid(row=row, column=1, sticky="ew", pady=3, padx=(0, 20))
            self.entries[key] = txt

        self.rowconfigure(0, weight=0)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=0)
        self.columnconfigure(0, weight=1)

        #кнопка Далее >> внизу справа
        bottom = ttk.Frame(self)
        bottom.grid(row=2, column=0, sticky="se", padx=20, pady=20)

        btn_next = ttk.Button(bottom, text="Далее >>", command=self.on_next)
        btn_next.grid(row=0, column=0, sticky="se")


    def on_show(self):
        #подставить уже введённые данные обратно в поля
        for key, widget in self.entries.items():
            widget.delete("1.0", "end")
            value = self.controller.form_data.get(key, "")
            if value:
                widget.insert("1.0", value)


    def on_next(self):
        #собрать данные из полей
        for key, widget in self.entries.items():
            value = widget.get("1.0", "end").strip()
            self.controller.form_data[key] = value

        self.controller.show_frame("PreviewFrame")


class PreviewFrame(ttk.Frame):
    def __init__(self, parent, controller: TzApp):
        super().__init__(parent)
        self.controller = controller

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        title = ttk.Label(
            self,
            text="Предпросмотр",
            font=("Arial", 20, "bold"),
            anchor="center",
        )
        title.grid(row=0, column=0, pady=(30, 10), padx=20)

        #рамка с текстом (большой скролл-блок)
        frame_preview = ttk.Frame(self, borderwidth=2, relief="solid")
        frame_preview.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        frame_preview.rowconfigure(0, weight=1)
        frame_preview.columnconfigure(0, weight=1)

        self.text_preview = tk.Text(frame_preview, wrap="word")
        self.text_preview.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(
            frame_preview, orient="vertical", command=self.text_preview.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.text_preview.configure(yscrollcommand=scrollbar.set)

        #нижняя панель с кнопками
        bottom = ttk.Frame(self)
        bottom.grid(row=2, column=0, pady=20)

        btn_edit = ttk.Button(bottom, text="<< Редактировать", command=self.on_edit)
        btn_edit.grid(row=0, column=0, padx=10, ipadx=10, ipady=5)

        btn_docx = ttk.Button(bottom, text="Сгенерировать DOCX", command=self.on_generate_docx)
        btn_docx.grid(row=0, column=1, padx=10, ipadx=10, ipady=5)

        btn_pdf = ttk.Button(bottom, text="Сгенерировать PDF", command=self.on_generate_pdf)
        btn_pdf.grid(row=0, column=2, padx=10, ipadx=10, ipady=5)

    def on_show(self):
        #обновляем текст предпросмотра
        self.text_preview.delete("1.0", "end")
        lines = []
        labels = dict(DEFAULT_FIELDS)
        for key, _ in DEFAULT_FIELDS:
            label = labels[key]
            value = self.controller.form_data.get(key, "") or "[не заполнено]"
            lines.append(f"{label}: {value}")
        self.text_preview.insert("1.0", "\n".join(lines))
        self.text_preview.mark_set("insert", "1.0")

    def _ensure_template(self) -> Optional[Path]:
        #если не выбран шаблон: используем дефолтный
        if self.controller.template_path is None:
            default = Path("templates/ts_template.docx")
            if not default.exists():
                messagebox.showerror(
                    "Ошибка",
                    f"Шаблон не выбран и файл по умолчанию не найден:\n{default}",
                )
                return None
            self.controller.template_path = default
        return self.controller.template_path

    def _generate_docx_internal(self) -> Optional[Path]:
        template = self._ensure_template()
        if template is None:
            return None

        out_dir = self.controller.output_dir
        out_docx = out_dir / f"{self.controller.base_name}.docx"
        try:
            render_docx(template, out_docx, self.controller.form_data)
        except Exception as e:
            messagebox.showerror("Ошибка при генерации DOCX", str(e))
            return None
        self.controller.generated_docx = out_docx
        return out_docx

    def on_edit(self):
        self.controller.show_frame("FormFrame")

    def on_generate_docx(self):
        docx_path = self._generate_docx_internal()
        if docx_path is None:
            return
        messagebox.showinfo("Готово", f"DOCX сгенерирован:\n{docx_path}")
        self.controller.show_frame("ResultFrame")

    def on_generate_pdf(self):
        #есть DOCX?
        docx_path = self.controller.generated_docx or self._generate_docx_internal()
        if docx_path is None:
            return

        out_pdf = self.controller.output_dir / f"{self.controller.base_name}.pdf"
        ok = try_convert_to_pdf(docx_path, out_pdf)
        if not ok:
            messagebox.showwarning(
                "PDF не сгенерирован",
                "Не удалось автоматически сконвертировать DOCX в PDF.\n"
                "Проверьте установку docx2pdf или LibreOffice.",
            )
            self.controller.generated_pdf = None
        else:
            self.controller.generated_pdf = out_pdf
            messagebox.showinfo("Готово", f"PDF сгенерирован:\n{out_pdf}")
        self.controller.show_frame("ResultFrame")



#Результат

class ResultFrame(ttk.Frame):
    def __init__(self, parent, controller: TzApp):
        super().__init__(parent)
        self.controller = controller

        self.columnconfigure(0, weight=1)

        title = ttk.Label(
            self,
            text="Техническое задание создано",
            font=("Arial", 20, "bold"),
            anchor="center",
        )
        title.grid(row=0, column=0, pady=(40, 30))

        center = ttk.Frame(self)
        center.grid(row=1, column=0, pady=80)

        btn_open_docx = ttk.Button(center, text="Открыть DOCX", style="Big.TButton", command=self.on_open_docx)
        btn_open_docx.grid(row=0, column=0, pady=5, ipadx=20, ipady=5)

        btn_open_pdf = ttk.Button(center, text="Открыть PDF", style="Big.TButton", command=self.on_open_pdf)
        btn_open_pdf.grid(row=1, column=0, pady=5, ipadx=20, ipady=5)

        btn_open_folder = ttk.Button(center, text="Открыть папку", style="Big.TButton", command=self.on_open_folder)
        btn_open_folder.grid(row=2, column=0, pady=5, ipadx=20, ipady=5)

        bottom = ttk.Frame(self)
        bottom.grid(row=3, column=0, sticky="se", padx=20, pady=20)
        btn_main = ttk.Button(bottom, text="В главное меню", style="Big.TButton", command=self.on_main_menu)
        btn_main.pack()

    def on_open_docx(self):
        if not self.controller.generated_docx:
            messagebox.showwarning("Нет файла", "DOCX ещё не был сгенерирован.")
            return
        open_file(self.controller.generated_docx)

    def on_open_pdf(self):
        if not self.controller.generated_pdf:
            messagebox.showwarning("Нет файла", "PDF ещё не был сгенерирован.")
            return
        open_file(self.controller.generated_pdf)

    def on_open_folder(self):
        open_folder(self.controller.output_dir)

    def on_main_menu(self):
        self.controller.show_frame("MainMenuFrame")

if __name__ == "__main__":
    app = TzApp()
    app.mainloop()