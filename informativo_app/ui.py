# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .config import (
    CLASSIFICATION_OPTIONS,
    CONTACT_FIELD_DEFINITIONS,
    FIELD_DEFINITIONS,
    REQUIRED_FIELD_DEFINITIONS,
)
from .docx_writer import create_docx
from .email_groups import (
    EmailGroupError,
    find_email_group,
    import_groups_from_xlsx,
    load_email_groups,
    save_email_groups,
)
from .outlook_mailer import OutlookDraftError, create_outlook_draft
from .text_builder import (
    build_attachment_lines,
    build_text,
    build_word_header,
    build_word_intro_lines,
)
from .utils import clean_filename, current_date_time


APP_TITLE = "PAINEL DE INFORMATIVOS - TRANSPORTADORA GARBUIO"
COLOR_BACKGROUND = "#f3f5fa"
COLOR_SURFACE = "#ffffff"
COLOR_NAVY = "#292746"
COLOR_MUTED = "#6b6680"
COLOR_GREEN = "#00985f"
COLOR_GREEN_DARK = "#007c4d"
COLOR_MINT = "#d8eee7"
COLOR_BORDER = "#dfe5ea"
COLOR_FIELD = "#f8fafc"


class InformativoApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        icon_path = Path(__file__).resolve().parent / "assets" / "garbuio_icon.ico"
        if icon_path.exists():
            self.root.iconbitmap(str(icon_path))
        self.root.minsize(1040, 760)
        self.root.configure(bg=COLOR_BACKGROUND)

        self.entries: dict[str, ttk.Entry | ttk.Combobox] = {}
        self.email_groups = load_email_groups()
        self.operation_field: ttk.Combobox | None = None
        self.classification_vars: dict[str, tk.BooleanVar] = {}
        self.selected_classification = "INFORMATIVO"
        self.open_after_save = tk.BooleanVar(value=False)

        self._configure_style()
        self._build_layout()
        self.update_preview()

    def _configure_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("App.TFrame", background=COLOR_BACKGROUND)
        style.configure("Card.TFrame", background=COLOR_SURFACE, relief="flat")
        style.configure("Header.TFrame", background=COLOR_SURFACE, relief="flat")
        style.configure("Controls.TFrame", background=COLOR_SURFACE)

        style.configure(
            "HeaderTitle.TLabel",
            background=COLOR_SURFACE,
            foreground=COLOR_NAVY,
            font=("Segoe UI", 18, "bold"),
        )
        style.configure(
            "HeaderSubtitle.TLabel",
            background=COLOR_SURFACE,
            foreground=COLOR_MUTED,
            font=("Segoe UI", 10),
        )
        style.configure(
            "CardTitle.TLabel",
            background=COLOR_SURFACE,
            foreground=COLOR_NAVY,
            font=("Segoe UI", 12, "bold"),
        )
        style.configure(
            "FieldLabel.TLabel",
            background=COLOR_SURFACE,
            foreground=COLOR_NAVY,
            font=("Segoe UI", 9, "bold"),
        )
        style.configure(
            "Status.TLabel",
            background=COLOR_SURFACE,
            foreground=COLOR_GREEN_DARK,
            font=("Segoe UI", 9, "bold"),
        )
        style.configure(
            "Footer.TLabel",
            background=COLOR_BACKGROUND,
            foreground=COLOR_MUTED,
            font=("Segoe UI", 8),
        )
        style.configure(
            "TEntry",
            fieldbackground=COLOR_FIELD,
            foreground=COLOR_NAVY,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            padding=(10, 7),
        )
        style.configure(
            "Primary.TButton",
            background=COLOR_GREEN,
            foreground="#ffffff",
            bordercolor=COLOR_GREEN,
            focusthickness=0,
            focuscolor=COLOR_GREEN,
            font=("Segoe UI", 10, "bold"),
            padding=(14, 9),
        )
        style.map(
            "Primary.TButton",
            background=[("pressed", COLOR_GREEN_DARK), ("active", COLOR_GREEN_DARK)],
            foreground=[("pressed", "#ffffff"), ("active", "#ffffff")],
        )
        style.configure(
            "Secondary.TButton",
            background=COLOR_MINT,
            foreground=COLOR_GREEN_DARK,
            bordercolor=COLOR_MINT,
            focusthickness=0,
            focuscolor=COLOR_MINT,
            font=("Segoe UI", 10, "bold"),
            padding=(14, 9),
        )
        style.map(
            "Secondary.TButton",
            background=[("pressed", "#c6e4da"), ("active", "#c6e4da")],
            foreground=[("pressed", COLOR_GREEN_DARK), ("active", COLOR_GREEN_DARK)],
        )
        style.configure(
            "Option.TCheckbutton",
            background=COLOR_SURFACE,
            foreground=COLOR_NAVY,
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Option.TCheckbutton",
            background=[("active", COLOR_SURFACE)],
            foreground=[("active", COLOR_GREEN_DARK)],
        )
        style.configure(
            "TCheckbutton",
            background=COLOR_SURFACE,
            foreground=COLOR_NAVY,
            font=("Segoe UI", 10),
        )
        style.map("TCheckbutton", background=[("active", COLOR_SURFACE)])

    def _add_entry(self, parent: ttk.Frame, row: int, label: str, key: str) -> None:
        ttk.Label(parent, text=f"{label}:", style="FieldLabel.TLabel").grid(
            row=row, column=0, sticky="w", pady=6
        )
        entry = ttk.Entry(parent, font=("Segoe UI", 10))
        entry.grid(row=row, column=1, sticky="ew", pady=6)
        entry.bind("<KeyRelease>", lambda _event: self.update_preview())
        self.entries[key] = entry

    def _add_operation_selector(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="Operação:", style="FieldLabel.TLabel").grid(
            row=row, column=0, sticky="w", pady=6
        )
        field = ttk.Combobox(
            parent,
            values=self._operation_names(),
            font=("Segoe UI", 10),
        )
        field.grid(row=row, column=1, sticky="ew", pady=6)
        field.bind("<KeyRelease>", lambda _event: self.update_preview())
        field.bind("<<ComboboxSelected>>", lambda _event: self.update_preview())
        self.operation_field = field
        self.entries["operacao"] = field

    def _add_classification_selector(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="Classificação:", style="FieldLabel.TLabel").grid(
            row=row, column=0, sticky="nw", pady=6
        )

        options_frame = ttk.Frame(parent, style="Controls.TFrame")
        options_frame.grid(row=row, column=1, sticky="ew", pady=(2, 8))
        for index, option in enumerate(CLASSIFICATION_OPTIONS):
            variable = tk.BooleanVar(value=option == self.selected_classification)
            self.classification_vars[option] = variable
            ttk.Checkbutton(
                options_frame,
                text=option,
                variable=variable,
                style="Option.TCheckbutton",
                command=lambda selected=option: self.select_classification(selected),
            ).grid(row=index // 3, column=index % 3, sticky="w", padx=(0, 18), pady=4)

    def select_classification(self, selected: str) -> None:
        if not self.classification_vars[selected].get():
            self.classification_vars[selected].set(True)

        self.selected_classification = selected
        for option, variable in self.classification_vars.items():
            variable.set(option == selected)

        self.update_preview()

    def _build_layout(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        shell = ttk.Frame(self.root, padding=18, style="App.TFrame")
        shell.grid(row=0, column=0, sticky="nsew")
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        header = ttk.Frame(shell, padding=(22, 18), style="Header.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text=APP_TITLE, style="HeaderTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text="Preencha, copie para WhatsApp e gere o Word padronizado em poucos cliques.",
            style="HeaderSubtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(5, 0))

        content = ttk.Frame(shell, style="App.TFrame")
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=4)
        content.columnconfigure(1, weight=5)
        content.rowconfigure(0, weight=1)

        form_card = ttk.Frame(content, padding=18, style="Card.TFrame")
        form_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        form_card.columnconfigure(1, weight=1)

        ttk.Label(form_card, text="Dados do Informativo", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12)
        )

        current_row = 1
        self._add_classification_selector(form_card, current_row)
        current_row += 1

        for label, key in FIELD_DEFINITIONS:
            if key == "operacao":
                self._add_operation_selector(form_card, current_row)
            else:
                self._add_entry(form_card, current_row, label, key)
            current_row += 1

        self.entries["data_hora"].insert(0, current_date_time())

        for label, key in CONTACT_FIELD_DEFINITIONS:
            self._add_entry(form_card, current_row, label, key)
            current_row += 1

        ttk.Label(form_card, text="Descrição:", style="FieldLabel.TLabel").grid(
            row=current_row, column=0, sticky="nw", pady=6
        )
        self.description = tk.Text(
            form_card,
            height=11,
            wrap="word",
            font=("Segoe UI", 10),
            relief="solid",
            bd=1,
            bg=COLOR_FIELD,
            fg=COLOR_NAVY,
            insertbackground=COLOR_GREEN_DARK,
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_GREEN,
            padx=8,
            pady=8,
        )
        self.description.grid(row=current_row, column=1, sticky="nsew", pady=6)
        self.description.bind("<KeyRelease>", lambda _event: self.update_preview())
        form_card.rowconfigure(current_row, weight=1)

        preview_card = ttk.Frame(content, padding=18, style="Card.TFrame")
        preview_card.grid(row=0, column=1, sticky="nsew")
        preview_card.columnconfigure(0, weight=1)
        preview_card.rowconfigure(1, weight=1)

        ttk.Label(preview_card, text="Prévia para WhatsApp", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 12)
        )

        self.preview = tk.Text(
            preview_card,
            height=18,
            wrap="word",
            font=("Consolas", 10),
            relief="solid",
            bd=1,
            bg="#fbfcfd",
            fg=COLOR_NAVY,
            insertbackground=COLOR_GREEN_DARK,
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_GREEN,
            padx=8,
            pady=8,
            state="disabled",
        )
        self.preview.grid(row=1, column=0, sticky="nsew")

        buttons = ttk.Frame(preview_card, style="Controls.TFrame")
        buttons.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        buttons.columnconfigure(4, weight=1)

        ttk.Button(
            buttons,
            text="Copiar para WhatsApp",
            style="Primary.TButton",
            command=self.copy_whatsapp_text,
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))

        ttk.Button(
            buttons,
            text="Gerar Word",
            style="Primary.TButton",
            command=self.save_word_file,
        ).grid(row=0, column=1, sticky="w", padx=(0, 8))

        ttk.Button(
            buttons,
            text="Criar e-mail no Outlook",
            style="Primary.TButton",
            command=self.create_email_draft,
        ).grid(row=0, column=2, sticky="w", padx=(0, 8))

        ttk.Button(buttons, text="Limpar", style="Secondary.TButton", command=self.clear_form).grid(
            row=0, column=3, sticky="w", padx=(0, 8)
        )

        ttk.Button(
            buttons,
            text="Importar grupos de e-mail",
            style="Secondary.TButton",
            command=self.import_email_groups,
        ).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(10, 0))

        ttk.Checkbutton(
            buttons,
            text="Abrir Word depois de gerar",
            variable=self.open_after_save,
        ).grid(row=1, column=1, columnspan=3, sticky="w", pady=(10, 0))

        if self.email_groups:
            ttk.Label(
                buttons,
                text=f"{len(self.email_groups)} grupos carregados",
                style="Status.TLabel",
            ).grid(row=1, column=4, sticky="e", pady=(10, 0))

        self.status = ttk.Label(preview_card, text="", style="Status.TLabel")
        self.status.grid(row=3, column=0, sticky="w", pady=(12, 0))
        footer = ttk.Frame(shell, style="App.TFrame")
        footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        footer.columnconfigure(0, weight=1)

        ttk.Label(
            footer,
            text="Desenvolvido por Abiézer W. Gonçalez",
            style="Footer.TLabel",
        ).grid(row=0, column=0, sticky="e")

    def _operation_names(self) -> list[str]:
        return [
            str(group.get("operacao", "")).strip()
            for group in self.email_groups
            if str(group.get("operacao", "")).strip()
        ]

    def _update_operation_values(self) -> None:
        if self.operation_field is not None:
            self.operation_field.configure(values=self._operation_names())

    def import_email_groups(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Importar grupos de e-mail",
            filetypes=[("Planilha Excel", "*.xlsx")],
        )
        if not file_path:
            return

        try:
            groups = import_groups_from_xlsx(file_path)
            save_email_groups(groups, file_path)
        except EmailGroupError as error:
            messagebox.showerror("Erro ao importar grupos", str(error))
            return

        self.email_groups = groups
        self._update_operation_values()
        self.status.configure(text=f"{len(groups)} grupos de e-mail importados.")
        messagebox.showinfo(
            "Grupos importados",
            f"{len(groups)} operações foram importadas com sucesso.",
        )

    def create_email_draft(self) -> None:
        if not self.confirm_when_missing():
            return

        if not self.email_groups:
            messagebox.showwarning(
                "Grupos não importados",
                "Importe a planilha de grupos de e-mail antes de criar o rascunho.",
            )
            return

        data = self.get_data()
        group = find_email_group(self.email_groups, data.get("operacao", ""))
        if not group:
            messagebox.showwarning(
                "Operação sem grupo",
                "Não encontrei grupo de e-mail para a operação selecionada.",
            )
            return

        try:
            create_outlook_draft(data, group)
        except OutlookDraftError as error:
            messagebox.showerror("Erro ao criar e-mail", str(error))
            return

        self.status.configure(text="Rascunho criado no Outlook.")

    def get_data(self) -> dict[str, str]:
        data = {key: entry.get() for key, entry in self.entries.items()}
        data["classificacao"] = self.selected_classification
        data["descricao"] = self.description.get("1.0", "end").strip()
        return data

    def missing_fields(self) -> list[str]:
        data = self.get_data()
        missing = [
            label
            for label, key in REQUIRED_FIELD_DEFINITIONS
            if not data.get(key, "").strip()
        ]
        if not data.get("descricao", "").strip():
            missing.append("Descrição")
        return missing

    def confirm_when_missing(self) -> bool:
        missing = self.missing_fields()
        if not missing:
            return True

        return messagebox.askyesno(
            "Campos vazios",
            "Alguns campos ainda estão vazios:\n\n"
            + "\n".join(f"- {field}" for field in missing)
            + "\n\nDeseja continuar mesmo assim?",
        )

    def update_preview(self) -> None:
        text = build_text(self.get_data(), include_request_line=True)
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", text)
        self.preview.configure(state="disabled")

    def copy_whatsapp_text(self) -> None:
        if not self.confirm_when_missing():
            return

        text = build_text(self.get_data(), include_request_line=True)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        self.status.configure(text="Texto copiado. Agora é só colar no WhatsApp.")

    def save_word_file(self) -> None:
        if not self.confirm_when_missing():
            return

        data = self.get_data()
        base_name = clean_filename(data.get("operacao", "")) or "informativo"
        default_name = f"{base_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"

        file_path = filedialog.asksaveasfilename(
            title="Salvar informativo em Word",
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Documento Word", "*.docx")],
        )
        if not file_path:
            return

        try:
            create_docx(
                file_path,
                build_text(data, include_request_line=False),
                build_word_header(data),
                build_word_intro_lines(data),
                build_attachment_lines(data),
            )
        except Exception as error:
            messagebox.showerror(
                "Erro ao gerar Word",
                f"Não foi possível gerar o arquivo Word.\n\nDetalhe: {error}",
            )
            return

        self.status.configure(text=f"Arquivo Word gerado: {file_path}")

        if self.open_after_save:
            try:
                os.startfile(str(Path(file_path).resolve()))
            except Exception:
                messagebox.showinfo(
                    "Arquivo gerado",
                    "O Word foi gerado, mas não consegui abrir o arquivo automaticamente.",
                )
        else:
            messagebox.showinfo("Arquivo gerado", "O arquivo Word foi gerado com sucesso.")

    def clear_form(self) -> None:
        for key, entry in self.entries.items():
            if isinstance(entry, ttk.Combobox):
                entry.set("")
            else:
                entry.delete(0, "end")
            if key == "data_hora":
                entry.insert(0, current_date_time())

        self.select_classification("INFORMATIVO")
        self.description.delete("1.0", "end")
        self.status.configure(text="")
        self.update_preview()


def main() -> None:
    root = tk.Tk()
    InformativoApp(root)
    root.mainloop()
