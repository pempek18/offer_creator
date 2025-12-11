import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import json
import os

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import reportlab.rl_config
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


def normalize_encoding(text):
    """
    Normalizuje kodowanie tekstu, próbując naprawić uszkodzone znaki.
    Konwertuje tekst do UTF-8, próbując różnych kodowań jeśli potrzeba.
    """
    if not text or not isinstance(text, str):
        return text
    
    # Jeśli tekst już jest poprawny i nie zawiera znaków zastępczych, zwróć go
    try:
        text.encode('utf-8')
        # Sprawdź czy nie zawiera znaków zastępczych (kwadraty, pytajniki itp.)
        if '\ufffd' not in text and '' not in text and '■' not in text:
            return text
    except UnicodeEncodeError:
        pass
    
    # Jeśli tekst zawiera znaki zastępcze, spróbuj naprawić
    # Najpierw spróbuj traktować jako Windows-1250 (typowe dla polskich systemów)
    try:
        # Jeśli tekst zawiera znaki zastępcze, spróbuj różnych kodowań
        if isinstance(text, bytes):
            # Jeśli to bytes, spróbuj zdekodować
            for encoding in ['utf-8', 'windows-1250', 'iso-8859-2', 'cp1250']:
                try:
                    decoded = text.decode(encoding)
                    if '\ufffd' not in decoded and '' not in decoded:
                        return decoded
                except (UnicodeDecodeError, UnicodeEncodeError):
                    continue
        else:
            # Jeśli to string z uszkodzonymi znakami, spróbuj naprawić
            # Najpierw zakoduj jako latin1 (zachowuje bajty) i zdekoduj jako windows-1250
            try:
                encoded = text.encode('latin1', errors='ignore')
                for encoding in ['windows-1250', 'iso-8859-2', 'cp1250']:
                    try:
                        decoded = encoded.decode(encoding, errors='ignore')
                        if '\ufffd' not in decoded and '' not in decoded and '■' not in decoded:
                            # Sprawdź czy zawiera polskie znaki (znak że naprawa zadziałała)
                            if any(c in decoded for c in 'ąćęłńóśźżĄĆĘŁŃÓŚŹŻ'):
                                return decoded
                    except (UnicodeDecodeError, UnicodeEncodeError):
                        continue
            except:
                pass
    except:
        pass
    
    # Jeśli nic nie zadziałało, zwróć oryginalny tekst
    return text


def fix_string_encoding(text):
    """
    Naprawia kodowanie pojedynczego stringa.
    """
    if isinstance(text, dict):
        return {k: fix_string_encoding(v) for k, v in text.items()}
    elif isinstance(text, list):
        return [fix_string_encoding(item) for item in text]
    elif isinstance(text, str):
        return normalize_encoding(text)
    else:
        return text


class OfferCreatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tworzenie Ofert")
        self.root.geometry("1000x700")
        
        # Dane firmy (stałe)
        self.company_data = {
            "name": "",
            "address": "",
            "city": "",
            "postal_code": "",
            "nip": "",
            "phone": "",
            "email": "",
            "bank_account": ""
        }
        
        # Odbiorcy (zmienne)
        self.recipients = []
        
        # Pozycje oferty
        self.offer_items = []
        
        self.setup_ui()
        self.load_company_data()
    
    def setup_ui(self):
        # Główny kontener z zakładkami
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Zakładka 1: Dane firmy
        company_frame = ttk.Frame(notebook)
        notebook.add(company_frame, text="Dane Firmy")
        self.setup_company_tab(company_frame)
        
        # Zakładka 2: Odbiorcy
        recipients_frame = ttk.Frame(notebook)
        notebook.add(recipients_frame, text="Odbiorcy")
        self.setup_recipients_tab(recipients_frame)
        
        # Zakładka 3: Tworzenie Oferty
        offer_frame = ttk.Frame(notebook)
        notebook.add(offer_frame, text="Tworzenie Oferty")
        self.setup_offer_tab(offer_frame)
    
    def setup_company_tab(self, parent):
        # Nagłówek
        header = ttk.Label(parent, text="Stałe Dane Firmy", font=("Arial", 16, "bold"))
        header.pack(pady=10)
        
        # Ramka z polami
        form_frame = ttk.LabelFrame(parent, text="Informacje o firmie", padding=15)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Pola formularza
        fields = [
            ("Nazwa firmy:", "name"),
            ("Adres:", "address"),
            ("Miasto:", "city"),
            ("Kod pocztowy:", "postal_code"),
            ("NIP:", "nip"),
            ("Telefon:", "phone"),
            ("Email:", "email"),
            ("Numer konta bankowego:", "bank_account")
        ]
        
        self.company_entries = {}
        for i, (label, key) in enumerate(fields):
            row_frame = ttk.Frame(form_frame)
            row_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(row_frame, text=label, width=20).pack(side=tk.LEFT)
            entry = ttk.Entry(row_frame, width=50)
            entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
            self.company_entries[key] = entry
        
        # Przyciski
        button_frame = ttk.Frame(parent)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Zapisz Dane Firmy", 
                  command=self.save_company_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Wczytaj Dane Firmy", 
                  command=self.load_company_data).pack(side=tk.LEFT, padx=5)
    
    def setup_recipients_tab(self, parent):
        # Nagłówek
        header = ttk.Label(parent, text="Zarządzanie Odbiorcami", font=("Arial", 16, "bold"))
        header.pack(pady=10)
        
        # Lista odbiorców
        list_frame = ttk.LabelFrame(parent, text="Lista Odbiorców", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Treeview dla listy
        columns = ("Nazwa", "Adres", "Miasto", "NIP")
        self.recipients_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)
        
        for col in columns:
            self.recipients_tree.heading(col, text=col)
            self.recipients_tree.column(col, width=150)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.recipients_tree.yview)
        self.recipients_tree.configure(yscrollcommand=scrollbar.set)
        
        self.recipients_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Formularz dodawania/edycji
        form_frame = ttk.LabelFrame(parent, text="Dodaj/Edytuj Odbiorcę", padding=15)
        form_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.recipient_entries = {}
        fields = [
            ("Nazwa:", "name"),
            ("Adres:", "address"),
            ("Miasto:", "city"),
            ("Kod pocztowy:", "postal_code"),
            ("NIP:", "nip"),
            ("Telefon:", "phone"),
            ("Email:", "email")
        ]
        
        for i, (label, key) in enumerate(fields):
            row_frame = ttk.Frame(form_frame)
            row_frame.pack(fill=tk.X, pady=3)
            
            ttk.Label(row_frame, text=label, width=15).pack(side=tk.LEFT)
            entry = ttk.Entry(row_frame, width=40)
            entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            self.recipient_entries[key] = entry
        
        # Przyciski
        button_frame = ttk.Frame(parent)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Dodaj Odbiorcę", 
                  command=self.add_recipient).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Edytuj Odbiorcę", 
                  command=self.edit_recipient).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Usuń Odbiorcę", 
                  command=self.delete_recipient).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Wyczyść Formularz", 
                  command=self.clear_recipient_form).pack(side=tk.LEFT, padx=5)
        
        self.recipients_tree.bind("<Double-1>", self.on_recipient_select)
        self.refresh_recipients_list()
    
    def setup_offer_tab(self, parent):
        # Wybór odbiorcy
        recipient_frame = ttk.LabelFrame(parent, text="Wybór Odbiorcy", padding=10)
        recipient_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Label(recipient_frame, text="Odbiorca:").pack(side=tk.LEFT, padx=5)
        self.selected_recipient = tk.StringVar()
        self.recipient_combo = ttk.Combobox(recipient_frame, textvariable=self.selected_recipient, 
                                           width=50, state="readonly")
        self.recipient_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.recipient_combo.bind("<<ComboboxSelected>>", self.update_recipient_combo)
        
        # Pozycje oferty
        items_frame = ttk.LabelFrame(parent, text="Pozycje Oferty", padding=10)
        items_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Treeview dla pozycji
        columns = ("Lp", "Nazwa", "Ilosc", "Cena jednostkowa", "Wartosc")
        self.items_tree = ttk.Treeview(items_frame, columns=columns, show="headings", height=8)
        
        for col in columns:
            self.items_tree.heading(col, text=col)
            if col == "Lp":
                self.items_tree.column(col, width=50)
            elif col == "Nazwa":
                self.items_tree.column(col, width=300)
            else:
                self.items_tree.column(col, width=120)
        
        scrollbar_items = ttk.Scrollbar(items_frame, orient=tk.VERTICAL, 
                                       command=self.items_tree.yview)
        self.items_tree.configure(yscrollcommand=scrollbar_items.set)
        
        self.items_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_items.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Formularz pozycji
        item_form_frame = ttk.LabelFrame(parent, text="Dodaj/Edytuj Pozycję", padding=15)
        item_form_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.item_entries = {}
        item_fields = [
            ("Nazwa pozycji:", "name"),
            ("Ilosc:", "quantity"),
            ("Cena jednostkowa:", "unit_price")
        ]
        
        for label, key in item_fields:
            row_frame = ttk.Frame(item_form_frame)
            row_frame.pack(fill=tk.X, pady=3)
            
            ttk.Label(row_frame, text=label, width=20).pack(side=tk.LEFT)
            entry = ttk.Entry(row_frame, width=40)
            entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            self.item_entries[key] = entry
        
        # Przyciski pozycji
        item_button_frame = ttk.Frame(parent)
        item_button_frame.pack(pady=5)
        
        ttk.Button(item_button_frame, text="Dodaj Pozycję", 
                  command=self.add_item).pack(side=tk.LEFT, padx=5)
        ttk.Button(item_button_frame, text="Edytuj Pozycję", 
                  command=self.edit_item).pack(side=tk.LEFT, padx=5)
        ttk.Button(item_button_frame, text="Usuń Pozycję", 
                  command=self.delete_item).pack(side=tk.LEFT, padx=5)
        ttk.Button(item_button_frame, text="Wyczyść Formularz", 
                  command=self.clear_item_form).pack(side=tk.LEFT, padx=5)
        
        # Suma
        total_frame = ttk.Frame(parent)
        total_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.total_label = ttk.Label(total_frame, text="Suma: 0.00 PLN", 
                                    font=("Arial", 12, "bold"))
        self.total_label.pack(side=tk.RIGHT)
        
        # Przyciski oferty
        offer_button_frame = ttk.Frame(parent)
        offer_button_frame.pack(pady=10)
        
        ttk.Button(offer_button_frame, text="Generuj Ofertę (TXT)", 
                  command=self.generate_offer_txt).pack(side=tk.LEFT, padx=5)
        ttk.Button(offer_button_frame, text="Generuj Ofertę (PDF)", 
                  command=self.generate_offer_pdf).pack(side=tk.LEFT, padx=5)
        ttk.Button(offer_button_frame, text="Wczytaj Ofertę (JSON)", 
                  command=self.load_offer_json).pack(side=tk.LEFT, padx=5)
        ttk.Button(offer_button_frame, text="Zapisz Ofertę (JSON)", 
                  command=self.save_offer_json).pack(side=tk.LEFT, padx=5)
        ttk.Button(offer_button_frame, text="Wyczyść Ofertę", 
                  command=self.clear_offer).pack(side=tk.LEFT, padx=5)
        
        self.items_tree.bind("<Double-1>", self.on_item_select)
        self.update_recipient_combo()
    
    def save_company_data(self):
        for key, entry in self.company_entries.items():
            value = entry.get()
            # Normalizuj kodowanie przed zapisaniem
            self.company_data[key] = normalize_encoding(value)
        
        # Zapisz do pliku JSON
        try:
            with open("company_data.json", "w", encoding="utf-8") as f:
                json.dump(self.company_data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Sukces", "Dane firmy zostały zapisane!")
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się zapisać danych: {str(e)}")
    
    def load_company_data(self):
        # Wczytaj z pliku JSON
        if os.path.exists("company_data.json"):
            try:
                with open("company_data.json", "r", encoding="utf-8") as f:
                    self.company_data = json.load(f)
                
                # Napraw kodowanie danych
                self.company_data = fix_string_encoding(self.company_data)
                
                for key, entry in self.company_entries.items():
                    entry.delete(0, tk.END)
                    value = self.company_data.get(key, "")
                    entry.insert(0, value)
            except Exception as e:
                messagebox.showerror("Błąd", f"Nie udało się wczytać danych: {str(e)}")
        else:
            # Wypełnij puste wartości
            for key, entry in self.company_entries.items():
                entry.delete(0, tk.END)
                entry.insert(0, self.company_data.get(key, ""))
    
    def add_recipient(self):
        recipient = {}
        for key, entry in self.recipient_entries.items():
            value = entry.get()
            # Normalizuj kodowanie wszystkich wartości
            recipient[key] = normalize_encoding(value)
        
        if not recipient.get("name"):
            messagebox.showwarning("Uwaga", "Nazwa odbiorcy jest wymagana!")
            return
        
        self.recipients.append(recipient)
        self.save_recipients()
        self.refresh_recipients_list()
        self.clear_recipient_form()
        self.update_recipient_combo()
        messagebox.showinfo("Sukces", "Odbiorca został dodany!")
    
    def edit_recipient(self):
        selected = self.recipients_tree.selection()
        if not selected:
            messagebox.showwarning("Uwaga", "Wybierz odbiorcę do edycji!")
            return
        
        item = self.recipients_tree.item(selected[0])
        recipient_name = item['values'][0]
        
        # Znajdź odbiorcę
        for i, recipient in enumerate(self.recipients):
            if recipient.get("name") == recipient_name:
                # Zaktualizuj dane z normalizacją kodowania
                for key, entry in self.recipient_entries.items():
                    value = entry.get()
                    self.recipients[i][key] = normalize_encoding(value)
                
                self.save_recipients()
                self.refresh_recipients_list()
                self.clear_recipient_form()
                self.update_recipient_combo()
                messagebox.showinfo("Sukces", "Odbiorca został zaktualizowany!")
                return
    
    def delete_recipient(self):
        selected = self.recipients_tree.selection()
        if not selected:
            messagebox.showwarning("Uwaga", "Wybierz odbiorcę do usunięcia!")
            return
        
        if messagebox.askyesno("Potwierdzenie", "Czy na pewno chcesz usunąć tego odbiorcę?"):
            item = self.recipients_tree.item(selected[0])
            recipient_name = item['values'][0]
            
            self.recipients = [r for r in self.recipients if r.get("name") != recipient_name]
            self.save_recipients()
            self.refresh_recipients_list()
            self.clear_recipient_form()
            self.update_recipient_combo()
    
    def clear_recipient_form(self):
        for entry in self.recipient_entries.values():
            entry.delete(0, tk.END)
    
    def on_recipient_select(self, event):
        selected = self.recipients_tree.selection()
        if selected:
            item = self.recipients_tree.item(selected[0])
            recipient_name = item['values'][0]
            
            # Znajdź i wypełnij formularz
            for recipient in self.recipients:
                if recipient.get("name") == recipient_name:
                    for key, entry in self.recipient_entries.items():
                        entry.delete(0, tk.END)
                        entry.insert(0, recipient.get(key, ""))
                    break
    
    def refresh_recipients_list(self):
        # Wyczyść listę
        for item in self.recipients_tree.get_children():
            self.recipients_tree.delete(item)
        
        # Dodaj odbiorców
        for recipient in self.recipients:
            self.recipients_tree.insert("", tk.END, values=(
                recipient.get("name", ""),
                recipient.get("address", ""),
                recipient.get("city", ""),
                recipient.get("nip", "")
            ))
    
    def save_recipients(self):
        try:
            with open("recipients.json", "w", encoding="utf-8") as f:
                json.dump(self.recipients, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się zapisać odbiorców: {str(e)}")
    
    def load_recipients(self):
        if os.path.exists("recipients.json"):
            try:
                with open("recipients.json", "r", encoding="utf-8") as f:
                    self.recipients = json.load(f)
                
                # Napraw kodowanie danych
                self.recipients = fix_string_encoding(self.recipients)
                
                self.refresh_recipients_list()
            except Exception as e:
                messagebox.showerror("Błąd", f"Nie udało się wczytać odbiorców: {str(e)}")
    
    def update_recipient_combo(self, event=None):
        values = [r.get("name", "") for r in self.recipients]
        self.recipient_combo['values'] = values
        if values and not self.selected_recipient.get():
            self.selected_recipient.set(values[0] if values else "")
    
    def add_item(self):
        name = self.item_entries["name"].get()
        quantity_str = self.item_entries["quantity"].get()
        price_str = self.item_entries["unit_price"].get()
        
        if not name:
            messagebox.showwarning("Uwaga", "Nazwa pozycji jest wymagana!")
            return
        
        try:
            quantity = float(quantity_str.replace(",", ".")) if quantity_str else 0
            price = float(price_str.replace(",", ".")) if price_str else 0
            total = quantity * price
            
            # Normalizuj kodowanie nazwy
            name = normalize_encoding(name)
            
            item = {
                "name": name,
                "quantity": quantity,
                "unit_price": price,
                "total": total
            }
            
            self.offer_items.append(item)
            self.refresh_items_list()
            self.clear_item_form()
            self.update_total()
        except ValueError:
            messagebox.showerror("Błąd", "Nieprawidłowa wartość ilości lub ceny!")
    
    def edit_item(self):
        selected = self.items_tree.selection()
        if not selected:
            messagebox.showwarning("Uwaga", "Wybierz pozycję do edycji!")
            return
        
        item_index = int(self.items_tree.item(selected[0])['values'][0]) - 1
        
        name = self.item_entries["name"].get()
        quantity_str = self.item_entries["quantity"].get()
        price_str = self.item_entries["unit_price"].get()
        
        try:
            quantity = float(quantity_str.replace(",", ".")) if quantity_str else 0
            price = float(price_str.replace(",", ".")) if price_str else 0
            total = quantity * price
            
            # Normalizuj kodowanie nazwy
            name = normalize_encoding(name)
            
            self.offer_items[item_index] = {
                "name": name,
                "quantity": quantity,
                "unit_price": price,
                "total": total
            }
            
            self.refresh_items_list()
            self.clear_item_form()
            self.update_total()
        except (ValueError, IndexError):
            messagebox.showerror("Błąd", "Nieprawidłowa wartość lub pozycja!")
    
    def delete_item(self):
        selected = self.items_tree.selection()
        if not selected:
            messagebox.showwarning("Uwaga", "Wybierz pozycję do usunięcia!")
            return
        
        item_index = int(self.items_tree.item(selected[0])['values'][0]) - 1
        
        try:
            del self.offer_items[item_index]
            self.refresh_items_list()
            self.clear_item_form()
            self.update_total()
        except IndexError:
            messagebox.showerror("Błąd", "Nie można usunąć pozycji!")
    
    def clear_item_form(self):
        for entry in self.item_entries.values():
            entry.delete(0, tk.END)
    
    def on_item_select(self, event):
        selected = self.items_tree.selection()
        if selected:
            item_index = int(self.items_tree.item(selected[0])['values'][0]) - 1
            
            try:
                item = self.offer_items[item_index]
                self.item_entries["name"].delete(0, tk.END)
                self.item_entries["name"].insert(0, item["name"])
                self.item_entries["quantity"].delete(0, tk.END)
                self.item_entries["quantity"].insert(0, str(item["quantity"]))
                self.item_entries["unit_price"].delete(0, tk.END)
                self.item_entries["unit_price"].insert(0, str(item["unit_price"]))
            except IndexError:
                pass
    
    def refresh_items_list(self):
        # Wyczyść listę
        for item in self.items_tree.get_children():
            self.items_tree.delete(item)
        
        # Dodaj pozycje
        for i, item in enumerate(self.offer_items, 1):
            self.items_tree.insert("", tk.END, values=(
                i,
                item["name"],
                f"{item['quantity']:.2f}",
                f"{item['unit_price']:.2f} PLN",
                f"{item['total']:.2f} PLN"
            ))
    
    def update_total(self):
        total = sum(item["total"] for item in self.offer_items)
        self.total_label.config(text=f"Suma: {total:.2f} PLN")
    
    def clear_offer(self):
        self.offer_items = []
        self.refresh_items_list()
        self.clear_item_form()
        self.update_total()
        self.selected_recipient.set("")
    
    def get_selected_recipient_data(self):
        recipient_name = self.selected_recipient.get()
        for recipient in self.recipients:
            if recipient.get("name") == recipient_name:
                return recipient
        return None
    
    def generate_offer_txt(self):
        if not self.offer_items:
            messagebox.showwarning("Uwaga", "Dodaj pozycje do oferty!")
            return
        
        recipient = self.get_selected_recipient_data()
        if not recipient:
            messagebox.showwarning("Uwaga", "Wybierz odbiorcę oferty!")
            return
        
        # Wybierz miejsce zapisu
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Pliki tekstowe", "*.txt"), ("Wszystkie pliki", "*.*")],
            initialfile=f"Oferta_{recipient.get('name', '')}_{datetime.now().strftime('%Y%m%d')}.txt"
        )
        
        if not filename:
            return
        
        try:
            with open(filename, "w", encoding="utf-8") as f:
                # Nagłówek
                f.write("=" * 80 + "\n")
                f.write("OFERTA\n")
                f.write("=" * 80 + "\n\n")
                
                # Data
                offer_date = datetime.now()
                valid_until = offer_date + timedelta(days=30)  # +1 miesiąc (około 30 dni)
                f.write(f"Data: {offer_date.strftime('%d.%m.%Y')}\n")
                f.write(f"Oferta ważna do: {valid_until.strftime('%d.%m.%Y')}\n\n")
                
                # Dane firmy
                f.write("SPRZEDAWCA:\n")
                f.write("-" * 80 + "\n")
                for key, label in [
                    ("name", "Nazwa:"),
                    ("address", "Adres:"),
                    ("city", "Miasto:"),
                    ("postal_code", "Kod pocztowy:"),
                    ("nip", "NIP:"),
                    ("phone", "Telefon:"),
                    ("email", "Email:"),
                    ("bank_account", "Konto bankowe:")
                ]:
                    if self.company_data.get(key):
                        f.write(f"{label} {self.company_data[key]}\n")
                f.write("\n")
                
                # Dane odbiorcy
                f.write("ODBIORCA:\n")
                f.write("-" * 80 + "\n")
                for key, label in [
                    ("name", "Nazwa:"),
                    ("address", "Adres:"),
                    ("city", "Miasto:"),
                    ("postal_code", "Kod pocztowy:"),
                    ("nip", "NIP:"),
                    ("phone", "Telefon:"),
                    ("email", "Email:")
                ]:
                    if recipient.get(key):
                        f.write(f"{label} {recipient[key]}\n")
                f.write("\n")
                
                # Pozycje oferty
                f.write("POZYCJE OFERTY:\n")
                f.write("-" * 80 + "\n")
                f.write(f"{'Lp':<5} {'Nazwa':<40} {'Ilość':>10} {'Cena':>12} {'Wartość':>12}\n")
                f.write("-" * 80 + "\n")
                
                total = 0
                for i, item in enumerate(self.offer_items, 1):
                    f.write(f"{i:<5} {item['name']:<40} {item['quantity']:>10.2f} "
                           f"{item['unit_price']:>12.2f} PLN {item['total']:>12.2f} PLN\n")
                    total += item['total']
                
                f.write("-" * 80 + "\n")
                f.write(f"{'SUMA:':<57} {total:>12.2f} PLN\n")
                f.write("=" * 80 + "\n")
            
            messagebox.showinfo("Sukces", f"Oferta została zapisana do pliku:\n{filename}")
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się zapisać oferty: {str(e)}")
    
    def generate_offer_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Błąd", 
                "Biblioteka reportlab nie jest zainstalowana!\n\n"
                "Zainstaluj ją poleceniem:\npip install reportlab")
            return
        
        if not self.offer_items:
            messagebox.showwarning("Uwaga", "Dodaj pozycje do oferty!")
            return
        
        recipient = self.get_selected_recipient_data()
        if not recipient:
            messagebox.showwarning("Uwaga", "Wybierz odbiorcę oferty!")
            return
        
        # Wybierz miejsce zapisu
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Pliki PDF", "*.pdf"), ("Wszystkie pliki", "*.*")],
            initialfile=f"Oferta_{recipient.get('name', '')}_{datetime.now().strftime('%Y%m%d')}.pdf"
        )
        
        if not filename:
            return
        
        try:
            # Rejestruj fonty obsługujące polskie znaki
            # Próbuj użyć fontów DejaVu, jeśli są dostępne
            font_name = 'Helvetica'  # Domyślny font
            font_bold = 'Helvetica-Bold'
            
            # Próbuj zarejestrować DejaVu Sans (obsługuje polskie znaki)
            try:
                # Sprawdź dostępność fontów DejaVu w typowych lokalizacjach
                dejavu_paths = [
                    'C:/Windows/Fonts/DejaVuSans.ttf',
                    'C:/Windows/Fonts/dejavu/DejaVuSans.ttf',
                    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
                    '/usr/share/fonts/TTF/DejaVuSans.ttf',
                ]
                
                # Sprawdź również fonty Arial Unicode MS (Windows) lub Liberation Sans (Linux)
                arial_unicode_paths = [
                    'C:/Windows/Fonts/ARIALUNI.TTF',
                    'C:/Windows/Fonts/arialuni.ttf',
                ]
                
                liberation_paths = [
                    '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
                    '/usr/share/fonts/TTF/LiberationSans-Regular.ttf',
                ]
                
                dejavu_found = False
                
                # Najpierw próbuj DejaVu
                for path in dejavu_paths:
                    if os.path.exists(path):
                        try:
                            pdfmetrics.registerFont(TTFont('DejaVuSans', path))
                            bold_path = path.replace('DejaVuSans.ttf', 'DejaVuSans-Bold.ttf')
                            if os.path.exists(bold_path):
                                pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', bold_path))
                            font_name = 'DejaVuSans'
                            font_bold = 'DejaVuSans-Bold'
                            dejavu_found = True
                            break
                        except Exception as e:
                            continue
                
                # Jeśli DejaVu nie znaleziono, próbuj Arial Unicode MS
                if not dejavu_found:
                    for path in arial_unicode_paths:
                        if os.path.exists(path):
                            try:
                                pdfmetrics.registerFont(TTFont('ArialUnicode', path))
                                font_name = 'ArialUnicode'
                                font_bold = 'ArialUnicode'  # Użyj tego samego fontu dla bold
                                dejavu_found = True
                                break
                            except Exception as e:
                                continue
                
                # Jeśli nadal nie znaleziono, próbuj Liberation Sans
                if not dejavu_found:
                    for path in liberation_paths:
                        if os.path.exists(path):
                            try:
                                pdfmetrics.registerFont(TTFont('LiberationSans', path))
                                bold_path = path.replace('LiberationSans-Regular.ttf', 'LiberationSans-Bold.ttf')
                                if os.path.exists(bold_path):
                                    pdfmetrics.registerFont(TTFont('LiberationSans-Bold', bold_path))
                                font_name = 'LiberationSans'
                                font_bold = 'LiberationSans-Bold'
                                dejavu_found = True
                                break
                            except Exception as e:
                                continue
                
                # Jeśli żaden font Unicode nie został znaleziony, użyj domyślnych
                # Helvetica w reportlab ma ograniczone wsparcie dla Unicode
                if not dejavu_found:
                    font_name = 'Helvetica'
                    font_bold = 'Helvetica-Bold'
            except Exception as e:
                # W razie problemów, użyj domyślnych fontów
                font_name = 'Helvetica'
                font_bold = 'Helvetica-Bold'
            
            # Utwórz dokument PDF
            doc = SimpleDocTemplate(filename, pagesize=A4,
                                   rightMargin=20*mm, leftMargin=20*mm,
                                   topMargin=20*mm, bottomMargin=20*mm)
            
            # Kontener na elementy
            story = []
            
            # Style z fontami obsługującymi polskie znaki
            styles = getSampleStyleSheet()
            
            # Utwórz style z odpowiednimi fontami
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontName=font_bold,
                fontSize=24,
                textColor=colors.HexColor('#1a1a1a'),
                spaceAfter=30,
                alignment=TA_CENTER
            )
            
            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading2'],
                fontName=font_bold,
                fontSize=14,
                textColor=colors.HexColor('#2c3e50'),
                spaceAfter=12,
                spaceBefore=12
            )
            
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontName=font_name,
                fontSize=10
            )
            
            # Tytuł
            story.append(Paragraph("OFERTA", title_style))
            story.append(Spacer(1, 10*mm))
            
            # Data
            offer_date = datetime.now()
            valid_until = offer_date + timedelta(days=30)  # +1 miesiąc (około 30 dni)
            date_text = f"Data: {offer_date.strftime('%d.%m.%Y')}"
            valid_until_text = f"Oferta wazna do: {valid_until.strftime('%d.%m.%Y')}"
            story.append(Paragraph(date_text, normal_style))
            story.append(Paragraph(valid_until_text, normal_style))
            story.append(Spacer(1, 5*mm))
            
            # Przygotuj dane sprzedawcy jako tekst
            company_text_parts = []
            company_text_parts.append(Paragraph("<b>SPRZEDAWCA:</b>", heading_style))
            for key, label in [
                ("name", "Nazwa:"),
                ("address", "Adres:"),
                ("city", "Miasto:"),
                ("postal_code", "Kod pocztowy:"),
                ("nip", "NIP:"),
                ("phone", "Telefon:"),
                ("email", "Email:"),
                ("bank_account", "Konto bankowe:")
            ]:
                value = self.company_data.get(key, "")
                if value:
                    company_text_parts.append(Paragraph(f"<b>{label}</b> {value}", normal_style))
            
            # Przygotuj dane odbiorcy jako tekst
            recipient_text_parts = []
            recipient_text_parts.append(Paragraph("<b>ODBIORCA:</b>", heading_style))
            for key, label in [
                ("name", "Nazwa:"),
                ("address", "Adres:"),
                ("city", "Miasto:"),
                ("postal_code", "Kod pocztowy:"),
                ("nip", "NIP:"),
                ("phone", "Telefon:"),
                ("email", "Email:")
            ]:
                value = recipient.get(key, "")
                if value:
                    recipient_text_parts.append(Paragraph(f"<b>{label}</b> {value}", normal_style))
            
            # Utwórz tabelę z dwiema kolumnami (sprzedawca po lewej, odbiorca po prawej)
            # Oblicz szerokość kolumn
            page_width = A4[0] - 40*mm  # Szerokość strony minus marginesy
            col_width = (page_width - 10*mm) / 2  # Połowa szerokości minus odstęp między kolumnami
            
            # Znajdź maksymalną liczbę wierszy
            max_rows = max(len(company_text_parts), len(recipient_text_parts))
            
            # Utwórz wiersze tabeli
            side_by_side_data = []
            for i in range(max_rows):
                left_cell = company_text_parts[i] if i < len(company_text_parts) else Paragraph("", normal_style)
                right_cell = recipient_text_parts[i] if i < len(recipient_text_parts) else Paragraph("", normal_style)
                side_by_side_data.append([left_cell, right_cell])
            
            side_by_side_table = Table(side_by_side_data, colWidths=[col_width, col_width])
            side_by_side_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (0, -1), 0),
                ('RIGHTPADDING', (0, 0), (0, -1), 5*mm),
                ('LEFTPADDING', (1, 0), (1, -1), 5*mm),
                ('RIGHTPADDING', (1, 0), (1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ]))
            story.append(side_by_side_table)
            
            story.append(Spacer(1, 8*mm))
            
            # Pozycje oferty
            story.append(Paragraph("<b>POZYCJE OFERTY:</b>", heading_style))
            
            # Styl dla komórek tabeli
            cell_style = ParagraphStyle(
                'TableCell',
                parent=normal_style,
                fontName=font_name,
                fontSize=10,
                leading=12,
                textColor=colors.black
            )
            
            cell_style_bold = ParagraphStyle(
                'TableCellBold',
                parent=normal_style,
                fontName=font_bold,
                fontSize=10,
                leading=12,
                textColor=colors.black
            )
            
            # Styl dla nagłówka tabeli (biały tekst)
            header_cell_style = ParagraphStyle(
                'TableHeader',
                parent=normal_style,
                fontName=font_bold,
                fontSize=11,
                leading=13,
                textColor=colors.whitesmoke
            )
            
            # Nagłówek tabeli - użyj Paragraph dla lepszej obsługi Unicode
            items_data = [[
                Paragraph('Lp', header_cell_style),
                Paragraph('Nazwa', header_cell_style),
                Paragraph('Ilosc', header_cell_style),
                Paragraph('Cena jedn.', header_cell_style),
                Paragraph('Wartosc', header_cell_style)
            ]]
            
            total = 0
            for i, item in enumerate(self.offer_items, 1):
                # Użyj Paragraph dla nazwy (może zawierać polskie znaki)
                # Dla pozostałych pól też użyj Paragraph dla spójności
                items_data.append([
                    Paragraph(str(i), cell_style),
                    Paragraph(item['name'], cell_style),  # To jest kluczowe - nazwa może mieć polskie znaki
                    Paragraph(f"{item['quantity']:.2f}", cell_style),
                    Paragraph(f"{item['unit_price']:.2f} PLN", cell_style),
                    Paragraph(f"{item['total']:.2f} PLN", cell_style)
                ])
                total += item['total']
            
            # Wiersz sumy
            items_data.append([
                Paragraph('', cell_style),
                Paragraph('', cell_style),
                Paragraph('', cell_style),
                Paragraph('<b>SUMA:</b>', cell_style_bold),
                Paragraph(f'<b>{total:.2f} PLN</b>', cell_style_bold)
            ])
            
            items_table = Table(items_data, colWidths=[15*mm, 80*mm, 25*mm, 30*mm, 30*mm])
            items_table.setStyle(TableStyle([
                # Nagłówek
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#34495e')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                # Uwaga: FONTNAME i FONTSIZE są ignorowane gdy używamy Paragraph,
                # fonty są określone w ParagraphStyle
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('TOPPADDING', (0, 0), (-1, 0), 12),
                
                # Wiersze danych
                ('VALIGN', (0, 1), (-1, -2), 'MIDDLE'),
                ('BOTTOMPADDING', (0, 1), (-1, -2), 8),
                ('TOPPADDING', (0, 1), (-1, -2), 8),
                ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f8f9fa')]),
                
                # Wiersz sumy
                ('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#ecf0f1')),
                ('BOTTOMPADDING', (0, -1), (-1, -1), 10),
                ('TOPPADDING', (0, -1), (-1, -1), 10),
                
                # Obramowanie
                ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#bdc3c7')),
                ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor('#34495e')),
                ('LINEABOVE', (0, -1), (-1, -1), 2, colors.HexColor('#34495e')),
            ]))
            
            story.append(items_table)
            
            # Generuj PDF
            doc.build(story)
            
            messagebox.showinfo("Sukces", f"Oferta PDF została zapisana do pliku:\n{filename}")
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się wygenerować PDF: {str(e)}")
    
    def load_offer_json(self):
        # Wybierz plik do wczytania
        filename = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("Pliki JSON", "*.json"), ("Wszystkie pliki", "*.*")]
        )
        
        if not filename:
            return
        
        try:
            with open(filename, "r", encoding="utf-8") as f:
                offer_data = json.load(f)
            
            # Napraw kodowanie danych
            offer_data = fix_string_encoding(offer_data)
            
            # Sprawdź czy plik ma poprawną strukturę
            if not isinstance(offer_data, dict):
                messagebox.showerror("Błąd", "Nieprawidłowy format pliku JSON!")
                return
            
            # Wczytaj dane firmy (opcjonalnie - tylko jeśli są w pliku)
            if "company" in offer_data and offer_data["company"]:
                company_from_file = offer_data["company"]
                # Aktualizuj tylko puste pola w danych firmy
                for key in self.company_data.keys():
                    if key in company_from_file and company_from_file[key]:
                        # Aktualizuj w słowniku
                        self.company_data[key] = company_from_file[key]
                        # Aktualizuj w interfejsie (jeśli istnieje pole)
                        if key in self.company_entries:
                            self.company_entries[key].delete(0, tk.END)
                            self.company_entries[key].insert(0, company_from_file[key])
            
            # Wczytaj odbiorcę
            if "recipient" not in offer_data or not offer_data["recipient"]:
                messagebox.showerror("Błąd", "Plik nie zawiera danych odbiorcy!")
                return
            
            recipient = offer_data["recipient"]
            recipient_name = recipient.get("name", "")
            
            if not recipient_name:
                messagebox.showerror("Błąd", "Odbiorca w pliku nie ma nazwy!")
                return
            
            # Sprawdź czy odbiorca istnieje w liście
            recipient_exists = False
            for r in self.recipients:
                if r.get("name") == recipient_name:
                    # Zaktualizuj dane istniejącego odbiorcy
                    for key, value in recipient.items():
                        r[key] = value
                    recipient_exists = True
                    break
            
            # Jeśli odbiorca nie istnieje, dodaj go
            if not recipient_exists:
                self.recipients.append(recipient)
                self.save_recipients()
                self.refresh_recipients_list()
            
            # Ustaw odbiorcę w combobox
            self.update_recipient_combo()
            self.selected_recipient.set(recipient_name)
            
            # Wczytaj pozycje oferty
            if "items" not in offer_data or not offer_data["items"]:
                messagebox.showwarning("Uwaga", "Plik nie zawiera pozycji oferty!")
                self.offer_items = []
            else:
                self.offer_items = offer_data["items"]
                # Upewnij się, że każda pozycja ma obliczoną wartość total
                for item in self.offer_items:
                    if "total" not in item or item["total"] == 0:
                        quantity = item.get("quantity", 0)
                        unit_price = item.get("unit_price", 0)
                        item["total"] = quantity * unit_price
            
            # Odśwież listę pozycji i sumę
            self.refresh_items_list()
            self.update_total()
            self.clear_item_form()
            
            # Wyświetl informację o dacie oferty jeśli jest dostępna
            date_info = ""
            if "date" in offer_data:
                date_info = f"\nData oferty: {offer_data['date']}"
            
            messagebox.showinfo("Sukces", 
                f"Oferta została wczytana z pliku:\n{filename}{date_info}\n\n"
                f"Odbiorca: {recipient_name}\n"
                f"Pozycji: {len(self.offer_items)}")
            
        except json.JSONDecodeError:
            messagebox.showerror("Błąd", "Nieprawidłowy format pliku JSON!")
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się wczytać oferty: {str(e)}")
    
    def save_offer_json(self):
        if not self.offer_items:
            messagebox.showwarning("Uwaga", "Dodaj pozycje do oferty!")
            return
        
        recipient = self.get_selected_recipient_data()
        if not recipient:
            messagebox.showwarning("Uwaga", "Wybierz odbiorcę oferty!")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("Pliki JSON", "*.json"), ("Wszystkie pliki", "*.*")],
            initialfile=f"Oferta_{recipient.get('name', '')}_{datetime.now().strftime('%Y%m%d')}.json"
        )
        
        if not filename:
            return
        
        offer_data = {
            "date": datetime.now().strftime("%Y-%m-%d"),
            "company": self.company_data,
            "recipient": recipient,
            "items": self.offer_items,
            "total": sum(item["total"] for item in self.offer_items)
        }
        
        try:
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(offer_data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Sukces", f"Oferta została zapisana do pliku:\n{filename}")
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się zapisać oferty: {str(e)}")


def main():
    root = tk.Tk()
    app = OfferCreatorApp(root)
    
    # Wczytaj odbiorców przy starcie
    app.load_recipients()
    
    root.mainloop()


if __name__ == "__main__":
    main()

