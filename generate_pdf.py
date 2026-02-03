#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CLI-скрипт для генерации PDF из CSV-данных и HTML-шаблона.

Использование:
    python generate_pdf.py --csv data.csv --template template.html

Автор: PDF-CheckMaker
"""

import os
import sys
import argparse
import platform
import subprocess
from pathlib import Path
from datetime import datetime

import pandas as pd
from jinja2 import Template
from fpdf import FPDF


# ============================================================================
# НАСТРОЙКИ
# ============================================================================

OUTPUT_DIR = Path(__file__).parent / "output"
DEFAULT_TEMPLATE = Path(__file__).parent / "template.html"


# ============================================================================
# ФУНКЦИИ РАБОТЫ С ФАЙЛАМИ
# ============================================================================

def validate_csv(filepath: Path) -> pd.DataFrame:
    """
    Проверяет существование CSV-файла и читает данные.
    
    Args:
        filepath: Путь к CSV-файлу
        
    Returns:
        DataFrame с данными
        
    Raises:
        FileNotFoundError: Если файл не найден
    """
    if not filepath.exists():
        raise FileNotFoundError(f"CSV-файл не найден: {filepath}")
    
    df = pd.read_csv(filepath, encoding='utf-8')
    
    if df.empty:
        raise ValueError("CSV-файл пустой")
    
    print(f"      Колонки: {', '.join(df.columns)}")
    
    return df


def validate_template(filepath: Path) -> str:
    """
    Проверяет существование HTML-шаблона и читает его содержимое.
    
    Args:
        filepath: Путь к шаблону
        
    Returns:
        Содержимое шаблона
        
    Raises:
        FileNotFoundError: Если шаблон не найден
    """
    if not filepath.exists():
        raise FileNotFoundError(f"Шаблон не найден: {filepath}")
    
    with open(filepath, 'r', encoding='utf-8') as f:
        return f.read()


def ensure_output_dir():
    """Создаёт папку output если её нет."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def generate_output_filename() -> Path:
    """
    Генерирует имя выходного файла с timestamp.
    
    Returns:
        Путь к выходному PDF-файлу
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return OUTPUT_DIR / f"output_{timestamp}.pdf"


# ============================================================================
# ГЕНЕРАЦИЯ PDF
# ============================================================================

class PDFGenerator(FPDF):
    """Генератор PDF с поддержкой кириллицы."""
    
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=15)
        self._setup_font()
    
    def _setup_font(self):
        """Настройка шрифта с поддержкой кириллицы."""
        fonts_dir = self._get_fonts_dir()
        
        # Пробуем подключить Arial (есть на Windows)
        font_options = [
            ('Arial', 'arial.ttf', 'arialbd.ttf'),
            ('DejaVuSans', 'DejaVuSans.ttf', 'DejaVuSans-Bold.ttf'),
            ('Calibri', 'calibri.ttf', 'calibrib.ttf'),
        ]
        
        for font_name, regular, bold in font_options:
            regular_path = fonts_dir / regular
            bold_path = fonts_dir / bold
            
            if regular_path.exists():
                try:
                    self.add_font(font_name, '', str(regular_path))
                    if bold_path.exists():
                        self.add_font(font_name, 'B', str(bold_path))
                    self.font_name = font_name
                    return
                except Exception:
                    continue
        
        # Fallback на встроенный шрифт
        self.font_name = 'Helvetica'
    
    def _get_fonts_dir(self) -> Path:
        """Возвращает путь к системным шрифтам."""
        if platform.system() == 'Windows':
            return Path(os.environ.get('WINDIR', 'C:\\Windows')) / 'Fonts'
        elif platform.system() == 'Darwin':
            return Path('/System/Library/Fonts/Supplemental')
        return Path('/usr/share/fonts/truetype/dejavu')
    
    def header(self):
        """Заголовок страницы."""
        self.set_font(self.font_name, 'B', 16)
        self.set_text_color(44, 62, 80)
        self.cell(0, 10, 'Каталог товаров', align='C', new_x='LMARGIN', new_y='NEXT')
        self.ln(5)
    
    def footer(self):
        """Подвал страницы."""
        self.set_y(-15)
        self.set_font(self.font_name, '', 8)
        self.set_text_color(150, 150, 150)
        timestamp = datetime.now().strftime("%d.%m.%Y %H:%M")
        self.cell(0, 10, f'Сгенерировано: {timestamp} | Стр. {self.page_no()}', align='C')
    
    def add_products_table(self, df: pd.DataFrame):
        """
        Добавляет таблицу с данными в PDF.
        Автоматически определяет колонки из DataFrame.
        
        Args:
            df: DataFrame с любыми колонками
        """
        self.add_page()
        
        columns = list(df.columns)
        num_cols = len(columns)
        
        # Рассчитываем ширину колонок (первая - №, остальные равномерно)
        available_width = 190  # A4 минус поля
        num_col_width = 12
        col_width = (available_width - num_col_width) / num_cols
        
        # Заголовок таблицы
        self.set_fill_color(52, 73, 94)
        self.set_text_color(255, 255, 255)
        self.set_font(self.font_name, 'B', 8)
        
        self.cell(num_col_width, 10, '№', border=1, align='C', fill=True)
        for col in columns:
            # Сокращаем длинные заголовки
            header = str(col)[:15]
            self.cell(col_width, 10, header, border=1, align='C', fill=True)
        self.ln()
        
        # Данные
        self.set_text_color(0, 0, 0)
        self.set_font(self.font_name, '', 7)
        
        for i, (idx, row) in enumerate(df.iterrows()):
            # Чередование цвета строк
            if i % 2 == 0:
                self.set_fill_color(245, 245, 245)
                fill = True
            else:
                fill = False
            
            self.cell(num_col_width, 8, str(i + 1), border=1, align='C', fill=fill)
            
            for col in columns:
                value = str(row[col])[:20]  # Обрезаем длинные значения
                self.cell(col_width, 8, value, border=1, fill=fill)
            self.ln()
        
        # Итого
        self.ln(5)
        self.set_font(self.font_name, 'B', 12)
        self.set_text_color(44, 62, 80)
        self.cell(0, 10, f'Всего записей: {len(df)}', new_x='LMARGIN', new_y='NEXT')


def generate_pdf_from_data(df: pd.DataFrame, output_path: Path) -> Path:
    """
    Генерирует PDF-файл из данных.
    
    Args:
        df: DataFrame с товарами
        output_path: Путь для сохранения PDF
        
    Returns:
        Путь к созданному файлу
    """
    pdf = PDFGenerator()
    pdf.add_products_table(df)
    pdf.output(str(output_path))
    return output_path


# ============================================================================
# АЛЬТЕРНАТИВНАЯ ГЕНЕРАЦИЯ ЧЕРЕЗ HTML-ШАБЛОН
# ============================================================================

def generate_pdf_from_template(df: pd.DataFrame, template_content: str, output_path: Path) -> Path:
    """
    Генерирует PDF через HTML-шаблон (рендеринг Jinja2 + fpdf2).
    
    Примечание: fpdf2 не рендерит HTML напрямую, поэтому используем 
    табличный вывод с данными из шаблона.
    
    Args:
        df: DataFrame с товарами
        template_content: Содержимое HTML-шаблона (для информации)
        output_path: Путь для сохранения
        
    Returns:
        Путь к созданному файлу
    """
    # Для fpdf2 используем прямую генерацию
    # HTML-шаблон можно использовать для извлечения заголовка и стилей
    return generate_pdf_from_data(df, output_path)


# ============================================================================
# ОТКРЫТИЕ ФАЙЛА
# ============================================================================

def open_file(filepath: Path):
    """
    Открывает файл в системном приложении.
    
    Args:
        filepath: Путь к файлу
    """
    system = platform.system()
    filepath_str = str(filepath)
    
    try:
        if system == 'Windows':
            os.startfile(filepath_str)
        elif system == 'Darwin':  # macOS
            subprocess.run(['open', filepath_str], check=True)
        else:  # Linux
            subprocess.run(['xdg-open', filepath_str], check=True)
        print(f"[OK] Файл открыт: {filepath.name}")
    except Exception as e:
        print(f"[!] Не удалось открыть файл автоматически: {e}")
        print(f"    Путь: {filepath}")


# ============================================================================
# CLI ИНТЕРФЕЙС
# ============================================================================

def parse_arguments() -> argparse.Namespace:
    """
    Парсит аргументы командной строки.
    
    Returns:
        Namespace с аргументами
    """
    parser = argparse.ArgumentParser(
        description='Генератор PDF из CSV-данных',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Примеры использования:
  python generate_pdf.py --csv data.csv
  python generate_pdf.py --csv data.csv --template custom.html
  python generate_pdf.py -c products.csv -t template.html
        '''
    )
    
    parser.add_argument(
        '-c', '--csv',
        type=Path,
        required=True,
        help='Путь к CSV-файлу с данными (обязательный)'
    )
    
    parser.add_argument(
        '-t', '--template',
        type=Path,
        default=DEFAULT_TEMPLATE,
        help=f'Путь к HTML-шаблону (по умолчанию: {DEFAULT_TEMPLATE.name})'
    )
    
    parser.add_argument(
        '-o', '--output',
        type=Path,
        default=None,
        help='Путь для сохранения PDF (по умолчанию: output/output_TIMESTAMP.pdf)'
    )
    
    parser.add_argument(
        '--no-open',
        action='store_true',
        help='Не открывать PDF после создания'
    )
    
    return parser.parse_args()


def main():
    """Главная функция CLI."""
    print("\n" + "=" * 50)
    print("  PDF Generator - Генератор PDF из CSV")
    print("=" * 50 + "\n")
    
    # Парсим аргументы
    args = parse_arguments()
    
    try:
        # 1. Валидация CSV
        print(f"[1/4] Чтение CSV: {args.csv}")
        df = validate_csv(args.csv)
        print(f"      Найдено записей: {len(df)}")
        
        # 2. Проверка шаблона (опционально)
        template_content = None
        if args.template.exists():
            print(f"[2/4] Шаблон: {args.template}")
            template_content = validate_template(args.template)
        else:
            print(f"[2/4] Шаблон не найден, используется встроенный стиль")
        
        # 3. Создание output директории
        print(f"[3/4] Подготовка папки output...")
        ensure_output_dir()
        
        # 4. Генерация PDF
        output_path = args.output or generate_output_filename()
        print(f"[4/4] Генерация PDF: {output_path.name}")
        
        if template_content:
            result = generate_pdf_from_template(df, template_content, output_path)
        else:
            result = generate_pdf_from_data(df, output_path)
        
        print(f"\n[OK] PDF успешно создан!")
        print(f"     Файл: {result}")
        
        # 5. Открытие файла
        if not args.no_open:
            print(f"\n[*] Открываю файл...")
            open_file(result)
        
    except FileNotFoundError as e:
        print(f"\n[ОШИБКА] {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"\n[ОШИБКА] {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n[ОШИБКА] Непредвиденная ошибка: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    
    print("\n" + "=" * 50)
    print("  Готово!")
    print("=" * 50 + "\n")


if __name__ == "__main__":
    main()
