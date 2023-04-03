using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;


namespace W2E
{
    class translator
    {
        // Окно
        public FormW2E frm;
        public translator(FormW2E frm)
        {
            this.frm = frm;
        }

        /// <summary>
        /// Просто начало обработки
        /// </summary>
        /// <param name="file">Путь к файлу</param>
        public void Startup(string file)
        {
            System.Windows.Forms.TextBox tb = frm.getLogTextBox();
            tb.Text += "Начинаю обработку..." + "\r\n";

            Word.Application word = null;
            Word.Document wDoc = null;
            Excel.Application excel = null;
            Excel.Workbook eDoc = null;

            try
            {
                tb.Text += "Пытаюсь запустить MS Word..." + "\r\n";
                word = new Word.Application();
                tb.Text += "MS Word запущен" + "\r\n";
                word.Visible = false;
                tb.Text += "Пытаюсь открыть выбранный документ..." + "\r\n";
                wDoc = word.Documents.Open(file);
                tb.Text += "Выбранный документ открыт" + "\r\n";
                tb.Text += "Пытаюсь запустить MS Excel..." + "\r\n";
                excel = new Excel.Application();
                tb.Text += "MS Excel заущен" + "\r\n";
                excel.Visible = false;
                tb.Text += "Пытаюсь создать нровую книгу Excel..." + "\r\n";
                eDoc = excel.Workbooks.Add();
                tb.Text += "Создал новую книгу Excel" + "\r\n";
                tb.Text += "Добавляю новый лист..." + "\r\n";
                Excel.Worksheet worksheet = eDoc.Worksheets.Add();
                tb.Text += "Добавил новый лист" + "\r\n";
                int row = 1;
                tb.Text += "Начинаю перебирать таблицы документа MS Word..." + "\r\n";
                int id = 1;
                foreach (Word.Table table in wDoc.Tables)
                {
                    tb.Text += $"Копирую таблицу {id}..." + "\r\n";
                    table.Range.Copy();
                    tb.Text += "Скопировал" + "\r\n";
                    tb.Text += "Указываю место вставки таблицы на листе Excel..." + "\r\n";
                    Excel.Range cell = worksheet.Cells[row, 1];
                    tb.Text += "Место вставки указал, вставляю таблицу..." + "\r\n";
                    worksheet.Paste(cell);
                    tb.Text += "Вставил таблицу из MS Word в MS Excel" + "\r\n";
                    tb.Text += "Обновляю место вставки таблицы на листе Excel..." + "\r\n";
                    row = worksheet.UsedRange.Rows.Count + 2;
                    //worksheet.UsedRange.ColumnWidth = 100;
                    id++;
                }
                tb.Text += "Перебрал все таблицы" + "\r\n";
                tb.Text += "Подгоняю размеры таблиц..." + "\r\n";
                worksheet.UsedRange.Columns.AutoFit();
                tb.Text += "Подогнал" + "\r\n";
                tb.Text += "Сохраняю книгу MS Excel..." + "\r\n";
                eDoc.SaveAs(file + ".xlsx");
                tb.Text += "Сохранил" + "\r\n";
                tb.Text += "закрываю книгу MS Excel..." + "\r\n";
                eDoc.Close();
                tb.Text += "Закрыл" + "\r\n";
                tb.Text += "Выхожу из программы MS Excel..." + "\r\n";
                excel.Quit();
                tb.Text += "Вышел" + "\r\n";
                tb.Text += "Закрываю документ MS Word..." + "\r\n";
                wDoc.Close();
                tb.Text += "Закрыл" + "\r\n";
                tb.Text += "Выхожу из программы MS Word ..." + "\r\n";
                word.Quit();
                tb.Text += "Вышел" + "\r\n";
                tb.Text += "Обработка завершена!" + "\r\n";
            }
            catch (Exception ex)
            {
                if (eDoc != null)   eDoc.Close();
                if (excel != null)  excel.Quit();
                if (wDoc != null)   wDoc.Close();
                if (word != null)   word.Quit();

                tb.Text += ex.ToString();
            }
            
        }
    }
}
