using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;

namespace School_Helper_Version_2._0
{
    public class Exports
    {
        //Процедура экспорта таблицы в Excel
        public void ExportToExcel(DataGridView dGV)
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel;
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            SaveFileDialog SaveDialog = new SaveFileDialog();
            SaveDialog.Filter = "Файл Excel|*.XLSX;*.XLS";
            SaveDialog.ShowDialog();
            try
            {
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Определяем последний запущенный процесс Excel
                System.Diagnostics.Process excelProc = System.Diagnostics.Process.GetProcessesByName("EXCEL").Last();
                ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                for (int i = 0; i < dGV.Rows.Count; i++)
                {
                    for (int j = 0; j < dGV.Columns.Count; j++)
                    {
                        ObjExcel.Cells[i + 1, j + 1] = dGV.Rows[i].Cells[j].Value.ToString();
                    }
                }
                ObjWorkBook.SaveAs(SaveDialog.FileName);
                ObjWorkBook.Close();
                ObjWorkBook = null;
                ObjWorkSheet = null;
                ObjExcel = null;
                excelProc.Kill();

                MessageBox.Show("Экспорт был произведен", "Готово!");
            }
            catch (Exception ex)
            {

                System.Diagnostics.Process excelProc = System.Diagnostics.Process.GetProcessesByName("EXCEL").Last();
                excelProc.Kill();
                MessageBox.Show("Отмена печати документа!", "Отмена!");
                MessageBox.Show("Ошибка: " + ex.Message, "Ошибка");
            }
        }
    }
}
