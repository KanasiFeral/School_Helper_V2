using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using School_Helper_Version_2._0.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO; //Ссылка на Excel компоненты

namespace School_Helper_Version_2._0.Forms
{
    public partial class Classnaya : Form
    {
        public BindingSource binSource;

        public Exports ExportsTo = new Exports();
        public ConnectorAccess conAccess;
        public Classnaya(ConnectorAccess ClassConSQL)
        {
            InitializeComponent();
            this.conAccess = ClassConSQL;
        }

        public int Check_Button;

        //Загрузка формы
        private void Classnaya_Load(object sender, EventArgs e)
        {
            conAccess.QueryToDataGrid("SELECT * FROM Teachers", dataGridClassnaya, NavigatorClassnaya, "Классный руководитель");
            dataGridClassnaya.ReadOnly = true;
            dataGridClassnaya.AllowUserToAddRows = false;
            dataGridClassnaya.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridClassnaya.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridClassnaya.MultiSelect = false;
            binSource = conAccess.binSourceClassnaya;

            //Переименовываем названия столбцов с системного на русский язык
            try
            {
                if (dataGridClassnaya.RowCount != 0)
                {
                    dataGridClassnaya.Columns[0].Visible = false;
                    dataGridClassnaya.Columns[1].HeaderText = "Имя";
                    dataGridClassnaya.Columns[2].HeaderText = "Фамилия";
                    dataGridClassnaya.Columns[3].HeaderText = "Отчество";
                    dataGridClassnaya.Columns[4].HeaderText = "Должность";
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        //Процедура очистки текста
        public void ClearText()
        {
            textAddImyaPrepod.Clear();
            textAddFamPrepod.Clear();
            textAddOtchPrepod.Clear();
            textAddDisPrepod.Clear();
        }

        //Клик по таблице
        private void dataGridClassnaya_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                textIdPrepod.DataBindings.Add(new Binding("Text", binSource, "ID_Teacher"));
                textImyaPrepod.DataBindings.Add(new Binding("Text", binSource, "Imya"));
                textFamPrepod.DataBindings.Add(new Binding("Text", binSource, "Familiya"));
                textOtchPrepod.DataBindings.Add(new Binding("Text", binSource, "Otchestvo"));
                textDisPrepod.DataBindings.Add(new Binding("Text", binSource, "Doljnost"));
            }
            catch //Чтобы работало при навигации туда сюда, не обрабатываю исключение, гы :D
            { }
        }

        //Кнопка "Удалить запись"
        private void buttonDeleteRecordClassnaya_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Teachers") == false)
            {
                MessageBox.Show("Все строки были удалены из базы", "Ошибка удаления!");
            }
            else
            {
                //Определяем индекс выбранной строки
                int i = dataGridClassnaya.CurrentRow.Index;
                string id_Prepod = Convert.ToString(dataGridClassnaya[0, i].Value);
                //Удаление строки
                conAccess.QueryToBool("DELETE FROM Teachers WHERE ID_Teacher = " + id_Prepod);
                //Зачем здесь эта строка? Во славу Сатане конечно :3
                binSource.RemoveAt(i);
                conAccess.QueryToDataGrid("SELECT * FROM Teachers", dataGridClassnaya, NavigatorClassnaya, "Классный руководитель");
            }
        }

        //Кнопка "Очистка"
        private void buttonClearPrepod_Click(object sender, EventArgs e)
        {
            ClearText();
        }

        //Кнопка "Добавить/Изменить запись"
        private void buttonAddRecordPrepod_Click(object sender, EventArgs e)
        {
            int ID_Teacher = 0;

            if (textAddImyaPrepod.Text.Equals("") ||
                textAddFamPrepod.Text.Equals("") || textAddOtchPrepod.Text.Equals("") ||
                textAddDisPrepod.Text.Equals(""))
            {
                MessageBox.Show("Не все поля введены", "Ошибка!");
            }
            else
            {
                if (Check_Button == 0) //Была нажата кнопка "Добавить"
                {
                    if (conAccess.QueryToBool("SELECT * FROM Teachers") == true)
                    {
                        string ID = conAccess.AgregateQueryToDataGrid("SELECT MAX(ID_Teacher) FROM Teachers");
                        try
                        {
                            ID_Teacher = Convert.ToInt32(ID);
                            ID_Teacher++;
                        }
                        catch (Exception exc)
                        {
                            MessageBox.Show(exc.Message);
                        }
                    }
                    else
                    {
                        ID_Teacher = 1;
                    }

                    string queryString = "INSERT INTO Teachers VALUES (" + ID_Teacher + ",'" +
                        textAddImyaPrepod.Text + "','" + textAddFamPrepod.Text + "','" + textAddOtchPrepod.Text + "','" +
                        textAddDisPrepod.Text + "')";
                    conAccess.QueryToBool(queryString);
                    conAccess.QueryToDataGrid("SELECT * FROM Teachers", dataGridClassnaya, NavigatorClassnaya, "Классный руководитель");
                    ClearText();
                }
                else
                {
                    //Определяем индекс выбранной строки
                    int i = dataGridClassnaya.CurrentRow.Index;
                    //Забор значения из 0 столбца i-тый строки
                    string id_prepod = Convert.ToString(dataGridClassnaya[0, i].Value);
                    string queryString = "UPDATE Teachers SET Imya = '"
                        + textAddImyaPrepod.Text + "', Familiya = '"
                        + textAddFamPrepod.Text + "', Otchestvo = '"
                        + textAddOtchPrepod.Text + "', Doljnost = '"
                        + textAddDisPrepod.Text + "' WHERE ID_Teacher = " + id_prepod;
                    conAccess.QueryToBool(queryString);
                    conAccess.QueryToDataGrid("SELECT * FROM Teachers", dataGridClassnaya, NavigatorClassnaya, "Классный руководитель");
                    ClearText();
                    panelPrepod.Visible = false;
                    buttonAddPrepod.Enabled = true;
                    buttonEditPrepod.Text = "Изменить";
                }
            }
        }

        //Кнопка "Добавить"
        private void buttonAddPrepod_Click(object sender, EventArgs e)
        {
            if (buttonAddPrepod.Text == "Добавить")
            {
                Check_Button = 0;
                buttonEditPrepod.Enabled = false;
                panelPrepod.Visible = true;
                buttonAddPrepod.Text = "Скрыть";
                ClearText();
                label7.Text = "Добавление нового преподователя";
            }
            else
            {
                Check_Button = 2;
                panelPrepod.Visible = false;
                buttonEditPrepod.Enabled = true;
                buttonAddPrepod.Text = "Добавить";
                ClearText();
            }
        }

        //Кнопка "Изменить"
        private void buttonEditPrepod_Click(object sender, EventArgs e)
        {
            if (buttonEditPrepod.Text == "Изменить")
            {
                ClearText();
                int x = dataGridClassnaya.CurrentRow.Index;
                //Забираем значение ячейки
                //textAddIdPrepod.Text = Convert.ToString(dataGridClassnaya[0, x].Value);
                textAddImyaPrepod.Text = Convert.ToString(dataGridClassnaya[1, x].Value);
                textAddFamPrepod.Text = Convert.ToString(dataGridClassnaya[2, x].Value);
                textAddOtchPrepod.Text = Convert.ToString(dataGridClassnaya[3, x].Value);
                textAddDisPrepod.Text = Convert.ToString(dataGridClassnaya[4, x].Value);

                Check_Button = 1;
                buttonAddPrepod.Enabled = false;
                panelPrepod.Visible = true;
                buttonEditPrepod.Text = "Скрыть";
                label7.Text = "Изменение данных преподователя";
            }
            else
            {
                Check_Button = 2;
                panelPrepod.Visible = false;
                buttonAddPrepod.Enabled = true;
                buttonEditPrepod.Text = "Изменить";
                ClearText();
            }
        }

        //Кнопка "Сортировка по убыванию"
        private void buttonSortMinusPrepod_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            if (comboBoxSortPrepod.Text.Equals(""))
            {
                MessageBox.Show("Выберите критерий сортировки", "Ошибка!");
            }
            else
            {
                int x = comboBoxSortPrepod.SelectedIndex;
                main.Sort_Minus(dataGridClassnaya, x);
            }
        }

        //Кнопка "Сортировка по возрастанию"
        private void buttonSortPlusPrepod_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            if (comboBoxSortPrepod.Text.Equals(""))
            {
                MessageBox.Show("Выберите критерий сортировки", "Ошибка!");
            }
            else
            {
                int x = comboBoxSortPrepod.SelectedIndex;
                main.Sort_Plus(dataGridClassnaya, x);
            }
        }

        //Кнопка "Поиск по таблице"
        private void buttonSearchComp_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            main.search_datagrid(dataGridClassnaya, textSearchPrepod);
        }

        //Кнопка "Очистка"
        private void buttonClearSearchComp_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            main.clear_datagrid(dataGridClassnaya);
        }

        //Кнопка "Экспорт в Excel"
        private void buttonExportExcelPrepod_Click(object sender, EventArgs e)
        {
            ExportsTo.ExportToExcel(dataGridClassnaya);
        }

        //Кнопка "Фильтрация данных"
        private void buttonFilter_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Teachers") == false)
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
            else
            {
                string Name_Column = "";
                if (comboBoxSortPrepod.SelectedIndex == 0)
                {
                    MessageBox.Show("Фильтрация не работает с полями целочисленных типов(Int)!", "Предупреждение!");
                }
                else if (comboBoxSortPrepod.SelectedIndex == 1) { Name_Column = "Imya"; }
                else if (comboBoxSortPrepod.SelectedIndex == 2) { Name_Column = "Familiya"; }
                else if (comboBoxSortPrepod.SelectedIndex == 3) { Name_Column = "Otchestvo"; }
                else if (comboBoxSortPrepod.SelectedIndex == 4) { Name_Column = "Doljnost"; }

                try
                {
                    binSource.Filter = "[" + Name_Column + "] LIKE '" + textBoxFilter.Text + "%'";
                }
                catch (Exception exp)
                {
                    binSource.Filter = "";
                    MessageBox.Show(exp.Message);
                }
            }
        }

        //Запрет на все символы кроме цифр и клавишы BackSpace
        private void textAddIdPrepod_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                e.Handled = true;
        }
    }
}
