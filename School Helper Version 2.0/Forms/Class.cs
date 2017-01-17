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
    public partial class Class : Form
    {
        public BindingSource binSource;

        public Exports ExportsTo = new Exports();
        public ConnectorAccess conAccess;
        public Class(ConnectorAccess ClassConSQL)
        {
            InitializeComponent();
            this.conAccess = ClassConSQL;
        }

        public int Check_Button;

        //Загрузка формы
        private void Class_Load(object sender, EventArgs e)
        {
            conAccess.QueryToDataGrid("SELECT * FROM Classes", dataGridClass, NavigatorClass, "Класс");
            dataGridClass.ReadOnly = true;
            dataGridClass.AllowUserToAddRows = false;
            dataGridClass.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridClass.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridClass.MultiSelect = false;
            binSource = conAccess.binSourceClass;

            //Переименовываем названия столбцов с системного на русский язык
            try
            {
                if (dataGridClass.RowCount != 0)
                {
                    dataGridClass.Columns[0].Visible = false;
                    dataGridClass.Columns[1].HeaderText = "Номер класса";
                    dataGridClass.Columns[2].HeaderText = "Буква";
                    dataGridClass.Columns[3].HeaderText = "Форма обучения";
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
            textAddNomerClass.Clear();
            textAddBukvaClass.Clear();
            textAddFormClass.Clear();
        }

        //Кнопка "Удалить"
        private void buttonDeleteRecord_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Classes") == false)
            {
                MessageBox.Show("Все строки были удалены из базы", "Ошибка удаления!");
            }
            else
            {
                //Определяем индекс выбранной строки
                int i = dataGridClass.CurrentRow.Index;
                string id_Class = Convert.ToString(dataGridClass[0, i].Value);
                //Удаление строки
                conAccess.QueryToBool("DELETE FROM Classes WHERE ID_classa = " + id_Class);
                //Зачем здесь эта строка? Во славу Сатане конечно :3
                binSource.RemoveAt(i);
                conAccess.QueryToDataGrid("SELECT * FROM Classes", dataGridClass, NavigatorClass, "Класс");
            }
        }

        //Кнопка "Очистка"
        private void buttonClearPrepod_Click(object sender, EventArgs e)
        {
            ClearText();
        }

        //Кнопка "Добавить/Изменить запись"
        private void buttonAddRecordClass_Click(object sender, EventArgs e)
        {
            int ID_Class = 0;

            if (textAddNomerClass.Text.Equals("") ||
                textAddBukvaClass.Text.Equals("") || textAddFormClass.Text.Equals(""))
            {
                MessageBox.Show("Не все поля введены", "Ошибка!");
            }
            else
            {
                if (Check_Button == 0) //Была нажата кнопка "Добавить"
                {
                    if (conAccess.QueryToBool("SELECT * FROM Classes") == true)
                    {
                        string ID = conAccess.AgregateQueryToDataGrid("SELECT MAX(ID_classa) FROM Classes");
                        try
                        {
                            ID_Class = Convert.ToInt32(ID);
                            ID_Class++;
                        }
                        catch (Exception exc)
                        {
                            MessageBox.Show(exc.Message);
                        }
                    }
                    else
                    {
                        ID_Class = 1;
                    }

                    string queryString = "INSERT INTO Classes VALUES (" + ID_Class + ",'" +
                        textAddNomerClass.Text + "','" + textAddBukvaClass.Text + "','" + textAddFormClass.Text + "')";
                    conAccess.QueryToBool(queryString);
                    conAccess.QueryToDataGrid("SELECT * FROM Classes", dataGridClass, NavigatorClass, "Класс");
                    ClearText();
                }
                else
                {
                    //Определяем индекс выбранной строки
                    int i = dataGridClass.CurrentRow.Index;
                    //Забор значения из 0 столбца i-тый строки
                    string id_Class = Convert.ToString(dataGridClass[0, i].Value);
                    string queryString = "UPDATE Classes SET Nomer_Classa = '"
                        + textAddNomerClass.Text + "', Bukva = '"
                        + textAddBukvaClass.Text + "', Forma_Ob = '"
                        + textAddFormClass.Text + "' WHERE ID_classa = " + id_Class;
                    conAccess.QueryToBool(queryString);
                    conAccess.QueryToDataGrid("SELECT * FROM Classes", dataGridClass, NavigatorClass, "Класс");
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
                label7.Text = "Добавление нового класса";
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
                int x = dataGridClass.CurrentRow.Index;
                //Забираем значение ячейки
                //textAddIdClass.Text = Convert.ToString(dataGridClass[0, x].Value);
                textAddNomerClass.Text = Convert.ToString(dataGridClass[1, x].Value);
                textAddBukvaClass.Text = Convert.ToString(dataGridClass[2, x].Value);
                textAddFormClass.Text = Convert.ToString(dataGridClass[3, x].Value);

                Check_Button = 1;
                buttonAddPrepod.Enabled = false;
                panelPrepod.Visible = true;
                buttonEditPrepod.Text = "Скрыть";
                label7.Text = "Изменение данных класса";
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
                main.Sort_Minus(dataGridClass, x);
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
                main.Sort_Plus(dataGridClass, x);
            }
        }

        //Кнопка "Экспорт в Excel"
        private void buttonExportExcelPrepod_Click(object sender, EventArgs e)
        {
            ExportsTo.ExportToExcel(dataGridClass);
        }

        //Кнопка "Поиск по таблице"
        private void buttonSearchComp_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            main.search_datagrid(dataGridClass, textSearchPrepod);
        }

        //Кнопка "Очистка"
        private void buttonClearSearchComp_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            main.clear_datagrid(dataGridClass);
        }

        //Кнопка "Фильтрация данных"
        private void buttonFilter_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Classes") == false)
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
                else if (comboBoxSortPrepod.SelectedIndex == 1) { Name_Column = "Nomer_Classa"; }
                else if (comboBoxSortPrepod.SelectedIndex == 2) { Name_Column = "Bukva"; }
                else if (comboBoxSortPrepod.SelectedIndex == 3) { Name_Column = "Forma_Ob"; }

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

        //Клик по таблице
        private void dataGridClass_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                textIdClass.DataBindings.Add(new Binding("Text", binSource, "ID_classa"));
                textNomerClass.DataBindings.Add(new Binding("Text", binSource, "Nomer_Classa"));
                textBukvaClass.DataBindings.Add(new Binding("Text", binSource, "Bukva"));
                textFormClass.DataBindings.Add(new Binding("Text", binSource, "Forma_Ob"));
            }
            catch //Чтобы работало при навигации туда сюда, не обрабатываю исключение, гы :D
            { }
        }

        //Запрет на все символы кроме цифр и клавишы BackSpace
        private void textAddIdClass_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                e.Handled = true;
        }

        //Запрет на все символы кроме цифр и клавишы BackSpace
        private void textAddNomerClass_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                e.Handled = true;
        }
    }
}
