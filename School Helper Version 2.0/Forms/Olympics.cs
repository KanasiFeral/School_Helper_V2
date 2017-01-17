using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace School_Helper_Version_2._0.Forms
{
    public partial class Olympics : Form
    {
        public BindingSource binSource;

        public Exports ExportsTo = new Exports();
        public ConnectorAccess conAccess;
        public Olympics(ConnectorAccess ClassConSQL)
        {
            InitializeComponent();
            this.conAccess = ClassConSQL;
        }

        public int Check_Button;

        //Загрузка формы
        private void Olympics_Load(object sender, EventArgs e)
        {
            conAccess.QueryToDataGrid("SELECT * FROM Olympics", dataGridOlympics, NavigatorOlympics, "Олимпиады");
            dataGridOlympics.ReadOnly = true;
            dataGridOlympics.AllowUserToAddRows = false;
            dataGridOlympics.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridOlympics.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridOlympics.MultiSelect = false;
            binSource = conAccess.binSourceOlympics;

            //Переименовываем названия столбцов с системного на русский язык
            try
            {
                if (dataGridOlympics.RowCount != 0)
                {
                    dataGridOlympics.Columns[0].Visible = false;
                    dataGridOlympics.Columns[1].HeaderText = "Название";
                    dataGridOlympics.Columns[2].HeaderText = "Кто победил";
                    dataGridOlympics.Columns[3].HeaderText = "Кто учавствовал";
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
            textAddNazvanieOlympics.Clear();
            textAddWhoIsWinOlympics.Clear();
            textAddPeopleOlympics.Clear();
        }

        //Кнопка "Удалить"
        private void buttonDeleteRecord_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Olympics") == false)
            {
                MessageBox.Show("Все строки были удалены из базы", "Ошибка удаления!");
            }
            else
            {
                //Определяем индекс выбранной строки
                int i = dataGridOlympics.CurrentRow.Index;
                string id_Olympics = Convert.ToString(dataGridOlympics[0, i].Value);
                //Удаление строки
                conAccess.QueryToBool("DELETE FROM Olympics WHERE ID_olympics = " + id_Olympics);
                //Зачем здесь эта строка? Во славу Сатане конечно :3
                binSource.RemoveAt(i);
                conAccess.QueryToDataGrid("SELECT * FROM Olympics", dataGridOlympics, NavigatorOlympics, "Олимпиады");
            }
        }

        //Кнопка "Очистка"
        private void buttonClearOlympics_Click(object sender, EventArgs e)
        {
            ClearText();
        }

        //Кнопка "Добавить/Изменить запись"
        private void buttonAddRecordOlympics_Click(object sender, EventArgs e)
        {
            int ID_Olympics = 0;

            if (textAddNazvanieOlympics.Text.Equals("") ||
                textAddWhoIsWinOlympics.Text.Equals("") || textAddPeopleOlympics.Text.Equals(""))
            {
                MessageBox.Show("Не все поля введены", "Ошибка!");
            }
            else
            {
                if (Check_Button == 0) //Была нажата кнопка "Добавить"
                {
                    if (conAccess.QueryToBool("SELECT * FROM Olympics") == true)
                    {
                        string ID = conAccess.AgregateQueryToDataGrid("SELECT MAX(ID_olympics) FROM Olympics");
                        try
                        {
                            ID_Olympics = Convert.ToInt32(ID);
                            ID_Olympics++;
                        }
                        catch (Exception exc)
                        {
                            MessageBox.Show(exc.Message);
                        }
                    }
                    else
                    {
                        ID_Olympics = 1;
                    }

                    string queryString = "INSERT INTO Olympics VALUES (" + ID_Olympics + ",'" +
                        textAddNazvanieOlympics.Text + "','" + textAddWhoIsWinOlympics.Text + "','" + textAddPeopleOlympics.Text + "')";
                    conAccess.QueryToBool(queryString);
                    conAccess.QueryToDataGrid("SELECT * FROM Olympics", dataGridOlympics, NavigatorOlympics, "Олимпиады");
                    ClearText();
                }
                else
                {
                    //Определяем индекс выбранной строки
                    int i = dataGridOlympics.CurrentRow.Index;
                    //Забор значения из 0 столбца i-тый строки
                    string id_Olympics = Convert.ToString(dataGridOlympics[0, i].Value);
                    string queryString = "UPDATE Olympics SET NameOlympic = '"
                        + textAddNazvanieOlympics.Text + "', WhoIsWin = '"
                        + textAddWhoIsWinOlympics.Text + "', People = '"
                        + textAddPeopleOlympics.Text + "' WHERE ID_olympics = " + id_Olympics;
                    conAccess.QueryToBool(queryString);
                    conAccess.QueryToDataGrid("SELECT * FROM Olympics", dataGridOlympics, NavigatorOlympics, "Олимпиады");
                    ClearText();
                    panelPrepod.Visible = false;
                    buttonAddOlympics.Enabled = true;
                    buttonEditOlympics.Text = "Изменить";
                }
            }
        }

        //Кнопка "Добавить"
        private void buttonAddOlympics_Click(object sender, EventArgs e)
        {
            if (buttonAddOlympics.Text == "Добавить")
            {
                Check_Button = 0;
                buttonEditOlympics.Enabled = false;
                panelPrepod.Visible = true;
                buttonAddOlympics.Text = "Скрыть";
                ClearText();
                label7.Text = "Добавление новой олимпиады";
            }
            else
            {
                Check_Button = 2;
                panelPrepod.Visible = false;
                buttonEditOlympics.Enabled = true;
                buttonAddOlympics.Text = "Добавить";
                ClearText();
            }
        }

        //Кнопка "Изменить"
        private void buttonEditOlympics_Click(object sender, EventArgs e)
        {
            if (buttonEditOlympics.Text == "Изменить")
            {
                ClearText();
                int x = dataGridOlympics.CurrentRow.Index;
                //Забираем значение ячейки
                //textAddIdOlympics.Text = Convert.ToString(dataGridOlympics[0, x].Value);
                textAddNazvanieOlympics.Text = Convert.ToString(dataGridOlympics[1, x].Value);
                textAddWhoIsWinOlympics.Text = Convert.ToString(dataGridOlympics[2, x].Value);
                textAddPeopleOlympics.Text = Convert.ToString(dataGridOlympics[3, x].Value);

                Check_Button = 1;
                buttonAddOlympics.Enabled = false;
                panelPrepod.Visible = true;
                buttonEditOlympics.Text = "Скрыть";
                label7.Text = "Изменение данных олимпиады";
            }
            else
            {
                Check_Button = 2;
                panelPrepod.Visible = false;
                buttonAddOlympics.Enabled = true;
                buttonEditOlympics.Text = "Изменить";
                ClearText();
            }
        }

        //Кнопка "Сортировка по убыванию"
        private void buttonSortMinusOlympics_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            if (comboBoxSortOlympics.Text.Equals(""))
            {
                MessageBox.Show("Выберите критерий сортировки", "Ошибка!");
            }
            else
            {
                int x = comboBoxSortOlympics.SelectedIndex;
                main.Sort_Minus(dataGridOlympics, x);
            }
        }

        //Кнопка "Сортировка по убыванию"
        private void buttonSortPlusOlympics_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            if (comboBoxSortOlympics.Text.Equals(""))
            {
                MessageBox.Show("Выберите критерий сортировки", "Ошибка!");
            }
            else
            {
                int x = comboBoxSortOlympics.SelectedIndex;
                main.Sort_Plus(dataGridOlympics, x);
            }
        }

        //Кнопка "Экспорт в Excel"
        private void buttonExportExcelOlympics_Click(object sender, EventArgs e)
        {
            ExportsTo.ExportToExcel(dataGridOlympics);
        }

        //Кнопка "Поиск по таблице"
        private void buttonSearchOlympics_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            main.search_datagrid(dataGridOlympics, textSearchOlympics);
        }

        //Кнопка "Очистка"
        private void buttonClearSearchOlympics_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            main.clear_datagrid(dataGridOlympics);
        }

        //Кнопка "Фильтрация данных"
        private void buttonFilter_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Olympics") == false)
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
            else
            {
                string Name_Column = "";
                if (comboBoxSortOlympics.SelectedIndex == 0)
                {
                    MessageBox.Show("Фильтрация не работает с полями целочисленных типов(Int)!", "Предупреждение!");
                }
                else if (comboBoxSortOlympics.SelectedIndex == 1) { Name_Column = "NameOlympic"; }
                else if (comboBoxSortOlympics.SelectedIndex == 2) { Name_Column = "WhoIsWin"; }
                else if (comboBoxSortOlympics.SelectedIndex == 3) { Name_Column = "People"; }

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
        private void dataGridOlympics_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                textIdOlympics.DataBindings.Add(new Binding("Text", binSource, "ID_olympics"));
                textNazvanieOlympics.DataBindings.Add(new Binding("Text", binSource, "NameOlympic"));
                textWhoIsWinOlympics.DataBindings.Add(new Binding("Text", binSource, "WhoIsWin"));
                textPeopleOlympics.DataBindings.Add(new Binding("Text", binSource, "People"));
            }
            catch //Чтобы работало при навигации туда сюда, не обрабатываю исключение, гы :D
            { }
        }

        //Запрет на все символы кроме цифр и клавишы BackSpace
        private void textAddIdOlympics_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                e.Handled = true;
        }        
    }
}