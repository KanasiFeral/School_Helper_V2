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
    public partial class Events : Form
    {
        public BindingSource binSource;

        public Exports ExportsTo = new Exports();
        public ConnectorAccess conAccess;
        public Events(ConnectorAccess ClassConSQL)
        {
            InitializeComponent();
            this.conAccess = ClassConSQL;
        }

        public int Check_Button;

        //Загрузка формы
        private void Events_Load(object sender, EventArgs e)
        {
            conAccess.QueryToDataGrid("SELECT * FROM Events", dataGridEvents, NavigatorEvents, "События");
            dataGridEvents.ReadOnly = true;
            dataGridEvents.AllowUserToAddRows = false;
            dataGridEvents.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridEvents.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridEvents.MultiSelect = false;
            binSource = conAccess.binSourceEvents;

            //Переименовываем названия столбцов с системного на русский язык
            try
            {
                if (dataGridEvents.RowCount != 0)
                {
                    dataGridEvents.Columns[0].Visible = false;
                    dataGridEvents.Columns[1].HeaderText = "Название";
                    dataGridEvents.Columns[2].HeaderText = "Дата проведения";
                    dataGridEvents.Columns[3].HeaderText = "Место проведения";
                    dataGridEvents.Columns[4].HeaderText = "Количество человек";
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
            //textAddDataProvEvents.Clear();
            textAddKolChelEvents.Clear();
            textAddMestoProvEvents.Clear();
            textAddNazvanieEvents.Clear();
        }

        //Клик по таблице
        private void dataGridEvents_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                textIdEvents.DataBindings.Add(new Binding("Text", binSource, "Id_Events"));
                textNazvanieEvents.DataBindings.Add(new Binding("Text", binSource, "Nazvanie"));
                textDataProvEvents.DataBindings.Add(new Binding("Text", binSource, "DataProvedeniya"));
                textMestoProvEvents.DataBindings.Add(new Binding("Text", binSource, "MestoProvedeniya"));
                textKolChelEvents.DataBindings.Add(new Binding("Text", binSource, "KolichestvoChelovek"));
            }
            catch //Чтобы работало при навигации туда сюда, не обрабатываю исключение, гы :D
            { }
        }

        //Кнопка "Удалить запись"
        private void buttonDeleteRecordEvents_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Events") == false)
            {
                MessageBox.Show("Все строки были удалены из базы", "Ошибка удаления!");
            }
            else
            {
                //Определяем индекс выбранной строки
                int i = dataGridEvents.CurrentRow.Index;
                string id_Events = Convert.ToString(dataGridEvents[0, i].Value);
                //Удаление строки
                conAccess.QueryToBool("DELETE FROM Events WHERE Id_Events = " + id_Events);
                //Зачем здесь эта строка? Во славу Сатане конечно :3
                binSource.RemoveAt(i);
                conAccess.QueryToDataGrid("SELECT * FROM Events", dataGridEvents, NavigatorEvents, "События");
            }
        }

        //Кнопка "Очистка"
        private void buttonClearEvents_Click(object sender, EventArgs e)
        {
            ClearText();
        }

        //Кнопка "Добавить/Изменить запись"
        private void buttonAddRecordEvents_Click(object sender, EventArgs e)
        {
            int ID_Events = 0;

            if (textAddDataProvEvents.Text.Equals("") ||
                textAddKolChelEvents.Text.Equals("") || textAddMestoProvEvents.Text.Equals("") ||
                textAddNazvanieEvents.Text.Equals(""))
            {
                MessageBox.Show("Не все поля введены", "Ошибка!");
            }
            else
            {
                if (Check_Button == 0) //Была нажата кнопка "Добавить"
                {
                    if (conAccess.QueryToBool("SELECT * FROM Events") == true)
                    {
                        string ID = conAccess.AgregateQueryToDataGrid("SELECT MAX(Id_Events) FROM Events");
                        try
                        {
                            ID_Events = Convert.ToInt32(ID);
                            ID_Events++;
                        }
                        catch (Exception exc)
                        {
                            MessageBox.Show(exc.Message);
                        }
                    }
                    else
                    {
                        ID_Events = 1;
                    }

                    string queryString = "INSERT INTO Events VALUES (" + ID_Events + ",'" +
                        textAddNazvanieEvents.Text + "','" + textAddDataProvEvents.Text + "','" + textAddMestoProvEvents.Text + "','" +
                        textAddKolChelEvents.Text + "')";
                    conAccess.QueryToBool(queryString);
                    conAccess.QueryToDataGrid("SELECT * FROM Events", dataGridEvents, NavigatorEvents, "События");
                    ClearText();
                }
                else
                {
                    //Определяем индекс выбранной строки
                    int i = dataGridEvents.CurrentRow.Index;
                    //Забор значения из 0 столбца i-тый строки
                    string id_Events = Convert.ToString(dataGridEvents[0, i].Value);
                    string queryString = "UPDATE Events SET Nazvanie = '"
                        + textNazvanieEvents.Text + "', DataProvedeniya = '"
                        + textDataProvEvents.Text + "', MestoProvedeniya = '"
                        + textMestoProvEvents.Text + "', KolichestvoChelovek = '"
                        + textKolChelEvents.Text + "' WHERE Id_Events = " + id_Events;
                    conAccess.QueryToBool(queryString);
                    conAccess.QueryToDataGrid("SELECT * FROM Events", dataGridEvents, NavigatorEvents, "События");
                    ClearText();
                    panelEvents.Visible = false;
                    buttonAddEvents.Enabled = true;
                    buttonEditEvents.Text = "Изменить";
                }
            }
        }

        //Кнопка "Добавить"
        private void buttonAddEvents_Click(object sender, EventArgs e)
        {
            if (buttonAddEvents.Text == "Добавить")
            {
                Check_Button = 0;
                buttonEditEvents.Enabled = false;
                panelEvents.Visible = true;
                buttonAddEvents.Text = "Скрыть";
                ClearText();
                label7.Text = "Добавление нового преподователя";
            }
            else
            {
                Check_Button = 2;
                panelEvents.Visible = false;
                buttonEditEvents.Enabled = true;
                buttonAddEvents.Text = "Добавить";
                ClearText();
            }
        }

        //Кнопка "Изменить"
        private void buttonEditEvents_Click(object sender, EventArgs e)
        {
            if (buttonEditEvents.Text == "Изменить")
            {
                ClearText();
                int x = dataGridEvents.CurrentRow.Index;
                //Забираем значение ячейки
                //textAddIdEvents.Text = Convert.ToString(dataGridEvents[0, x].Value);
                textAddNazvanieEvents.Text = Convert.ToString(dataGridEvents[1, x].Value);
                textAddDataProvEvents.Text = Convert.ToString(dataGridEvents[2, x].Value);
                textAddMestoProvEvents.Text = Convert.ToString(dataGridEvents[3, x].Value);
                textAddKolChelEvents.Text = Convert.ToString(dataGridEvents[4, x].Value);

                Check_Button = 1;
                buttonAddEvents.Enabled = false;
                panelEvents.Visible = true;
                buttonEditEvents.Text = "Скрыть";
                label7.Text = "Изменение данных события";
            }
            else
            {
                Check_Button = 2;
                panelEvents.Visible = false;
                buttonAddEvents.Enabled = true;
                buttonEditEvents.Text = "Изменить";
                ClearText();
            }
        }

        //Кнопка "Сортировка по убыванию"
        private void buttonSortMinusEvents_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            if (comboBoxSortEvents.Text.Equals(""))
            {
                MessageBox.Show("Выберите критерий сортировки", "Ошибка!");
            }
            else
            {
                int x = comboBoxSortEvents.SelectedIndex;
                main.Sort_Minus(dataGridEvents, x);
            }
        }

        //Кнопка "Сортировка по возрастанию"
        private void buttonSortPlusEvents_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            if (comboBoxSortEvents.Text.Equals(""))
            {
                MessageBox.Show("Выберите критерий сортировки", "Ошибка!");
            }
            else
            {
                int x = comboBoxSortEvents.SelectedIndex;
                main.Sort_Plus(dataGridEvents, x);
            }
        }

        //Кнопка "Поиск по таблице"
        private void buttonSearchEvents_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            main.search_datagrid(dataGridEvents, textSearchEvents);
        }

        //Кнопка "Очистка"
        private void buttonClearSearchEvents_Click(object sender, EventArgs e)
        {
            Main main = new Main(conAccess);
            main.clear_datagrid(dataGridEvents);
        }

        //Кнопка "Экспорт в Excel"
        private void buttonExportExcelEvents_Click(object sender, EventArgs e)
        {
            ExportsTo.ExportToExcel(dataGridEvents);
        }

        //Кнопка "Фильтрация данных"
        private void buttonFilter_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Events") == false)
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
            else
            {
                string Name_Column = "";
                if (comboBoxSortEvents.SelectedIndex == 0)
                {
                    MessageBox.Show("Фильтрация не работает с полями целочисленных типов(Int)!", "Предупреждение!");
                }
                else if (comboBoxSortEvents.SelectedIndex == 1) { Name_Column = "Nazvanie"; }
                else if (comboBoxSortEvents.SelectedIndex == 2) { Name_Column = "DataProvedeniya"; }
                else if (comboBoxSortEvents.SelectedIndex == 3) { Name_Column = "MestoProvedeniya"; }
                else if (comboBoxSortEvents.SelectedIndex == 4) { Name_Column = "KolichestvoChelovek"; }

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
        private void textAddIdEvents_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                e.Handled = true;
        }
    }
}