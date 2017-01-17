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

namespace School_Helper_Version_2._0
{
    public partial class Main : Form
    {
        public int Check_Button; //Если 0 - Кнопка "Добавить", если 1 - Кнопка "Изменить"

        public BindingSource binSourceAll;

        public bool bAdminStatus = true;

        //------------- Стандартные процедура -------------//
        public Exports ExportsTo = new Exports();
        public ConnectorAccess conAccess;
        public Main(ConnectorAccess ClassConSQL)
        {
            InitializeComponent();
            this.conAccess = ClassConSQL;
        }

        private const int CP_NOCLOSE_BUTTON = 0x200;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams myCp = base.CreateParams;
                myCp.ClassStyle = myCp.ClassStyle | CP_NOCLOSE_BUTTON;
                return myCp;
            }
        }

        //Кнопка "Выход", для закрытия программы
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            conAccess.CloseConnection();
            Application.Exit();
        }

        //Кнопка "Настройки соединения", для открытия формы настройки соединения с базой данных
        private void настройкиСоединенияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Settings Stngs = new Settings(conAccess);
            Stngs.button1.Enabled = false;
            Stngs.button3.Enabled = true;
            Stngs.Show();
            this.Close();
            /* Main main = new Main(conAccess);
             main.Close();*/
        }

        //Кнопка "О программе", для окрытия формы с информацией о программе
        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutProgram About = new AboutProgram();
            About.Show();
        }

        //Кнопка "Справка(F1)", вызывает справочное средство
        private void справкаF1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Help.ShowHelp(this, helpProvider1.HelpNamespace);
        }

        //Кнопка "Классные руководители", для открытия формы классных руководителей
        private void классныеРуководителиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Classnaya classnaya = new Classnaya(conAccess);
            classnaya.Show();
        }

        //Кнопка "Классы", для открытия формы классов
        private void классыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Class klass = new Class(conAccess);
            klass.Show();
        }

        //Кнопка "События", для открытия формы событий
        private void мерроприятияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Events events = new Events(conAccess);
            events.Show();
        }

        //Кнопка "Олимпиады", для открытия формы олимпиад
        private void олимпиадыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Olympics olymp = new Olympics(conAccess);
            olymp.Show();
        }

        //Аторизация
        private void авторизацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Splashscreen splash = new Splashscreen(conAccess);
            splash.Show();
            this.Close();
        }

        //Загрузка формы
        private void Main_Load(object sender, EventArgs e)
        {
            //Проверка на администрирование
            if (bAdminStatus == true)
            {
                классныеРуководителиToolStripMenuItem.Enabled = true;
                классыToolStripMenuItem.Enabled = true;
                мерроприятияToolStripMenuItem.Enabled = true;
                олимпиадыToolStripMenuItem.Enabled = true;
                настройкиToolStripMenuItem.Enabled = true;

                buttonAddDeti.Enabled = true;
                buttonEditDeti.Enabled = true;

                buttonAddRod.Enabled = true;
                buttonEditRod.Enabled = true;
            }
            else
            {
                классныеРуководителиToolStripMenuItem.Enabled = false;
                классыToolStripMenuItem.Enabled = false;
                мерроприятияToolStripMenuItem.Enabled = false;
                олимпиадыToolStripMenuItem.Enabled = false;
                настройкиToolStripMenuItem.Enabled = false;

                buttonAddDeti.Enabled = false;
                buttonEditDeti.Enabled = false;

                buttonAddRod.Enabled = false;
                buttonEditRod.Enabled = false;
            }

            //Загрузка списка учителей
            if (conAccess.QueryToBool("SELECT * FROM Teachers") == true)
            {
                conAccess.QueryToComboBox("SELECT DISTINCT (Familiya + ' ' + Imya + ' ' + Otchestvo) AS FIO FROM Teachers", textAddClassnayaDeti, "FIO");
            }
            //Загрузка списка классов(буквы)
            if (conAccess.QueryToBool("SELECT * FROM Classes") == true)
            {
                conAccess.QueryToComboBox("SELECT DISTINCT Bukva FROM Classes", textAddBukvaDeti, "Bukva");
            }
            //Загрузка списка классов (номера)
            if (conAccess.QueryToBool("SELECT * FROM Classes") == true)
            {
                conAccess.QueryToComboBox("SELECT DISTINCT Nomer_Classa FROM Classes", textAddClassDeti, "Nomer_Classa");
            }
            //Загрузка списка родителей(мама)
            if (conAccess.QueryToBool("SELECT * FROM Parents") == true)
            {
                conAccess.QueryToComboBox("SELECT DISTINCT (Familiya + ' ' + Imya + ' ' + Otchestvo) AS FIO FROM Parents WHERE Parents.pol = 'Женский'", 
                    textAddMama, "FIO");
            }
            //Загрузка списка родителей(папа)
            if (conAccess.QueryToBool("SELECT * FROM Parents") == true)
            {
                conAccess.QueryToComboBox("SELECT DISTINCT (Familiya + ' ' + Imya + ' ' + Otchestvo) AS FIO FROM Parents WHERE Parents.pol = 'Мужской'",
                    textAddPapa, "FIO");
            }
            //Загрузка списка классов(буквы)
            if (conAccess.QueryToBool("SELECT * FROM Classes") == true)
            {
                conAccess.QueryToComboBox("SELECT DISTINCT Bukva FROM Classes", comboBoxBukva, "Bukva");
            }
            //Загрузка списка классов (номера)
            if (conAccess.QueryToBool("SELECT * FROM Classes") == true)
            {
                conAccess.QueryToComboBox("SELECT DISTINCT Nomer_Classa FROM Classes", comboBoxClass, "Nomer_Classa");
            }

            //Смена текста на черный
            this.dataGridDeti.DefaultCellStyle.ForeColor = Color.Black;

            helpProvider1.HelpNamespace = @"Data\Help\Helpout.chm";
            helpProvider1.SetHelpNavigator(this, HelpNavigator.Topic);
            helpProvider1.SetShowHelp(this, true);
            //Выбор активной вкладки, вкладки "Компьютеры"
            this.tabControlUchet.SelectedTab = tabPageDeti;
            conAccess.QueryToDataGrid("SELECT * FROM Childrens", dataGridDeti, NavigatorDeti, "Дети");
            dataGridSettings(dataGridDeti);
            binSourceAll = conAccess.binSourceDeti;
            //Таблицы которые есть в базе: Children, Classes, Parents, Teachers, Events
            //Переименовываем названия столбцов с системного на русский язык
            try
            {
                if (dataGridDeti.RowCount != 0)
                {
                    dataGridDeti.Columns[0].Visible = false;
                    dataGridDeti.Columns[1].HeaderText = "Имя";
                    dataGridDeti.Columns[2].HeaderText = "Фамилия";
                    dataGridDeti.Columns[3].HeaderText = "Отчество";
                    dataGridDeti.Columns[4].HeaderText = "Дата рождения";
                    dataGridDeti.Columns[5].HeaderText = "Адрес";
                    dataGridDeti.Columns[6].HeaderText = "Класс";
                    dataGridDeti.Columns[7].HeaderText = "Буква";
                    dataGridDeti.Columns[8].HeaderText = "Форма обучения";
                    dataGridDeti.Columns[9].HeaderText = "Дата зачисления";
                    dataGridDeti.Columns[10].HeaderText = "Номер приказа";
                    dataGridDeti.Columns[11].HeaderText = "Дата окончания";
                    dataGridDeti.Columns[12].HeaderText = "Причина отчисления";
                    dataGridDeti.Columns[13].HeaderText = "Куда выбыл";
                    dataGridDeti.Columns[14].HeaderText = "Классная";
                    dataGridDeti.Columns[15].HeaderText = "Семья";
                    dataGridDeti.Columns[16].HeaderText = "Номер приказа об отчислении";
                    dataGridDeti.Columns[17].HeaderText = "Статус";
                    dataGridDeti.Columns[18].HeaderText = "Дом телефон";
                    dataGridDeti.Columns[19].HeaderText = "Моб телефон";
                    dataGridDeti.Columns[20].HeaderText = "Друг телефон";
                    dataGridDeti.Columns[21].HeaderText = "Медецинские показания";
                    dataGridDeti.Columns[22].HeaderText = "Мама";
                    dataGridDeti.Columns[23].HeaderText = "Папа";
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        //Происходит при смене вкладки
        private void tabControlUchet_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (tabControlUchet.SelectedTab == tabPageDeti) //Вкладка "Дети"
            {
                conAccess.QueryToDataGrid("SELECT * FROM Childrens", dataGridDeti, NavigatorDeti, "Дети");
                dataGridSettings(dataGridDeti);
                binSourceAll = conAccess.binSourceDeti;


                //Переименовываем названия столбцов с системного на русский язык
                try
                {
                    if (dataGridDeti.RowCount != 0)
                    {
                        dataGridDeti.Columns[0].Visible = false;
                        dataGridDeti.Columns[1].HeaderText = "Имя";
                        dataGridDeti.Columns[2].HeaderText = "Фамилия";
                        dataGridDeti.Columns[3].HeaderText = "Отчество";
                        dataGridDeti.Columns[4].HeaderText = "Дата рождения";
                        dataGridDeti.Columns[5].HeaderText = "Адрес";
                        dataGridDeti.Columns[6].HeaderText = "Класс";
                        dataGridDeti.Columns[7].HeaderText = "Буква";
                        dataGridDeti.Columns[8].HeaderText = "Форма обучения";
                        dataGridDeti.Columns[9].HeaderText = "Дата зачисления";
                        dataGridDeti.Columns[10].HeaderText = "Номер приказа";
                        dataGridDeti.Columns[11].HeaderText = "Дата окончания";
                        dataGridDeti.Columns[12].HeaderText = "Причина отчисления";
                        dataGridDeti.Columns[13].HeaderText = "Куда выбыл";
                        dataGridDeti.Columns[14].HeaderText = "Классная";
                        dataGridDeti.Columns[15].HeaderText = "Семья";
                        dataGridDeti.Columns[16].HeaderText = "Номер приказа об отчислении";
                        dataGridDeti.Columns[17].HeaderText = "Статус";
                        dataGridDeti.Columns[18].HeaderText = "Дом телефон";
                        dataGridDeti.Columns[19].HeaderText = "Моб телефон";
                        dataGridDeti.Columns[20].HeaderText = "Друг телефон";
                        dataGridDeti.Columns[21].HeaderText = "Медецинские показания";
                        dataGridDeti.Columns[22].HeaderText = "Мама";
                        dataGridDeti.Columns[23].HeaderText = "Папа";
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message);
                }
            }
            else if (tabControlUchet.SelectedTab == tabPageRoditeli) //Вкладка "Родители"
            {
                conAccess.QueryToDataGrid("SELECT * FROM Parents", dataGridRoditeli, NavigatorRoditeli, "Родители");
                dataGridSettings(dataGridRoditeli);
                binSourceAll = conAccess.binSourceParents;


                //Переименовываем названия столбцов с системного на русский язык
                try
                {
                    if (dataGridRoditeli.RowCount != 0)
                    {
                        dataGridRoditeli.Columns[0].Visible = false;
                        dataGridRoditeli.Columns[1].HeaderText = "Фамилия";
                        dataGridRoditeli.Columns[2].HeaderText = "Имя";
                        dataGridRoditeli.Columns[3].HeaderText = "Отчество";
                        dataGridRoditeli.Columns[4].HeaderText = "Пол";
                        dataGridRoditeli.Columns[5].HeaderText = "Возраст";
                        dataGridRoditeli.Columns[6].HeaderText = "Тел моб";
                        dataGridRoditeli.Columns[7].HeaderText = "Тел дом";
                        dataGridRoditeli.Columns[8].HeaderText = "Тел раб";
                        dataGridRoditeli.Columns[9].HeaderText = "Адрес";
                        dataGridRoditeli.Columns[10].HeaderText = "Место работы";
                        dataGridRoditeli.Columns[11].HeaderText = "Должность";
                        dataGridRoditeli.Columns[12].HeaderText = "Семья";
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message);
                }
            }
        }

        //------------- Конец Стандартные процедура -------------//

        //------------- Дополнительные процедуры -------------//

        //Процедура настройки дата грида
        public void dataGridSettings(DataGridView dataGV)
        {
            dataGV.ReadOnly = true;
            dataGV.AllowUserToAddRows = false;
            dataGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGV.MultiSelect = false;
        }

        //Процедура очистки результатов поиска
        public void clear_datagrid(DataGridView dGV)
        {
            int i;
            int j;
            for (i = 0; i <= dGV.ColumnCount - 1; i++)
            {
                for (j = 0; j <= dGV.RowCount - 1; j++)
                {
                    dGV[i, j].Style.BackColor = Color.White;
                    dGV[i, j].Style.ForeColor = Color.Black;
                }
            }
        }

        //Процедура "Поиска по datagrid"
        public void search_datagrid(DataGridView dGV, TextBox text)
        {
            int i, j = 0;
            for (i = 0; i < dGV.ColumnCount; i++)
            {
                for (j = 0; j < dGV.RowCount; j++)
                {
                    dGV[i, j].Style.BackColor = Color.Black;
                    dGV[i, j].Style.ForeColor = Color.Orange;
                }
            }
            for (i = 0; i < dGV.ColumnCount; i++)
            {
                for (j = 0; j < dGV.RowCount; j++)
                {
                    if ((dGV[i, j].FormattedValue.ToString().Contains(text.Text.Trim())))
                    {
                        dGV[i, j].Style.BackColor = Color.White;
                        dGV[i, j].Style.ForeColor = Color.Green;
                    }
                }
            }
        }

        //Процедура сортировки плюс
        public void Sort_Plus(DataGridView dGV, int x)
        {
            dGV.Sort(dGV.Columns[x], ListSortDirection.Ascending);
        }

        //Процедура сортировки минус
        public void Sort_Minus(DataGridView dGV, int x)
        {
            dGV.Sort(dGV.Columns[x], ListSortDirection.Descending);
        }

        //Процедура удаления фото
        public void Delete_Photo(DataGridView dg, string Folder)
        {
            //Определяем индекс выбранной строки
            int i = dg.CurrentRow.Index;
            //Забор значения из 0 столбца i-тый строки
            string name_file = Convert.ToString(dg[0, i].Value);
            //Удаление старых фотографий
            if (File.Exists(@"Data\Img\" + Folder + "\\" + name_file + ".bmp"))
            {
                File.Delete(@"Data\Img\" + Folder + "\\" + name_file + ".bmp");
            }
        }

        //Процедура изменения картинки
        public void Change_picture(DataGridView dGV, string Folder)
        {
            //Проверка на пустоту базы
            if (dGV.RowCount == 0)
            {
                MessageBox.Show("Отсутсвуют записи в таблице", "Ошибка добавления фото");
            }
            else
            {
                int i = dGV.SelectedCells[0].RowIndex;
                string s = Convert.ToString(dGV[0, i].Value);
                //Вызов процедуры добавления фото
                add_picture(s, Folder);
            }
        }

        //Процедура добавления фото
        public void add_picture(string Name_file, string Folder)
        {
            openFileDialog1.Filter = ("Image Files(*.BMP;*.JPG;*.PNG)|*.BMP;*.JPG;*.PNG|All files (*.*)|*.*");
            try
            {
                //Добавление фотографии
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(@"Data\Img\" + Folder + '\\' + Name_file + ".bmp"))
                    {
                        File.Delete(@"Data\Img\" + Folder + '\\' + Name_file + ".bmp");
                        File.Copy(openFileDialog1.FileName, @"Data\Img\" + Folder + '\\' + Name_file + ".bmp");
                        MessageBox.Show("Повторное добавление", "Уже существует");
                    }
                    else
                    {
                        File.Copy(openFileDialog1.FileName, @"Data\Img\" + Folder + '\\' + Name_file + ".bmp");
                    }
                }
                else
                {
                    MessageBox.Show("Отмена добавления", "Отмена действия!");
                }
            }
            catch { }
        }

        //Процедура "клик по таблице"
        public void Click_for_table(DataGridView dg, PictureBox pb, string Folder)
        {
            if (dg.RowCount == 0)
            {
                MessageBox.Show("Отсутсвуют строки в таблице!", "Ошибка!");
            }
            else
            {
                //Очистка поля для фото
                pb.Image = null;
                //Установка растяжения по всей площади
                pb.BackgroundImageLayout = ImageLayout.Stretch;
                int i = 0;
                //Определяем индекс строки
                i = dg.SelectedCells[0].RowIndex;
                try
                {
                    //Если файл существует по пути, то загружаем фото, если нету, то картинку с ошибкой
                    if (File.Exists(@"Data\Img\" + Folder + "\\" +
                        Convert.ToString(dg[0, i].Value) + ".bmp"))
                    {
                        System.IO.FileStream fs = new System.IO.FileStream(@"Data\Img\" +
                            Folder + "\\" + Convert.ToString(dg[0, i].Value) + ".bmp", System.IO.FileMode.Open);
                        System.Drawing.Image img = System.Drawing.Image.FromStream(fs);
                        fs.Close();
                        pb.Image = img;
                    }
                    else
                    {
                        pb.BackgroundImage = Image.FromFile(@"Data\Img\404.png");
                    }
                }
                catch { }
            }
        }

        //------------- Конец Дополнительные процедуры -------------//

        // ------------- Вкладка Дети ------------- //

        public string id_deti_old;

        //Процедура очистки полей ввода
        public void ClearTextDeti()
        {
            textAddImyaDeti.Clear();
            textAddFamDeti.Clear();
            textAddOtchDeti.Clear();
            textAddAdresDeti.Clear();
            textAddNomerPrikDeti.Clear();
            textAddPrichOtchDeti.Clear();
            textAddKudaDeti.Clear();
            textAddSemiaDeti.Clear();
            textAddNomerPrikOtDeti.Clear();
            textAddDomTelDeti.Clear();
            textAddMobTelDeti.Clear();
            textAddDrugTelDeti.Clear();
            textAddMedPokazania.Clear();
        }

        //Кнопка "Экспорт в Excel", для печати таблицы "Дети"
        private void buttonExportExcelDeti_Click(object sender, EventArgs e)
        {
            ExportsTo.ExportToExcel(dataGridDeti);
        }

        //Кнопка "Первая запись" на навигаторе, вкладка "Дети"
        private void bindingNavigatorMoveFirstItem_Click(object sender, EventArgs e)
        {
            Click_for_table(dataGridDeti, pictureBoxDeti, "Childrens");
        }

        //Кнопка "Предыдущая запись" на навигаторе, вкладка "Дети"
        private void bindingNavigatorMovePreviousItem_Click(object sender, EventArgs e)
        {
            Click_for_table(dataGridDeti, pictureBoxDeti, "Childrens");
        }

        //Кнопка "Следующая запись" на навигаторе, вкладка "Дети"
        private void bindingNavigatorMoveNextItem_Click(object sender, EventArgs e)
        {
            Click_for_table(dataGridDeti, pictureBoxDeti, "Childrens");
        }

        //Кнопка "Последняя запись" на навигаторе, вкладка "Дети"
        private void bindingNavigatorMoveLastItem_Click(object sender, EventArgs e)
        {
            Click_for_table(dataGridDeti, pictureBoxDeti, "Childrens");
        }

        //Кнопка "Удалить запись" на навигаторе, вкладка "дети"
        private void buttonDelRecordDeti_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Childrens") == false)
            {
                MessageBox.Show("Все строки были удалены из базы", "Ошибка удаления!");
            }
            else
            {
                //Определяем индекс выбранной строки
                int i = dataGridDeti.CurrentRow.Index;
                string id_Deti = Convert.ToString(dataGridDeti[0, i].Value);
                //Удаление строки
                conAccess.QueryToBool("DELETE FROM Childrens WHERE ID_rebenka = " + id_Deti);
                //Удаление картинки, привязанной к строке
                Delete_Photo(dataGridDeti, "Childrens");

                //Зачем здесь эта строка? Во славу Сатане конечно :3
                binSourceAll.RemoveAt(i);
                conAccess.QueryToDataGrid("SELECT * FROM Childrens", dataGridDeti, NavigatorDeti, "Дети");

                pictureBoxDeti.Image = null;
            }
        }

        //Кнопка "Изменить фото" на навигаторе, вкладка "Дети"
        private void buttonChangePhotoDeti_Click(object sender, EventArgs e)
        {
            Change_picture(dataGridDeti, "Childrens");
        }

        //Кнопка "Вперед", вкладка "Дети"
        private void buttonNextDeti_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Childrens") == true)
            {
                binSourceAll.MoveNext();
                Click_for_table(dataGridDeti, pictureBoxDeti, "Childrens");
            }
            else
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
        }

        //Кнопка "Назад", вкладка "дети"
        private void buttonPreviousDeti_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Childrens") == true)
            {
                binSourceAll.MovePrevious();
                Click_for_table(dataGridDeti, pictureBoxDeti, "Childrens");
            }
            else
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
        }

        //Кнопка "Первая", вкладка "Дети"
        private void buttonFirstDeti_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Childrens") == true)
            {
                binSourceAll.MoveFirst();
                Click_for_table(dataGridDeti, pictureBoxDeti, "Childrens");
            }
            else
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
        }

        //Кнопка "Последняя", вкладка "Дети"
        private void buttonLastDeti_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Childrens") == true)
            {
                binSourceAll.MoveLast();
                Click_for_table(dataGridDeti, pictureBoxDeti, "Childrens");
            }
            else
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
        }

        //Кнопка "Сортировка минус", вкладка "Дети"
        private void buttonSortMinus_Click(object sender, EventArgs e)
        {
            if (comboBoxDeti.Text.Equals(""))
            {
                MessageBox.Show("Выберите критерий сортировки", "Ошибка!");
            }
            else
            {
                int x = comboBoxDeti.SelectedIndex;
                Sort_Minus(dataGridDeti, x);
            }
        }

        //Кнопка "Сортировка плюс", вкладка "Дети"
        private void buttonSortPlus_Click(object sender, EventArgs e)
        {
            if (comboBoxDeti.Text.Equals(""))
            {
                MessageBox.Show("Выберите критерий сортировки", "Ошибка!");
            }
            else
            {
                int x = comboBoxDeti.SelectedIndex;
                Sort_Plus(dataGridDeti, x);
            }
        }

        //Вкладка "Поиск по таблице", вкладка "Дети"
        private void buttonSearchDeti_Click(object sender, EventArgs e)
        {
            search_datagrid(dataGridDeti, textSearchDeti);
        }

        //Вкладка "Очистка", вкладка "Дети"
        private void button1_Click(object sender, EventArgs e)
        {
            clear_datagrid(dataGridDeti);
        }

        //Кнопка "Добавить", вкладка "Дети"
        private void buttonAddDeti_Click(object sender, EventArgs e)
        {
            if (buttonAddDeti.Text == "Добавить")
            {
                Check_Button = 0;
                buttonEditDeti.Enabled = false;
                panelDeti.Visible = true;
                buttonAddDeti.Text = "Скрыть";
                ClearTextDeti();
                label15.Text = "Добавление нового ребенка";

                //Загрузка списка учителей
                if (conAccess.QueryToBool("SELECT * FROM Teachers") == true)
                {
                    conAccess.QueryToComboBox("SELECT DISTINCT (Familiya + ' ' + Imya + ' ' + Otchestvo) AS FIO FROM Teachers", textAddClassnayaDeti, "FIO");
                }
                //Загрузка списка классов(буквы)
                if (conAccess.QueryToBool("SELECT * FROM Classes") == true)
                {
                    conAccess.QueryToComboBox("SELECT DISTINCT Bukva FROM Classes", textAddBukvaDeti, "Bukva");
                }
                //Загрузка списка классов (номера)
                if (conAccess.QueryToBool("SELECT * FROM Classes") == true)
                {
                    conAccess.QueryToComboBox("SELECT DISTINCT Nomer_Classa FROM Classes", textAddClassDeti, "Nomer_Classa");
                }
                //Загрузка списка родителей(мама)
                if (conAccess.QueryToBool("SELECT * FROM Parents") == true)
                {
                    conAccess.QueryToComboBox("SELECT DISTINCT (Familiya + ' ' + Imya + ' ' + Otchestvo) AS FIO FROM Parents WHERE Parents.pol = 'Женский'",
                        textAddMama, "FIO");
                }
                //Загрузка списка родителей(папа)
                if (conAccess.QueryToBool("SELECT * FROM Parents") == true)
                {
                    conAccess.QueryToComboBox("SELECT DISTINCT (Familiya + ' ' + Imya + ' ' + Otchestvo) AS FIO FROM Parents WHERE Parents.pol = 'Мужской'",
                        textAddPapa, "FIO");
                }
            }
            else
            {
                Check_Button = 2;
                panelDeti.Visible = false;
                buttonEditDeti.Enabled = true;
                buttonAddDeti.Text = "Добавить";
                ClearTextDeti();
            }
        }

        //Кнопка "Редактировать запись", вкладка "Дети"
        private void buttonEditDeti_Click(object sender, EventArgs e)
        {
            if (buttonEditDeti.Text == "Изменить")
            {
                label15.Text = "Изменение данных ребенка";
                //Определяем индекс выбранной строки
                int i = dataGridDeti.CurrentRow.Index;
                id_deti_old = Convert.ToString(dataGridDeti[0, i].Value);
                ClearTextDeti();
                //Забираем значение ячейки   
                int x = dataGridDeti.CurrentRow.Index;
                //textAddIdDeti.Text = Convert.ToString(dataGridDeti[0, x].Value);
                textAddImyaDeti.Text = Convert.ToString(dataGridDeti[1, x].Value);
                textAddFamDeti.Text = Convert.ToString(dataGridDeti[2, x].Value);
                textAddOtchDeti.Text = Convert.ToString(dataGridDeti[3, x].Value);
                textAddDataRogdDeti.Text = Convert.ToString(dataGridDeti[4, x].Value);
                textAddAdresDeti.Text = Convert.ToString(dataGridDeti[5, x].Value);
                textAddClassDeti.Text = Convert.ToString(dataGridDeti[6, x].Value);
                textAddBukvaDeti.Text = Convert.ToString(dataGridDeti[7, x].Value);
                textAddFormObdeti.Text = Convert.ToString(dataGridDeti[8, x].Value);
                textAddDataZachDeti.Text = Convert.ToString(dataGridDeti[9, x].Value);
                textAddNomerPrikDeti.Text = Convert.ToString(dataGridDeti[10, x].Value);
                textAddDataOkDeti.Text = Convert.ToString(dataGridDeti[11, x].Value);
                textAddPrichOtchDeti.Text = Convert.ToString(dataGridDeti[12, x].Value);
                textAddKudaDeti.Text = Convert.ToString(dataGridDeti[13, x].Value);
                textAddClassnayaDeti.Text = Convert.ToString(dataGridDeti[14, x].Value);
                textAddSemiaDeti.Text = Convert.ToString(dataGridDeti[15, x].Value);
                textAddNomerPrikOtDeti.Text = Convert.ToString(dataGridDeti[16, x].Value);
                textAddStatusDeti.Text = Convert.ToString(dataGridDeti[17, x].Value);
                textAddDomTelDeti.Text = Convert.ToString(dataGridDeti[18, x].Value);
                textAddMobTelDeti.Text = Convert.ToString(dataGridDeti[19, x].Value);
                textAddDrugTelDeti.Text = Convert.ToString(dataGridDeti[20, x].Value);
                textAddMedPokazania.Text = Convert.ToString(dataGridDeti[21, x].Value);
                textAddMama.Text = Convert.ToString(dataGridDeti[22, x].Value);
                textAddPapa.Text = Convert.ToString(dataGridDeti[23, x].Value);

                Check_Button = 1;
                buttonAddDeti.Enabled = false;
                panelDeti.Visible = true;
                buttonEditDeti.Text = "Скрыть";
            }
            else
            {
                Check_Button = 2;
                panelDeti.Visible = false;
                buttonAddDeti.Enabled = true;
                buttonEditDeti.Text = "Изменить";
                ClearTextDeti();
            }
        }

        //Кнопка "Открыть картинку", вкладка "Дети"
        private void buttonAddPictDeti_Click(object sender, EventArgs e)
        {
            if ((textAddFamDeti.Text.Equals("")) || (textAddImyaDeti.Text.Equals("")))
            {
                if (conAccess.QueryToBool("SELECT * FROM Childrens") == true)
                {
                    string ID = conAccess.AgregateQueryToDataGrid("SELECT MAX(ID_rebenka) FROM Childrens");
                    try
                    {
                        ID_Deti = Convert.ToInt32(ID);
                        ID_Deti++;
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show(exc.Message);
                    }
                }
                else
                {
                    ID_Deti = 1;
                }

                add_picture(Convert.ToString(ID_Deti) + "_" + textAddFamDeti.Text + "_" + textAddImyaDeti.Text + "_" + textAddOtchDeti.Text, "Childrens");
            }
            else
            {
                MessageBox.Show("Не все поля введены, введите Имя и Фамилию", "Ошибка!");
            }
        }

        //Кнопка "Очистить", вкладка "Дети"
        private void buttonClearDeti_Click(object sender, EventArgs e)
        {
            ClearTextDeti();
        }

        //Двойной клик по таблице, вкладка "Дети"
        private void dataGridDeti_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //Определяем индекс выбранной строки
            int i = dataGridDeti.CurrentRow.Index;

            string Mama, Papa, idRebenok;

            idRebenok = Convert.ToString(dataGridDeti[0, i].Value);
            Mama = Convert.ToString(dataGridDeti[22, i].Value);
            Papa = Convert.ToString(dataGridDeti[23, i].Value);

            PolnayaAnketa anketa = new PolnayaAnketa(conAccess);
            anketa.Mama = Mama;
            anketa.Papa = Papa;
            anketa.idRebenok = idRebenok;

            anketa.Show();
        }

        //Клик по таблице "Дети"
        private void dataGridDeti_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Click_for_table(dataGridDeti, pictureBoxDeti, "Childrens");

            try
            {
                textIdDeti.DataBindings.Add(new Binding("Text", binSourceAll, "ID_rebenka"));
                textImyaDeti.DataBindings.Add(new Binding("Text", binSourceAll, "Imya"));
                textFamDeti.DataBindings.Add(new Binding("Text", binSourceAll, "Familiya"));
                textOtchDeti.DataBindings.Add(new Binding("Text", binSourceAll, "Otchestvo"));
                textDataRogdDeti.DataBindings.Add(new Binding("Text", binSourceAll, "DataRogdenia"));
                textAdresDeti.DataBindings.Add(new Binding("Text", binSourceAll, "Adres"));
                textClassDeti.DataBindings.Add(new Binding("Text", binSourceAll, "Klass"));
                textBukvaDeti.DataBindings.Add(new Binding("Text", binSourceAll, "Bukva"));
                textFormObDeti.DataBindings.Add(new Binding("Text", binSourceAll, "FormaObuch"));
                textDataZachDeti.DataBindings.Add(new Binding("Text", binSourceAll, "DataZachislenia"));
                textNomerPrikDeti.DataBindings.Add(new Binding("Text", binSourceAll, "NomerPrikaza"));
                textDataOkonDeti.DataBindings.Add(new Binding("Text", binSourceAll, "DataOkonch"));
                textPrichOtchDeti.DataBindings.Add(new Binding("Text", binSourceAll, "PrichinaOtchis"));
                textPosleDeti.DataBindings.Add(new Binding("Text", binSourceAll, "KudaVibil"));
                textClassnayaDeti.DataBindings.Add(new Binding("Text", binSourceAll, "Classnaya"));
                textSemiaDeti.DataBindings.Add(new Binding("Text", binSourceAll, "Semia"));
                textNomerPrikOtchDeti.DataBindings.Add(new Binding("Text", binSourceAll, "NomerPrikazaOtch"));
                textStatusDeti.DataBindings.Add(new Binding("Text", binSourceAll, "Status"));
                textDomTelDeti.DataBindings.Add(new Binding("Text", binSourceAll, "DomTel"));
                textMobTelDeti.DataBindings.Add(new Binding("Text", binSourceAll, "MobTel"));
                textDrugTelDeti.DataBindings.Add(new Binding("Text", binSourceAll, "DrugTel"));
                textMedPokazania.DataBindings.Add(new Binding("Text", binSourceAll, "MedPokazania"));
                textMamaDeti.DataBindings.Add(new Binding("Text", binSourceAll, "Mama"));
                textPapaDeti.DataBindings.Add(new Binding("Text", binSourceAll, "Papa"));
            }
            catch //Чтобы работало при навигации туды сюды, не обрабатываю исключение, гы :D
            { }
        }

        public int ID_Deti;

        //Кнопка "Подтверждение", вкладка "Дети"
        private void buttonAddRecordDeti_Click(object sender, EventArgs e)
        {
            if (Check_Button == 0) //Была нажата кнопка "Добавить"
            {
                if (conAccess.QueryToBool("SELECT * FROM Childrens") == true)
                {
                    string ID = conAccess.AgregateQueryToDataGrid("SELECT MAX(ID_rebenka) FROM Childrens");
                    try
                    {
                        ID_Deti = Convert.ToInt32(ID);
                        ID_Deti++;
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show(exc.Message);
                    }
                }
                else
                {
                    ID_Deti = 1;
                }

                string queryString = "INSERT INTO Childrens (ID_rebenka, Imya, Familiya, Otchestvo, DataRogdenia, Adres, "
                    + "Klass, Bukva, FormaObuch, DataZachislenia, NomerPrikaza, DataOkonch, "
                    + "PrichinaOtchis, KudaVibil, Classnaya, Semia, "
                    + "NomerPrikazaOtch, Status, DomTel, MobTel, DrugTel, MedPokazania, Mama, Papa) VALUES (" +
                                    ID_Deti + ",'" + textAddImyaDeti.Text +
                                    "','" + textAddFamDeti.Text + "','" + textAddOtchDeti.Text +
                                    "','" + textAddDataRogdDeti.Text + "','" + textAddAdresDeti.Text +
                                    "','" + textAddClassDeti.Text + "','" + textAddBukvaDeti.Text +
                                    "','" + textAddFormObdeti.Text + "','" + textAddDataZachDeti.Text +
                                    "','" + textAddNomerPrikDeti.Text + "','" + textAddDataOkDeti.Text +
                                    "','" + textAddPrichOtchDeti.Text + "','" + textAddKudaDeti.Text +
                                    "','" + textAddClassnayaDeti.Text + "','" + textAddSemiaDeti.Text +
                                    "','" + textAddNomerPrikOtDeti.Text +
                                    "','" + textAddStatusDeti.Text + "','" + textAddDomTelDeti.Text +
                                    "','" + textAddMobTelDeti.Text + "','" + textAddDrugTelDeti.Text +
                                    "','" + textAddMedPokazania.Text + "','" + textAddMama.Text + "','" + textAddPapa.Text + 
                                    "')";

                conAccess.QueryToBool(queryString);
                conAccess.QueryToDataGrid("SELECT * FROM Childrens", dataGridDeti, NavigatorDeti, "Дети");
                ClearTextDeti();
                //MessageBox.Show(queryString);
            }
            else
            {
                string queryString = "UPDATE Childrens SET Imya = '"
                    + textAddImyaDeti.Text + "', Familiya = '"
                    + textAddFamDeti.Text + "', Otchestvo = '"
                    + textAddOtchDeti.Text + "', DataRogdenia = '"
                    + textAddDataRogdDeti.Text + "', Adres = '"
                    + textAddAdresDeti.Text + "', Klass = '"
                    + textAddClassDeti.Text + "', Bukva = '"
                    + textAddBukvaDeti.Text + "', FormaObuch = '"
                    + textAddFormObdeti.Text + "', DataZachislenia = '"
                    + textAddDataZachDeti.Text + "', NomerPrikaza = '"
                    + textAddNomerPrikDeti.Text + "', DataOkonch = '"
                    + textAddDataOkDeti.Text + "', PrichinaOtchis = '"
                    + textAddPrichOtchDeti.Text + "', KudaVibil = '"
                    + textAddKudaDeti.Text + "', Classnaya = '"
                    + textAddClassnayaDeti.Text + "', Semia = '"
                    + textAddSemiaDeti.Text + "', NomerPrikazaOtch = '"
                    + textAddNomerPrikOtDeti.Text + "', Status = '"
                    + textAddStatusDeti.Text + "', DomTel = '"
                    + textAddDomTelDeti.Text + "', MobTel = '"
                    + textAddMobTelDeti.Text + "', DrugTel = '"
                    + textAddDrugTelDeti.Text + "', MedPokazania = '"
                    + textAddMedPokazania.Text + "', Mama = '"
                    + textAddMama.Text + "', Papa = '"
                    + textAddPapa.Text
                    + "' WHERE ID_rebenka = " + id_deti_old;
                conAccess.QueryToBool(queryString);
                conAccess.QueryToDataGrid("SELECT * FROM Childrens", dataGridDeti, NavigatorDeti, "Дети");
                ClearTextDeti();
                panelDeti.Visible = false;
                buttonAddDeti.Enabled = true;
                buttonEditDeti.Text = "Изменить";
            }
        }

        //Блок всех символов кроме 0-9 и BackSpace
        private void textAddIdDeti_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                e.Handled = true;
        }

        //Блок всех символов кроме 0-9 и BackSpace
        private void textAddSemiaNomerDeti_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                e.Handled = true;
        }

        //Фильтрация значений по фамилии
        private void textFilter_KeyUp(object sender, KeyEventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Childrens") == false)
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
            else
            {
                try
                {
                    binSourceAll.Filter = "[Familiya] LIKE '" + textFilterDeti.Text + "%'";
                }
                catch (Exception exp)
                {
                    binSourceAll.Filter = "";
                    MessageBox.Show(exp.Message);
                }
            }
        }

        //Фильтрация значений по соц статусу
        private void comboBoxFilterStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Childrens") == false)
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
            else
            {
                if (comboBoxFilterStatus.SelectedIndex == 0)
                {
                    binSourceAll.Filter = "";
                }
                else
                {
                    try
                    {
                        binSourceAll.Filter = "[Status] LIKE '" + comboBoxFilterStatus.Text + "%'";
                    }
                    catch (Exception exp)
                    {
                        binSourceAll.Filter = "";
                        MessageBox.Show(exp.Message);
                    }
                }
            }
        }

        // ------------- Конец Вкладка Дети ------------- //

        // ------------- Вкладка Родители ------------- //

        public string id_rod_old;

        public void ClearTextRod()
        {
            textAddFamRod.Clear();
            textAddImyaRod.Clear();
            textAddOtchRod.Clear();
            textAddMobTelRod.Clear();
            textAddDomTelRod.Clear();
            textAddRabTelRod.Clear();
            textAddAdresRod.Clear();
            textAddMestoRabRod.Clear();
            textAddDolgnRod.Clear();
            textAddSemiaDeti.Clear();
        }

        //Кнопка "Удалить запись", вкладка "Родители"
        private void buttonDelRecordRod_Click(object sender, EventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Parents") == false)
            {
                MessageBox.Show("Все строки были удалены из базы", "Ошибка удаления!");
            }
            else
            {
                //Определяем индекс выбранной строки
                int i = dataGridRoditeli.CurrentRow.Index;
                string id_Rod = Convert.ToString(dataGridRoditeli[0, i].Value);
                //Удаление строки
                conAccess.QueryToBool("DELETE FROM Parents WHERE ID_rod = " + id_Rod);
                //Удаление картинки, привязанной к строке
                Delete_Photo(dataGridRoditeli, "Parents");

                //Зачем здесь эта строка? Во славу Сатане конечно :3
                binSourceAll.RemoveAt(i);
                conAccess.QueryToDataGrid("SELECT * FROM Parents", dataGridRoditeli, NavigatorRoditeli, "Родители");

                pictureBoxRod.Image = null;
            }
        }

        //Кнопка "Изменить фото", вкладка "Родители"
        private void buttonChangePhotoRod_Click(object sender, EventArgs e)
        {
            Change_picture(dataGridRoditeli, "Parents");
        }

        //Кнопка "Добавить", вкладка "Родители"
        private void buttonAddRod_Click(object sender, EventArgs e)
        {
            if (buttonAddRod.Text == "Добавить")
            {
                Check_Button = 0;
                buttonEditRod.Enabled = false;
                panelRod.Visible = true;
                buttonAddRod.Text = "Скрыть";
                ClearTextRod();
                label75.Text = "Добавление нового родителя";
            }
            else
            {
                Check_Button = 2;
                panelRod.Visible = false;
                buttonEditRod.Enabled = true;
                buttonAddRod.Text = "Добавить";
                ClearTextRod();
            }
        }

        //Кнопка "Изменить", вкладка "Родители"
        private void buttonEditRod_Click(object sender, EventArgs e)
        {
            if (buttonEditRod.Text == "Изменить")
            {
                label75.Text = "Изменение данных родителя";
                //Определяем индекс выбранной строки
                int i = dataGridRoditeli.CurrentRow.Index;
                id_rod_old = Convert.ToString(dataGridRoditeli[0, i].Value);
                ClearTextRod();
                //Забираем значение ячейки   
                int x = dataGridRoditeli.CurrentRow.Index;

                //textAddIdRod.Text = Convert.ToString(dataGridRoditeli[0, x].Value);
                textAddFamRod.Text = Convert.ToString(dataGridRoditeli[1, x].Value);
                textAddImyaRod.Text = Convert.ToString(dataGridRoditeli[2, x].Value);
                textAddOtchRod.Text = Convert.ToString(dataGridRoditeli[3, x].Value);
                textAddPolRod.Text = Convert.ToString(dataGridRoditeli[4, x].Value);
                textAddVozrRod.Text = Convert.ToString(dataGridRoditeli[5, x].Value);
                textAddMobTelRod.Text = Convert.ToString(dataGridRoditeli[6, x].Value);
                textAddDomTelRod.Text = Convert.ToString(dataGridRoditeli[7, x].Value);
                textAddRabTelRod.Text = Convert.ToString(dataGridRoditeli[8, x].Value);
                textAddAdresRod.Text = Convert.ToString(dataGridRoditeli[9, x].Value);
                textAddMestoRabRod.Text = Convert.ToString(dataGridRoditeli[10, x].Value);
                textAddDolgnRod.Text = Convert.ToString(dataGridRoditeli[11, x].Value);
                textAddSemyaRod.Text = Convert.ToString(dataGridRoditeli[12, x].Value);

                Check_Button = 1;
                buttonAddRod.Enabled = false;
                panelRod.Visible = true;
                buttonEditRod.Text = "Скрыть";
            }
            else
            {
                Check_Button = 2;
                panelRod.Visible = false;
                buttonAddRod.Enabled = true;
                buttonEditRod.Text = "Изменить";
                ClearTextRod();
            }
        }

        //Кнопка "Искать", вкладка "Родители"
        private void buttonSearchRod_Click(object sender, EventArgs e)
        {
            search_datagrid(dataGridRoditeli, textSearchRod);
        }

        //Кнопка "Очистка", вкладка "Родители"
        private void buttonClearSearchRod_Click(object sender, EventArgs e)
        {
            clear_datagrid(dataGridRoditeli);
        }

        //Кнопка "Сортировка по убыванию", вкладка "Родители"
        private void buttonSortMinusRod_Click(object sender, EventArgs e)
        {
            if (comboBoxRod.Text.Equals(""))
            {
                MessageBox.Show("Выберите критерий сортировки", "Ошибка!");
            }
            else
            {
                int x = comboBoxRod.SelectedIndex;
                Sort_Minus(dataGridRoditeli, x);
            }
        }

        //Кнопка "Сортировка по возрастанию", вкладка "Родители"
        private void buttonSortPlusRod_Click(object sender, EventArgs e)
        {
            if (comboBoxRod.Text.Equals(""))
            {
                MessageBox.Show("Выберите критерий сортировки", "Ошибка!");
            }
            else
            {
                int x = comboBoxRod.SelectedIndex;
                Sort_Plus(dataGridRoditeli, x);
            }
        }

        //Кнопка "Экспорт в Excel", вкладка "Родители"
        private void buttonExportRod_Click(object sender, EventArgs e)
        {
            ExportsTo.ExportToExcel(dataGridRoditeli);
        }

        //Кнопка "Добавить картинку", вкладка "Родители"
        private void buttonAddPictRod_Click(object sender, EventArgs e)
        {
            int ID_Parents = 0;

            if ((textAddFamRod.Text.Equals("")) || (textAddImyaRod.Text.Equals("")) || (textAddOtchRod.Text.Equals("")))
            {
                if (conAccess.QueryToBool("SELECT * FROM Parents") == true)
                    {
                        string ID = conAccess.AgregateQueryToDataGrid("SELECT MAX(ID_rod) FROM Parents");
                        try
                        {
                            ID_Parents = Convert.ToInt32(ID);
                            ID_Parents++;
                        }
                        catch (Exception exc)
                        {
                            MessageBox.Show(exc.Message);
                        }
                    }
                    else
                    {
                        ID_Parents = 1;
                    }

                add_picture(Convert.ToString(ID_Parents) + "_" + textAddFamRod.Text + "_" + textAddImyaRod.Text + "_" + textAddOtchRod.Text, "Parents");
            }
            else
            {
                MessageBox.Show("Не все поля введены, введите Имя и Фамилию", "Ошибка!");
            }
        }

        //Кнопка "Очистка полей", вкладка "Родители"
        private void buttonClearRod_Click(object sender, EventArgs e)
        {
            ClearTextRod();
        }

        //Кнопка "Подтверждение", вкладка "Родители"
        private void buttonAddRecordRod_Click(object sender, EventArgs e)
        {
            if ((textAddFamRod.Text.Equals("")) || (textAddImyaRod.Text.Equals("")) ||
               (textAddOtchRod.Text.Equals("")) || (textAddPolRod.Text.Equals("")) ||
               (textAddVozrRod.Text.Equals("")) || (textAddMobTelRod.Text.Equals("")) ||
               (textAddDomTelRod.Text.Equals("")) || (textAddRabTelRod.Text.Equals("")) ||
               (textAddAdresRod.Text.Equals("")) || (textAddMestoRabRod.Text.Equals("")) ||
               (textAddDolgnRod.Text.Equals("")) || (textAddSemyaRod.Text.Equals("")))
            {
                MessageBox.Show("Не все поля введены", "Ошибка!");
            }
            else
            {
                int ID_Parents = 0;

                if (Check_Button == 0) //Была нажата кнопка "Добавить"
                {
                    if (conAccess.QueryToBool("SELECT * FROM Parents") == true)
                    {
                        string ID = conAccess.AgregateQueryToDataGrid("SELECT MAX(ID_rod) FROM Parents");
                        try
                        {
                            ID_Parents = Convert.ToInt32(ID);
                            ID_Parents++;
                        }
                        catch (Exception exc)
                        {
                            MessageBox.Show(exc.Message);
                        }
                    }
                    else
                    {
                        ID_Parents = 1;
                    }

                    string queryString = "INSERT INTO Parents (ID_rod, Familiya, Imya, Otchestvo, Pol, Vozrsast, "
                        + "TelMob, TelDom, TelRab, Address, Mesto_Raboti, Doljnost, Semia) VALUES (" +
                                        ID_Parents + ",'" + textAddFamRod.Text +
                                        "','" + textAddImyaRod.Text + "','" + textAddOtchRod.Text +
                                        "','" + textAddPolRod.Text + "','" + textAddVozrRod.Text +
                                        "','" + textAddMobTelRod.Text + "','" + textAddDomTelRod.Text +
                                        "','" + textAddRabTelRod.Text + "','" + textAddAdresRod.Text +
                                        "','" + textAddMestoRabRod.Text + "','" + textAddDolgnRod.Text +
                                        "','" + textAddSemyaRod.Text + "')";

                    conAccess.QueryToBool(queryString);
                    conAccess.QueryToDataGrid("SELECT * FROM Parents", dataGridRoditeli, NavigatorRoditeli, "Родители");
                    ClearTextRod();
                }
                else
                {
                    string queryString = "UPDATE Parents SET Familiya = '"
                        + textAddFamRod.Text + "', Imya = '"
                        + textAddImyaRod.Text + "', Otchestvo = '"
                        + textAddOtchRod.Text + "', Pol = '"
                        + textAddPolRod.Text + "', Vozrsast = '"
                        + textAddVozrRod.Text + "', TelMob = '"
                        + textAddMobTelRod.Text + "', TelDom = '"
                        + textAddDomTelRod.Text + "', TelRab = '"
                        + textAddRabTelRod.Text + "', Address = '"
                        + textAddAdresRod.Text + "', Mesto_Raboti = '"
                        + textAddMestoRabRod.Text + "', Doljnost = '"
                        + textAddDolgnRod.Text + "', Semia = '"
                        + textAddSemyaRod.Text + "' WHERE ID_rod = " + id_rod_old;
                    conAccess.QueryToBool(queryString);
                    conAccess.QueryToDataGrid("SELECT * FROM Parents", dataGridRoditeli, NavigatorRoditeli, "Родители");
                    ClearTextRod();
                    panelRod.Visible = false;
                    buttonAddRod.Enabled = true;
                    buttonEditRod.Text = "Изменить";
                }
            }
        }

        //Клик по таблице, вкладка "Родители"
        private void dataGridRoditeli_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Click_for_table(dataGridRoditeli, pictureBoxRod, "Parents");

            try
            {
                textIdRod.DataBindings.Add(new Binding("Text", binSourceAll, "ID_rod"));
                textFamRod.DataBindings.Add(new Binding("Text", binSourceAll, "Familiya"));
                textImyaRod.DataBindings.Add(new Binding("Text", binSourceAll, "Imya"));
                textOtchRod.DataBindings.Add(new Binding("Text", binSourceAll, "Otchestvo"));
                textPolRod.DataBindings.Add(new Binding("Text", binSourceAll, "Pol"));
                textVozrRod.DataBindings.Add(new Binding("Text", binSourceAll, "Vozrsast"));
                textMobTelRod.DataBindings.Add(new Binding("Text", binSourceAll, "TelMob"));
                textDomTelRod.DataBindings.Add(new Binding("Text", binSourceAll, "TelDom"));
                textRabTelRod.DataBindings.Add(new Binding("Text", binSourceAll, "TelRab"));
                textAdresRod.DataBindings.Add(new Binding("Text", binSourceAll, "Address"));
                textMestoRabRod.DataBindings.Add(new Binding("Text", binSourceAll, "Mesto_Raboti"));
                textDolgnRod.DataBindings.Add(new Binding("Text", binSourceAll, "Doljnost"));
                textSemyaRod.DataBindings.Add(new Binding("Text", binSourceAll, "Semia"));
            }
            catch //Чтобы работало при навигации туды сюды, не обрабатываю исключение, гы :D
            { }
        }

        //Фильтрация значений
        private void textFilterRod_KeyUp(object sender, KeyEventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Parents") == false)
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
            else
            {
                try
                {
                    binSourceAll.Filter = "[Familiya] LIKE '" + textFilterRod.Text + "%'";
                }
                catch (Exception exp)
                {
                    binSourceAll.Filter = "";
                    MessageBox.Show(exp.Message);
                }
            }
        }

        //Запрет на все символы кроме цифр и клавишы BackSpace
        private void textAddIdRod_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                e.Handled = true;
        }

        //Запрет на все символы кроме цифр и клавишы BackSpace
        private void textAddSemyaNomerRod_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
                e.Handled = true;
        }

        //Поиск по базе
        private void textBoxSearch_KeyPress(object sender, KeyPressEventArgs e)
        {
            string sColumn = "";

            //Обработка выбора какой столбец фильтруем
            if (comboBoxSearch.SelectedIndex == 0)
            {
                sColumn = "[ID_rebenka] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 1)
            {
                sColumn = "[Imya] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 2)
            {
                sColumn = "[Familiya] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 3)
            {
                sColumn = "[Otchestvo] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 4)
            {
                sColumn = "[DataRogdenia] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 5)
            {
                sColumn = "[Adres] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 6)
            {
                sColumn = "[Klass] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 7)
            {
                sColumn = "[Bukva] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 8)
            {
                sColumn = "[FormaObuch] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 9)
            {
                sColumn = "[DataZachislenia] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 10)
            {
                sColumn = "[NomerPrikaza] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 11)
            {
                sColumn = "[DataOkonch] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 12)
            {
                sColumn = "[PrichinaOtchis] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 13)
            {
                sColumn = "[KudaVibil] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 14)
            {
                sColumn = "[Classnaya] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 15)
            {
                sColumn = "[Semia] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 16)
            {
                sColumn = "[NomerPrikazaOtch] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 17)
            {
                sColumn = "[Status] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 18)
            {
                sColumn = "[DomTel] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 19)
            {
                sColumn = "[MobTel] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 20)
            {
                sColumn = "[DrugTel] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 21)
            {
                sColumn = "[MedPokazania] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 22)
            {
                sColumn = "[Mama] LIKE '";
            }
            else if (comboBoxSearch.SelectedIndex == 23)
            {
                sColumn = "[Papa] LIKE '";
            }

            //Сам поиск, фильтрация так скажем
            if (conAccess.QueryToBool("SELECT * FROM Childrens") == false)
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
            else
            {
                try
                {
                    binSourceAll.Filter = sColumn + "%" + textBoxSearch.Text + "%'";
                }
                catch (Exception exp)
                {
                    binSourceAll.Filter = "";
                    MessageBox.Show(exp.Message);
                }
            }

        }

        //Фильтрация значений по должности
        private void textFilterRodDolgn_KeyUp(object sender, KeyEventArgs e)
        {
            if (conAccess.QueryToBool("SELECT * FROM Parents") == false)
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
            else
            {
                try
                {
                    binSourceAll.Filter = "[Doljnost] LIKE '" + textFilterRodDolgn.Text + "%'";
                }
                catch (Exception exp)
                {
                    binSourceAll.Filter = "";
                    MessageBox.Show(exp.Message);
                }
            }
        }

        //Поиск так скажем по всем полям
        private void textBoxFilterRod_KeyUp(object sender, KeyEventArgs e)
        {
            /*
                Номер(Системный)
                Имя
                Фамилия
                Отчество
                Пол
                Возраст
                Адрес
                Семья
                Телефон домашний
                Телефон мобильный
                Телефон рабочий
                Место работы
                Должность             
             */


            string sColumn = "";

            //Обработка выбора какой столбец фильтруем
            if (comboBoxFilterRod.SelectedIndex == 0)
            {
                sColumn = "[ID_rod] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 1)
            {
                sColumn = "[Imya] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 2)
            {
                sColumn = "[Familiya] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 3)
            {
                sColumn = "[Otchestvo] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 4)
            {
                sColumn = "[Pol] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 5)
            {
                sColumn = "[Vozrsast] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 6)
            {
                sColumn = "[Address] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 7)
            {
                sColumn = "[Semia] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 8)
            {
                sColumn = "[TelDom] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 9)
            {
                sColumn = "[TelMob] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 10)
            {
                sColumn = "[TelRab] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 11)
            {
                sColumn = "[Mesto_Raboti] LIKE '";
            }
            else if (comboBoxFilterRod.SelectedIndex == 12)
            {
                sColumn = "[Doljnost] LIKE '";
            }

            //Сам поиск, фильтрация так скажем
            if (conAccess.QueryToBool("SELECT * FROM Parents") == false)
            {
                MessageBox.Show("Отсутсвуют записи в таблице!", "Ошибка!");
            }
            else
            {
                try
                {
                    binSourceAll.Filter = sColumn + "%" + textBoxFilterRod.Text + "%'";
                }
                catch (Exception exp)
                {
                    binSourceAll.Filter = "";
                    MessageBox.Show(exp.Message);
                }
            }

        }

        //Поиск класса
        private void buttonClassFilter_Click(object sender, EventArgs e)
        {
            conAccess.QueryToDataGrid("SELECT * FROM Childrens WHERE Bukva = '" + comboBoxBukva.Text + "' AND Klass = '" + comboBoxClass.Text + "'",
                dataGridDeti, NavigatorDeti, "Дети");
        }

        //Возврат значений
        private void buttonClassClear_Click(object sender, EventArgs e)
        {
            conAccess.QueryToDataGrid("SELECT * FROM Childrens", dataGridDeti, NavigatorDeti, "Дети");
        }

        // ------------- Конец Вкладка Родители ------------- //

    }
}