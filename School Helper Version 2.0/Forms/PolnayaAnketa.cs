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
    public partial class PolnayaAnketa : Form
    {
        public BindingSource binSourceDeti;
        public BindingSource binSourceMama;
        public BindingSource binSourcePapa;

        public string Mama, Papa, idRebenok;

        public ConnectorAccess conAccess;
        public PolnayaAnketa(ConnectorAccess ClassConSQL)
        {
            InitializeComponent(); 
            this.conAccess = ClassConSQL;
        }

        //Процедура настройки дата грида
        public void dataGridSettings(DataGridView dataGV)
        {
            dataGV.ReadOnly = true;
            dataGV.AllowUserToAddRows = false;
            dataGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGV.MultiSelect = false;
        }

        //Загрузка формы
        private void PolnayaAnketa_Load(object sender, EventArgs e)
        {
            //Переменные для мамы
            string Fam_Mama;
            string Name_Otch_Mama;
            string Name_Mama;
            string Otch_Mama;
            //Переменные для папы
            string Fam_Papa;
            string Name_Otch_Papa;
            string Name_Papa;
            string Otch_Papa;

            //Обрезка
            //Селиверствова Дарья Николаевна
            int Probel_Posle_Fam_Mama = Mama.IndexOf(" ", 0); //Определяю вхождение первого пробела (После фамилии)
            int Probel_Posle_Fam_Papa = Papa.IndexOf(" ", 0); //Определяю вхождение первого пробела (После фамилии)

            Probel_Posle_Fam_Mama++;
            Probel_Posle_Fam_Papa++;

            Name_Otch_Mama = Mama.Substring(Probel_Posle_Fam_Mama, Mama.Length - Probel_Posle_Fam_Mama); //Копирование (Имя, Отчество)
            Name_Otch_Papa = Papa.Substring(Probel_Posle_Fam_Papa, Papa.Length - Probel_Posle_Fam_Papa); //Копирование (Имя, Отчество)

            Fam_Mama = Mama.Remove(Probel_Posle_Fam_Mama); //Копируем фамилию преподователя в переменную
            Fam_Papa = Papa.Remove(Probel_Posle_Fam_Papa); //Копируем фамилию преподователя в переменную

            //Дарья Николаевна
            int Probel_Posle_Imeni_Mama = Name_Otch_Mama.IndexOf(" ", 1); //Определяю вхождение первого пробела (После имени)
            int Probel_Posle_Imeni_Papa = Name_Otch_Papa.IndexOf(" ", 1); //Определяю вхождение первого пробела (После имени)

            Probel_Posle_Imeni_Mama++;
            Probel_Posle_Imeni_Papa++;

            Otch_Mama = Name_Otch_Mama.Substring(Probel_Posle_Imeni_Mama, Name_Otch_Mama.Length - Probel_Posle_Imeni_Mama); //Копирование (Отчество)
            Otch_Papa = Name_Otch_Papa.Substring(Probel_Posle_Imeni_Papa, Name_Otch_Papa.Length - Probel_Posle_Imeni_Papa); //Копирование (Отчество)

            Name_Mama = Name_Otch_Mama.Remove(Probel_Posle_Imeni_Mama); //Копируем имя преподователя в переменную
            Name_Papa = Name_Otch_Papa.Remove(Probel_Posle_Imeni_Papa); //Копируем имя преподователя в переменную

            Fam_Mama = Fam_Mama.Replace("  ", string.Empty);
            Name_Mama = Name_Mama.Replace("  ", string.Empty);
            Otch_Mama = Otch_Mama.Replace("  ", string.Empty);

            Fam_Papa = Fam_Papa.Replace("  ", string.Empty);
            Name_Papa = Name_Papa.Replace("  ", string.Empty);
            Otch_Papa = Otch_Papa.Replace("  ", string.Empty);
            //(Familiya + ' ' + Imya + ' ' + Otchestvo)

            string queryStringMama = "SELECT * FROM Parents WHERE Familiya = '" + Fam_Mama + "' AND Imya = '" + Name_Mama + "' AND Otchestvo = '" + 
                Otch_Mama + "' AND Pol ='Женский'";
            string queryStringPapa = "SELECT * FROM Parents WHERE Familiya = '" + Fam_Papa + "' AND Imya = '" + Name_Papa + "' AND Otchestvo = '" +
                Otch_Papa + "' AND Pol ='Мужской'";
            string queryStringRebenok = "SELECT * FROM Childrens WHERE ID_rebenka = " + idRebenok;

            conAccess.QueryToDataGridOneRecord(queryStringPapa, dataGridPapa, "Папа");
            dataGridSettings(dataGridPapa);
            binSourcePapa = conAccess.binSourcePapa;

            conAccess.QueryToDataGridOneRecord(queryStringMama, dataGridMama, "Мама");
            dataGridSettings(dataGridMama);
            binSourceMama = conAccess.binSourceMama;

            conAccess.QueryToDataGridOneRecord(queryStringRebenok, dataGridDeti, "Ребенок");
            dataGridSettings(dataGridDeti);
            binSourceDeti = conAccess.binSourceRebenok;

            //Прогрузка в текстовые поля ребенка
            int x = 0;
            textIdDeti.Text = Convert.ToString(dataGridDeti[0, x].Value);
            textImyaDeti.Text = Convert.ToString(dataGridDeti[1, x].Value);
            textFamDeti.Text = Convert.ToString(dataGridDeti[2, x].Value);
            textOtchDeti.Text = Convert.ToString(dataGridDeti[3, x].Value);
            textDataRogdDeti.Text = Convert.ToString(dataGridDeti[4, x].Value);
            textAdresDeti.Text = Convert.ToString(dataGridDeti[5, x].Value);
            textClassDeti.Text = Convert.ToString(dataGridDeti[6, x].Value);
            textBukvaDeti.Text = Convert.ToString(dataGridDeti[7, x].Value);
            textFormObdeti.Text = Convert.ToString(dataGridDeti[8, x].Value);
            textDataZachDeti.Text = Convert.ToString(dataGridDeti[9, x].Value);
            textNomerPrikDeti.Text = Convert.ToString(dataGridDeti[10, x].Value);
            textDataOkDeti.Text = Convert.ToString(dataGridDeti[11, x].Value);
            textPrichOtchDeti.Text = Convert.ToString(dataGridDeti[12, x].Value);
            textKudaDeti.Text = Convert.ToString(dataGridDeti[13, x].Value);
            textClassnayaDeti.Text = Convert.ToString(dataGridDeti[14, x].Value);
            textSemiaDeti.Text = Convert.ToString(dataGridDeti[15, x].Value);
            textNomerPrikOtDeti.Text = Convert.ToString(dataGridDeti[16, x].Value);
            textStatusDeti.Text = Convert.ToString(dataGridDeti[17, x].Value);
            textDomTelDeti.Text = Convert.ToString(dataGridDeti[18, x].Value);
            textMobTelDeti.Text = Convert.ToString(dataGridDeti[19, x].Value);
            textDrugTelDeti.Text = Convert.ToString(dataGridDeti[20, x].Value);
            textMedPokazaniaDeti.Text = Convert.ToString(dataGridDeti[21, x].Value);

            //Прогрузка в текстовые поля мамы
            textIdMama.Text = Convert.ToString(dataGridMama[0, x].Value);
            textFamMama.Text = Convert.ToString(dataGridMama[1, x].Value);
            textImyaMama.Text = Convert.ToString(dataGridMama[2, x].Value);
            textOtchMama.Text = Convert.ToString(dataGridMama[3, x].Value);
            textPolMama.Text = Convert.ToString(dataGridMama[4, x].Value);
            textVozrMama.Text = Convert.ToString(dataGridMama[5, x].Value);
            textAdresMama.Text = Convert.ToString(dataGridMama[9, x].Value);
            textDomTelMama.Text = Convert.ToString(dataGridMama[7, x].Value);
            textMobTelMama.Text = Convert.ToString(dataGridMama[6, x].Value);
            textRabTelMama.Text = Convert.ToString(dataGridMama[8, x].Value);
            textMestoRabMama.Text = Convert.ToString(dataGridMama[10, x].Value);
            textDolgnMama.Text = Convert.ToString(dataGridMama[11, x].Value);
            textSemyaMama.Text = Convert.ToString(dataGridMama[12, x].Value);

            //Прогрузка в текстовые поля папы
            textIdPapa.Text = Convert.ToString(dataGridPapa[0, x].Value);
            textFamPapa.Text = Convert.ToString(dataGridPapa[1, x].Value);
            textImyaPapa.Text = Convert.ToString(dataGridPapa[2, x].Value);
            textOtchPapa.Text = Convert.ToString(dataGridPapa[3, x].Value);
            textPolPapa.Text = Convert.ToString(dataGridPapa[4, x].Value);
            textVozrPapa.Text = Convert.ToString(dataGridPapa[5, x].Value);
            textAdresPapa.Text = Convert.ToString(dataGridPapa[9, x].Value);
            textDomTelPapa.Text = Convert.ToString(dataGridPapa[7, x].Value);
            textMobTelPapa.Text = Convert.ToString(dataGridPapa[6, x].Value);
            textRabTelPapa.Text = Convert.ToString(dataGridPapa[8, x].Value);
            textMestoRabPapa.Text = Convert.ToString(dataGridPapa[10, x].Value);
            textDolgnPapa.Text = Convert.ToString(dataGridPapa[11, x].Value);
            textSemyaPapa.Text = Convert.ToString(dataGridPapa[12, x].Value);
        }

        //Кнопка "Распечатать", для печати анкеты
        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Документ (*.doc)|*.doc|Все файлы (*.*)|*.*";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                StreamWriter streamWriter = new StreamWriter(saveFileDialog.FileName);
                streamWriter.WriteLine("Анкета семьи");
                streamWriter.WriteLine("РЕБЕНОК");
                streamWriter.WriteLine("ФИО ребенка: " + textImyaDeti.Text + " " + textFamDeti.Text + " " + textOtchDeti.Text);
                streamWriter.WriteLine("Дата рождения: " + textDataRogdDeti.Text);
                streamWriter.WriteLine("Адрес: " + textAdresDeti.Text);
                streamWriter.WriteLine("Класс: " + textClassDeti.Text);
                streamWriter.WriteLine("Буква: " + textBukvaDeti.Text);
                streamWriter.WriteLine("Форма обучения: " + textFormObdeti.Text);
                streamWriter.WriteLine("Дата зачисления: " + textDataZachDeti.Text);
                streamWriter.WriteLine("Номер приказа зачисления: " + textNomerPrikDeti.Text);
                streamWriter.WriteLine("Дата окончания: " + textDataOkDeti.Text);
                streamWriter.WriteLine("Причина отчисления: " + textPrichOtchDeti.Text);
                streamWriter.WriteLine("Куда выбыл: " + textKudaDeti.Text);
                streamWriter.WriteLine("ФИО Классного руководителя: " + textClassnayaDeti.Text);
                streamWriter.WriteLine("Семья: " + textSemiaDeti.Text);
                streamWriter.WriteLine("Номер приказа отчисления: " + textNomerPrikOtDeti.Text);
                streamWriter.WriteLine("Социальный статус ребенка: " + textStatusDeti.Text);
                //Мама
                streamWriter.WriteLine("МАМА");
                streamWriter.WriteLine("ФИО Мамы: " + textFamMama.Text + " " + textImyaMama.Text + " " + textOtchMama.Text);
                streamWriter.WriteLine("Пол Мамы: " + textPolMama.Text);
                streamWriter.WriteLine("Возраст Мамы: " + textVozrMama.Text);
                streamWriter.WriteLine("Адрес Мамы: " + textAdresMama.Text);
                streamWriter.WriteLine("Домашний телефон Мамы: " + textDomTelMama.Text);
                streamWriter.WriteLine("Мобильный телефон Мамы: " + textMobTelMama.Text);
                streamWriter.WriteLine("Рабочий телефон Мамы: " + textRabTelMama.Text);
                streamWriter.WriteLine("Место работы Мамы: " + textMestoRabMama.Text);
                streamWriter.WriteLine("Должность Мамы: " + textDolgnMama.Text);
                //Папа
                streamWriter.WriteLine("ПАПА");
                streamWriter.WriteLine("ФИО Папы: " + textFamPapa.Text + " " + textImyaPapa.Text + " " + textOtchPapa.Text);
                streamWriter.WriteLine("Пол Папы: " + textPolPapa.Text);
                streamWriter.WriteLine("Возраст Папы: " + textVozrPapa.Text);
                streamWriter.WriteLine("Адрес Папы: " + textAdresPapa.Text);
                streamWriter.WriteLine("Домашний телефон Папы: " + textDomTelPapa.Text);
                streamWriter.WriteLine("Мобильный телефон Папы: " + textMobTelPapa.Text);
                streamWriter.WriteLine("Рабочий телефон Папы: " + textRabTelPapa.Text);
                streamWriter.WriteLine("Место работы Папы: " + textMestoRabPapa.Text);
                streamWriter.WriteLine("Должность Папы: " + textDolgnPapa.Text);


                streamWriter.Close();
                MessageBox.Show("Файл был сохранен!","Выполненно!");
            }
        }
    }
}
