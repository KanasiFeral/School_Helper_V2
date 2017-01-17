using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace School_Helper_Version_2._0.Forms
{
    public partial class Settings : Form
    {
        public ConnectorAccess conAccess;
        public Settings(ConnectorAccess ClassConSQL)
        {
            InitializeComponent();
            this.conAccess = ClassConSQL;
        }

        //Кнопка "Отменить", для отмены действий
        private void button1_Click(object sender, EventArgs e)
        {
            Splashscreen SplashScreen = new Splashscreen(conAccess);
            SplashScreen.Show();
            this.Close();
        }

        //Кнопка "Создать соединение", для добавления настроек в конфиг файл подключения
        private void button2_Click(object sender, EventArgs e)
        {
            if (DBTextBox.Text.Equals("") || (ServerTextBox.Text.Equals("")))
            {
                MessageBox.Show("Вы не ввели все данные!", "Предупреждение!",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                StreamWriter wreader = new StreamWriter(@"Data\\MyAccessCon.cfg");
                wreader.WriteLine(@"Provider="+ServerTextBox.Text + @";Data Source=" + DBTextBox.Text + "");
                wreader.Close();
                wreader.Dispose();
                MessageBox.Show("Подключение было созданно!", "Созданно!");
                Splashscreen SP = new Splashscreen(conAccess);
                SP.Show();
                this.Close();
            }
        }

        //Кнопка "Выход в главное меню", чтобы закрыть форму настроек
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
            Main main = new Main(conAccess);
            main.Show();
        }

        //Кнопка "Указать путь через OpenFileDialog", для указания путя к файлу БД
        private void button4_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "DataBase Files (.mdb)|*.*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string PathBD = openFileDialog1.FileName;
                DBTextBox.Text = PathBD;
            }
        }
    }
}
