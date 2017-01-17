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
    public partial class Splashscreen : Form
    {
        public ConnectorAccess conAccess;
        public Splashscreen(ConnectorAccess ClassConSQL)
        {
            InitializeComponent();
            this.conAccess = ClassConSQL;
        }

        public string con_splash_screen;

        //Процедура загрузки формы
        private void Splashscreen_Load(object sender, EventArgs e)
        {
            ConLabel.Text = "Проверка подключения...";
            timer1.Interval = 1000;
            timer1.Start();
        }

        //Процедура работы с таймером, кароче соединение
        private void timer1_Tick(object sender, EventArgs e)
        {
            //Если существует файл настроек подключения
            if (File.Exists(@"Data\\MyAccessCon.cfg"))
            {
                timer1.Stop();
                //Main main = new Main(conAccess);
                Autorization vhod = new Autorization(conAccess);
                StreamReader sreader = new StreamReader(@"Data\\MyAccessCon.cfg");
                con_splash_screen = sreader.ReadLine(); //читаем данные о подключении в переменную 
                sreader.Close();
                sreader.Dispose();
                if (conAccess.Connection(con_splash_screen) == true)
                {
                    timer1.Stop();
                    vhod.Show();
                    //main.Show();
                    this.Hide();
                }
                else
                {
                    timer1.Stop();
                    ConLabel.Text = "Ошибка подключения...";                    
                    if (MessageBox.Show("Не удалось подключиться к базе данных, настроить подключение?",
                        "Ошибка", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        Settings Stngs = new Settings(conAccess);
                        Stngs.Show();
                        Stngs.button3.Enabled = false;
                        Stngs.button1.Enabled = true;
                        this.Hide();
                    }
                    else
                    {
                        Application.Exit();
                    }
                }
            }
            else
            {
                StreamWriter wreader = new StreamWriter(@"Data\\MyAccessCon.cfg");
                wreader.Close();
            }
        }
    }
}
