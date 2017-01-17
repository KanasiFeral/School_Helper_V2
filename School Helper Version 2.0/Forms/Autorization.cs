using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;

namespace School_Helper_Version_2._0.Forms
{
    public partial class Autorization : Form
    {
        public string sPassword;
        public bool bIsClose = true;
        public ConnectorAccess conAccess;
        public Autorization(ConnectorAccess ClassConSQL)
        {
            InitializeComponent();
            this.conAccess = ClassConSQL;
        }

        //Загрузка формы
        private void Autorization_Load(object sender, EventArgs e)
        {
            textBoxPassword.Visible = false;
            comboBoxTypeUsers.SelectedIndex = 0;
        }

        //Закрываем форму, закрываем приложение
        private void Autorization_FormClosed(object sender, FormClosedEventArgs e)
        {
            //Закрывать все приложение, только если флаг стоит в true
            if(bIsClose == true)
            {
                Application.Exit();
            }
        }

        //Процедура входа в программу
        public void vhod()
        {
            if (comboBoxTypeUsers.SelectedIndex == 0)
            {
                StreamReader sreader = new StreamReader(@"Data\\password.txt");
                sPassword = sreader.ReadLine(); //читаем данные о пароле в переменную 
                sreader.Close();
                sreader.Dispose();

                //Проверка на правильность
                while (sPassword != textBoxPassword.Text)
                {
                    MessageBox.Show("Пароль не верен, попробуйте снова!");
                    return;
                }

                //Ставим флаг закрытия приложения в false
                bIsClose = false;
                Main main = new Main(conAccess);
                main.bAdminStatus = true;
                main.Show();
                this.Close();
            }
            else
            {
                //Ставим флаг закрытия приложения в false
                bIsClose = false;
                Main main = new Main(conAccess);
                main.bAdminStatus = false;
                main.Show();
                this.Close();
            }
        }

        //Кнопка входа
        private void buttonEnter_Click(object sender, EventArgs e)
        {
            vhod();
        }

        //Происходит при смене элемента в списке
        private void comboBoxTypeUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Если выбран админ
            if(comboBoxTypeUsers.SelectedIndex == 0)
            {
                textBoxPassword.Visible = true;
            }
            else
            {
                textBoxPassword.Visible = false;
            }
        }

        private void textBoxPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                vhod();
            }
        }
    }
}