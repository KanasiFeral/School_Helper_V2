using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;

namespace School_Helper_Version_2._0.Forms
{
    public class ConnectorAccess
    {
        public BindingSource binSourceDeti = new BindingSource();
        public BindingSource binSourceParents = new BindingSource();
        public BindingSource binSourceClassnaya = new BindingSource();
        public BindingSource binSourceClass = new BindingSource();

        public BindingSource binSourcePapa = new BindingSource();
        public BindingSource binSourceMama = new BindingSource();
        public BindingSource binSourceRebenok = new BindingSource();

        public BindingSource binSourceEvents = new BindingSource();
        public BindingSource binSourceOlympics = new BindingSource();

        //Указываем в памяти место для хранения значений типа OdbcConnection
        public OleDbConnection con_ConnectionAccess;
        //Процедура подключения к базе данных
        public bool Connection(string Connection)
        {
            try
            {
                con_ConnectionAccess = new OleDbConnection(Connection);
                con_ConnectionAccess.Open();
            }
            catch(Exception exp)
            {
                MessageBox.Show(exp.Message);
                return false;
            }
            return true;
        }
        //Процедура закрытия соеденения с базой данных
        public bool CloseConnection()
        {
            con_ConnectionAccess.Close();
            return true;
        }
        //Процедура запросов к базе, вернет true, если имеются строки в таблице
        public bool QueryToBool(string queryString)
        { 
            OleDbCommand com;
            OleDbDataReader dataReader;
            com = new OleDbCommand(queryString, con_ConnectionAccess);
            try
            {
                dataReader = com.ExecuteReader();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
                return false;
            }
            if (dataReader.HasRows)
            {
                dataReader.Close();
                com.Dispose();
                return true;
            }
            dataReader.Close();
            com.Dispose();
            return false;
        }

        //Процедура вывода таблицы в dataGridView
        public bool QueryToDataGrid(string queryString, DataGridView DataGrid, BindingNavigator Navigator, string Name_BinSource)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();
            OleDbCommand com;
            OleDbDataReader dataReader;
            com = new OleDbCommand(queryString, con_ConnectionAccess);
            try
            {
                dataReader = com.ExecuteReader();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
                return false;
            }
            if (dataReader.HasRows)
            {
                dataTable.Load(dataReader);
                if (Name_BinSource == "Дети")//Таблицы Дети
                {
                    binSourceDeti.DataSource = dataTable;
                    Navigator.BindingSource = binSourceDeti;
                    DataGrid.DataSource = binSourceDeti;
                }
                else if (Name_BinSource == "Родители") //Таблицы Родители
                {
                    binSourceParents.DataSource = dataTable;
                    Navigator.BindingSource = binSourceParents;
                    DataGrid.DataSource = binSourceParents;
                }
                else if (Name_BinSource == "Классный руководитель") //Таблица Классный руководитель
                {
                    binSourceClassnaya.DataSource = dataTable;
                    Navigator.BindingSource = binSourceClassnaya;
                    DataGrid.DataSource = binSourceClassnaya;
                }
                else if(Name_BinSource == "События") //Таблица События
                {
                    binSourceEvents.DataSource = dataTable;
                    Navigator.BindingSource = binSourceEvents;
                    DataGrid.DataSource = binSourceEvents;
                }
                else if (Name_BinSource == "Олимпиады") //Таблица События
                {
                    binSourceOlympics.DataSource = dataTable;
                    Navigator.BindingSource = binSourceOlympics;
                    DataGrid.DataSource = binSourceOlympics;
                }
                else //Таблица Класс
                {
                    binSourceClass.DataSource = dataTable;
                    Navigator.BindingSource = binSourceClass;
                    DataGrid.DataSource = binSourceClass;
                }


                dataReader.Close();
                com.Dispose();
                return true;

            }
            dataReader.Close();
            com.Dispose();
            return false;
        }

        //Процедура вывода таблицы в dataGridView
        public bool QueryToDataGridOneRecord(string queryString, DataGridView DataGrid, string Name_BinSource)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();
            OleDbCommand com;
            OleDbDataReader dataReader;
            com = new OleDbCommand(queryString, con_ConnectionAccess);
            try
            {
                dataReader = com.ExecuteReader();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
                return false;
            }
            if (dataReader.HasRows)
            {
                dataTable.Load(dataReader);
                if (Name_BinSource == "Ребенок")//Таблицы Дети
                {
                    binSourceRebenok.DataSource = dataTable;
                    DataGrid.DataSource = binSourceRebenok;
                }
                else if (Name_BinSource == "Мама") //Таблицы Родители
                {
                    binSourceMama.DataSource = dataTable;
                    DataGrid.DataSource = binSourceMama;
                }
                else //Таблицы Родители
                {
                    binSourcePapa.DataSource = dataTable;
                    DataGrid.DataSource = binSourcePapa;
                }


                dataReader.Close();
                com.Dispose();
                return true;

            }
            dataReader.Close();
            com.Dispose();
            return false;
        }


        //Процедура вывода в комбо бокса столбца таблицы
        public bool QueryToComboBox(string queryString, ComboBox comboBox, string Name_Column)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();
            OleDbCommand com;
            OleDbDataReader dataReader;
            com = new OleDbCommand(queryString, con_ConnectionAccess);
            try
            {
                dataReader = com.ExecuteReader();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
                return false;
            }
            if (dataReader.HasRows)
            {
                dataTable.Load(dataReader);
                comboBox.DataSource = dataTable;
                comboBox.DisplayMember = Name_Column;

                dataReader.Close();
                com.Dispose();
                return true;
            }
            dataReader.Close();
            com.Dispose();
            return false;
        }

        //Процедура для агрегатных запросов
        public string AgregateQueryToDataGrid(string queryString)
        {
            int iResultQuery = 0;
            string sResultQuery = "";
            System.Data.DataTable dataTable = new System.Data.DataTable();
            OleDbCommand com;
            OleDbDataReader dataReader;
            com = new OleDbCommand(queryString, con_ConnectionAccess);
            try
            {
                dataReader = com.ExecuteReader();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return sResultQuery;
            }

            dataReader.Read();
            //resultQuery = dataReader.GetString(0); //Invalid attempt to access a field before calling Read()
            /*try
            {
                if (dataReader.HasRows)
                {
                    dataTable.Load(dataReader);
                    var MyValue = Convert.ToString(dataTable.Rows[0][0]);
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return resultQuery;
            }*/
            iResultQuery = dataReader.GetInt32(0);
            sResultQuery = Convert.ToString(iResultQuery);

            dataReader.Close();
            com.Dispose();
            return sResultQuery;
        }

    }
}