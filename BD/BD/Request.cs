using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace BD
{
    public partial class Request : Form
    {
        public Request()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            String connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);

            dataGridView2.Rows.Clear();
            this.Column10.HeaderText = "Идентификатор товара";
            this.Column11.HeaderText = "Цена";
            this.Column12.HeaderText = "";
            this.Column13.HeaderText = "";
            this.Column14.HeaderText = "";
            this.Column15.HeaderText = "";

            dbConnection.Open();// открываем соединение
            try
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = dbConnection;
                cmd.CommandText = "EXEC [Самые дорогие товары]";
                cmd.ExecuteNonQuery();
                OleDbDataReader dbReader = cmd.ExecuteReader();

                if (dbReader.HasRows == false)
                    MessageBox.Show("Ошибка");
                else
                    while (dbReader.Read())
                    {
                        dataGridView2.Rows.Add(dbReader.GetValue(dbReader.GetOrdinal("Идентификатор товара")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Цена")).ToString());
                    }

                dbReader.Close();
                dbConnection.Close();
            }
            catch
            {
                MessageBox.Show("Ошибка");
                dbConnection.Close();
            }
            finally
            {
                dbConnection.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);

            dataGridView2.Rows.Clear();
            this.Column10.HeaderText = "Идентификатор товара";
            this.Column11.HeaderText = "Номер магазина";
            this.Column12.HeaderText = "Номер отдела";
            this.Column13.HeaderText = "Цена";
            this.Column14.HeaderText = "Количество";
            this.Column15.HeaderText = "Срок годности";

            dbConnection.Open();// открываем соединение
            try
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = dbConnection;
                cmd.CommandText = "EXEC [SQL Истекший срок годности]";
                cmd.ExecuteNonQuery();
                OleDbDataReader dbReader = cmd.ExecuteReader();

                if (dbReader.HasRows == false)
                    MessageBox.Show("Ошибка");
                else
                    while (dbReader.Read())
                    {
                        dataGridView2.Rows.Add(dbReader.GetValue(dbReader.GetOrdinal("Идентификатор товара")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Номер магазина")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Номер отдела")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Цена")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Количество")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Срок годности")).ToString());
                    }

                dbReader.Close();
                dbConnection.Close();
            }
            catch
            {
                MessageBox.Show("Ошибка");
                dbConnection.Close();
            }
            finally
            {
                dbConnection.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            String connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);

            dataGridView2.Rows.Clear();
            this.Column10.HeaderText = "Номер отдела";
            this.Column11.HeaderText = "Номер магазина";
            this.Column12.HeaderText = "Название отдела";
            this.Column13.HeaderText = "Табельный номер заведующего";
            this.Column14.HeaderText = "";
            this.Column15.HeaderText = "";

            dbConnection.Open();// открываем соединение
            try
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = dbConnection;
                cmd.CommandText = "EXEC [SQL Магазин-Отделы]";
                cmd.ExecuteNonQuery();
                OleDbDataReader dbReader = cmd.ExecuteReader();

                if (dbReader.HasRows == false)
                    MessageBox.Show("Ошибка");
                else
                    while (dbReader.Read())
                    {
                        dataGridView2.Rows.Add(dbReader.GetValue(dbReader.GetOrdinal("Номер отдела")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Номер магазина")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Название отдела")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Табельный номер заведующего")).ToString());
                    }

                dbReader.Close();
                dbConnection.Close();
            }
            catch
            {
                MessageBox.Show("Ошибка");
                dbConnection.Close();
            }
            finally
            {
                dbConnection.Close();
            }
        }



        private void button5_Click(object sender, EventArgs e)
        {
            String connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);

            dataGridView2.Rows.Clear();
            this.Column10.HeaderText = "Табельный номер сотрудника";
            this.Column11.HeaderText = "Фамилия";
            this.Column12.HeaderText = "Название магазина";
            this.Column13.HeaderText = "Название отдела";
            this.Column14.HeaderText = "";
            this.Column15.HeaderText = "";

            dbConnection.Open();// открываем соединение
            try
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = dbConnection;
                cmd.Parameters.AddWithValue("Введите номер сотрудника", textBox1.Text.ToString());
                cmd.CommandText = "EXEC [Поиск сотрудника по табельному]";
                cmd.ExecuteNonQuery();
                OleDbDataReader dbReader = cmd.ExecuteReader();

                if (dbReader.HasRows == false)
                    MessageBox.Show("Ошибка");
                else
                    while (dbReader.Read())
                    {
                        dataGridView2.Rows.Add(dbReader.GetValue(dbReader.GetOrdinal("Табельный номер сотрудника")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Фамилия")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Название магазина")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Название отдела")).ToString());
                    }

                dbReader.Close();
                dbConnection.Close();
            }
            catch
            {
                MessageBox.Show("Ошибка");
                dbConnection.Close();
            }
            finally
            {
                dbConnection.Close();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            String connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);

            dataGridView2.Rows.Clear();
            this.Column10.HeaderText = "Идентификатор товара";
            this.Column11.HeaderText = "Номер магазина";
            this.Column12.HeaderText = "Номер отдела";
            this.Column13.HeaderText = "Цена";
            this.Column14.HeaderText = "";
            this.Column15.HeaderText = "";

            dbConnection.Open();// открываем соединение
            try
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = dbConnection;
                cmd.Parameters.AddWithValue("Введите стоимость", textBox2.Text.ToString());
                cmd.CommandText = "EXEC [ТоварыВышеСтоимости]";
                cmd.ExecuteNonQuery();
                OleDbDataReader dbReader = cmd.ExecuteReader();

                if (dbReader.HasRows == false)
                    MessageBox.Show("Ошибка");
                else
                    while (dbReader.Read())
                    {
                        dataGridView2.Rows.Add(dbReader.GetValue(dbReader.GetOrdinal("Идентификатор товара")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Номер магазина")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Номер отдела")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Цена")).ToString());
                    }

                dbReader.Close();
                dbConnection.Close();
            }
            catch
            {
                MessageBox.Show("Ошибка");
                dbConnection.Close();
            }
            finally
            {
                dbConnection.Close();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            String connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();// открываем соединение
            try
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = dbConnection;
                cmd.Parameters.AddWithValue("Номер поставщика", textBox3.Text.ToString());
                cmd.Parameters.AddWithValue("Название поставщика", textBox5.Text.ToString());
                cmd.Parameters.AddWithValue("Адрес поставщика", textBox4.Text.ToString());
                cmd.CommandText = "EXEC [Добавление поставщик]";
                cmd.ExecuteNonQuery();
                
                dbConnection.Close();
            }
            catch
            {
                MessageBox.Show("Ошибка");
                dbConnection.Close();
            }
            finally
            {
                MessageBox.Show("Данные успешно добавлены");
                dbConnection.Close();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            String connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();// открываем соединение
            try
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = dbConnection;
                cmd.Parameters.AddWithValue("Введите номер отдела", textBox6.Text.ToString());
                cmd.CommandText = "EXEC [Удаление с подзапросом всех товаров из отдела]";
                cmd.ExecuteNonQuery();
               
            }
            catch
            {
                MessageBox.Show("Ошибка");
                dbConnection.Close();
            }
            finally
            {
                MessageBox.Show("Данные успешно удалены");
                dbConnection.Close();
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            OleDbConnection dbConnection = new OleDbConnection(
                @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source=""lab3.accdb"";" + 
                @"Jet OLEDB:Create System Database=true;" +
                @"Jet OLEDB:System database=C:\Users\roman\AppData\" +
                @"Roaming\Microsoft\Access\System.mdw"
            );

            

            dataGridView2.Rows.Clear();
            this.Column10.HeaderText = "Id";
            this.Column11.HeaderText = "Name";
            this.Column12.HeaderText = "ParentId";
            this.Column13.HeaderText = "Type";
            this.Column14.HeaderText = "";
            this.Column15.HeaderText = "";

            dbConnection.Open(); // открываем соединение

            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = dbConnection;
            cmd.CommandText = "EXEC [MSysObjects отображение]";
            cmd.ExecuteNonQuery();
            OleDbDataReader dbReader = cmd.ExecuteReader();

            if (dbReader.HasRows == false)
                MessageBox.Show("Ошибка");
            else
                while (dbReader.Read())
                {
                    dataGridView2.Rows.Add(dbReader.GetValue(dbReader.GetOrdinal("Id")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Name")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("ParentId")).ToString(), dbReader.GetValue(dbReader.GetOrdinal("Type")).ToString());
                }

            dbReader.Close();
            dbConnection.Close();
        }
    }
}
