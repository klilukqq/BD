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
    public partial class Form1 : Form
    {



        public Form1()
        {
            InitializeComponent();
        }

        private void ЗагрузитьПоставщикToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();

            this.Column1.HeaderText = "Номер поставщика";
            this.Column2.HeaderText = "Название поставщика";
            this.Column3.HeaderText = "Адрес поставщика";
            this.Column4.HeaderText = "";
            this.Column5.HeaderText = "";
            this.Column6.HeaderText = "";
            this.Column7.HeaderText = "";
            this.Column8.HeaderText = "";
            this.Column9.HeaderText = "";
            this.Column10.HeaderText = "";

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "SELECT * FROM Поставщик";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();

            if (dbReader.HasRows == false)
            {
                MessageBox.Show("Данные не найдены");
            }
            else
            {

                while (dbReader.Read())
                {


                    dataGridView1.Rows.Add(dbReader["Номер поставщика"], dbReader["Название поставщика"], dbReader["Адрес поставщика"]);

                }
            }
            dbReader.Close();
            dbConnection.Close();
        }

        private void ДобавитьПоставщикToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;
            if (dataGridView1.Rows[index].Cells[0].Value == null ||
            dataGridView1.Rows[index].Cells[1].Value == null ||

            dataGridView1.Rows[index].Cells[2].Value == null 
            )
            {
                MessageBox.Show("Введены не все данные!");
                return;
            }

            string numpost = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string name = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string address = dataGridView1.Rows[index].Cells[2].Value.ToString();
           

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "INSERT INTO Поставщик VALUES (" + numpost + ", '" + name + "', '" + address + "')";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void ОбновитьПоставщикToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;

            string numpost = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string name = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string address = dataGridView1.Rows[index].Cells[2].Value.ToString();


            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "UPDATE [Поставщик] SET [Название поставщика]='" + name + "',[Адрес поставщика]= '" + address + "' WHERE [Номер поставщика]=" + numpost;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void УдалитьПоставщикToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;

            string numpost = dataGridView1.Rows[index].Cells[0].Value.ToString();


            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "DELETE FROM Поставщик WHERE [Номер поставщика]=" + numpost;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void ЗагрузитьДоговорыToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();

            this.Column1.HeaderText = "Номер договора";
            this.Column2.HeaderText = "Дата";
            this.Column3.HeaderText = "Номер поставщика";
            this.Column4.HeaderText = "Номер магазина";
            this.Column5.HeaderText = "";
            this.Column6.HeaderText = "";
            this.Column7.HeaderText = "";
            this.Column8.HeaderText = "";
            this.Column9.HeaderText = "";
            this.Column10.HeaderText = "";
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "SELECT * FROM Договоры";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();

            if (dbReader.HasRows == false)
            {
                MessageBox.Show("Данные не найдены");
            }
            else
            {

                while (dbReader.Read())
                {


                    dataGridView1.Rows.Add(dbReader["Номер договора"], dbReader["Дата"], dbReader["Номер поставщика"], dbReader["Номер магазина"]);

                }
            }
            dbReader.Close();
            dbConnection.Close();
        }

        private void ДобавитьДоговорыToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;
            if (dataGridView1.Rows[index].Cells[0].Value == null ||
            dataGridView1.Rows[index].Cells[1].Value == null ||

            dataGridView1.Rows[index].Cells[2].Value == null ||
            dataGridView1.Rows[index].Cells[3].Value == null 
            )
            {
                MessageBox.Show("Введены не все данные!");
                return;
            }

            string numdog = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string date = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string numpost = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string numshop = dataGridView1.Rows[index].Cells[3].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "INSERT INTO Договоры VALUES (" + numdog + ", '" + date + "', '" + numpost + "', '" + numshop + "')";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();

        }

        private void ОбновитьДоговорыToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;

            string numdog = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string date = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string numpost = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string numshop = dataGridView1.Rows[index].Cells[3].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "UPDATE [Договоры] SET [Дата]='" + date + "',[Номер поставщика]= '" + numpost + "',[Номер магазина]= '" + numshop + "'WHERE [Номер договора]="+ numdog;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void УдалитьДоговорыToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;

            string numdog = dataGridView1.Rows[index].Cells[0].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "DELETE FROM Договоры WHERE [Номер договора]=" + numdog;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void ЗагрузитьМагазинToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            this.Column1.HeaderText = "Номер магазина";
            this.Column2.HeaderText = "Название магазина";
            this.Column3.HeaderText = "Специализация";
            this.Column4.HeaderText = "ИНН";
            this.Column5.HeaderText = "Адрес";
            this.Column6.HeaderText = "Табельный номер директора";
            this.Column7.HeaderText = "";
            this.Column8.HeaderText = "";
            this.Column9.HeaderText = "";
            this.Column10.HeaderText = "";

            string connectionString = "provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "SELECT * FROM Магазин";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();//считываем данные

            if (dbReader.HasRows == false)
            {
                MessageBox.Show("Данные не найдены");
            }
            else
            {

                while (dbReader.Read())
                {

                    //Вывод данных
                    dataGridView1.Rows.Add(dbReader["Номер магазина"], dbReader["Название магазина"], dbReader["Специализация"], dbReader["ИНН"], dbReader["Адрес"], dbReader["Табельный номер директора"]);

                }
            }
            dbReader.Close();
            dbConnection.Close();


        }

        private void ДобавитьМагазинToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;
            if (dataGridView1.Rows[index].Cells[0].Value == null ||
            dataGridView1.Rows[index].Cells[1].Value == null ||
            dataGridView1.Rows[index].Cells[2].Value == null ||
            dataGridView1.Rows[index].Cells[3].Value == null ||
            dataGridView1.Rows[index].Cells[4].Value == null ||
            dataGridView1.Rows[index].Cells[5].Value == null)
            {
                MessageBox.Show("Введены не все данные!");
                return;
            }

            string nomshop = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string name = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string spec = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string inn = dataGridView1.Rows[index].Cells[3].Value.ToString();
            //messagebox.showMessageBox.Show


            dataGridView1.Rows[index].Cells[3].Value.ToString();
            string adr = dataGridView1.Rows[index].Cells[4].Value.ToString();
            string nomshopdir = dataGridView1.Rows[index].Cells[5].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "INSERT INTO МАГАЗИН VALUES (" + nomshop + ", '" + name + "', '" + spec + "', " + inn + ", '" + adr + "', " + nomshopdir + " )";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void ОбновитьМагазинToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;
            

            string nomshop = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string name = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string spec = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string inn = dataGridView1.Rows[index].Cells[3].Value.ToString();
            string adr = dataGridView1.Rows[index].Cells[4].Value.ToString();
            string nomshopdir = dataGridView1.Rows[index].Cells[5].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb;Jet OLEDB:Create System Database=true";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "UPDATE [МАГАЗИН] SET [Название магазина] = '" + name + "', [ИНН] = '" + inn + "', [Специализация] = '" + spec + "', [Адрес] = '" + adr + "', [Табельный номер директора] = " + nomshopdir + " WHERE [Номер магазина] =" + nomshop;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные изменены (обновлены)");

            dbConnection.Close();




        }

        private void УдалитьМагазинToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;

            string nomshop = dataGridView1.Rows[index].Cells[0].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "DELETE FROM МАГАЗИН WHERE [Номер магазина]=" + nomshop;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();

        }

        private void ЗагрузитьОтделToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            this.Column1.HeaderText = "Номер магазина";
            this.Column2.HeaderText = "Номер отдела";
            this.Column3.HeaderText = "Название отдела";
            this.Column4.HeaderText = "Табельный номер заведующего ";
            this.Column5.HeaderText = "";
            this.Column6.HeaderText = "";
            this.Column7.HeaderText = "";
            this.Column8.HeaderText = "";
            this.Column9.HeaderText = "";
            this.Column10.HeaderText = "";

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "SELECT * FROM Отдел";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();

            if (dbReader.HasRows == false)
            {
                MessageBox.Show("Данные не найдены");
            }
            else
            {

                while (dbReader.Read())
                {


                    dataGridView1.Rows.Add(dbReader["Номер магазина"], dbReader["Номер отдела"], dbReader["Название отдела"], dbReader["Табельный номер заведующего"]);

                }
            }
            dbReader.Close();
            dbConnection.Close();

        }

        private void ДобавитьОтделToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;
            if (dataGridView1.Rows[index].Cells[0].Value == null ||
            dataGridView1.Rows[index].Cells[1].Value == null ||
            dataGridView1.Rows[index].Cells[2].Value == null ||
            dataGridView1.Rows[index].Cells[3].Value == null
            )
            {
                MessageBox.Show("Введены не все данные!");
                return;
            }

            string nomshop = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string nomot = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string name = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string tabdir = dataGridView1.Rows[index].Cells[3].Value.ToString();


            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "INSERT INTO ОТДЕЛ VALUES (" + nomot + ", " + nomshop + ", '" + name + "', " + tabdir + " )";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();

        }

        private void ОбновитьОтделToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;
           

            string nomshop = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string nomot = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string name = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string tabdir = dataGridView1.Rows[index].Cells[3].Value.ToString();


            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "UPDATE [ОТДЕЛ] SET [Номер магазина]=" + nomshop  + ",[Название отдела]= '" + name + "',[Табельный номер заведующего]= " + tabdir + "WHERE [Номер отдела] =" + nomot;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void УдалитьОтделToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;

            string nomot = dataGridView1.Rows[index].Cells[1].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "DELETE FROM ОТДЕЛ WHERE [Номер отдела] =" + nomot;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void ЗагрузитьСотрудникToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();

            this.Column1.HeaderText = "Табельный номер сотрудника";
            this.Column2.HeaderText = "Номер магазина";
            this.Column3.HeaderText = "Фамилия";
            this.Column4.HeaderText = "Имя";
            this.Column5.HeaderText = "Отчество ";
            this.Column6.HeaderText = "Адрес";
            this.Column7.HeaderText = "Пол";
            this.Column8.HeaderText = "Дата рождения";
            this.Column9.HeaderText = "Семейное положение";
            this.Column10.HeaderText = "";




            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "SELECT * FROM Сотрудник";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();

            if (dbReader.HasRows == false)
            {
                MessageBox.Show("Данные не найдены");
            }
            else
            {

                while (dbReader.Read())
                {


                    dataGridView1.Rows.Add(dbReader["Табельный номер сотрудника"], dbReader["Номер магазина"], dbReader["Фамилия"], dbReader["Имя"], dbReader["Отчество"], dbReader["Адрес"], dbReader["Пол"], dbReader["Дата рождения"], dbReader["Семейное положение"]);
                
}
            }
            dbReader.Close();
            dbConnection.Close();
        }



        private void ДобавитьСотрудникToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;
            if (dataGridView1.Rows[index].Cells[0].Value == null ||
            dataGridView1.Rows[index].Cells[1].Value == null ||
            dataGridView1.Rows[index].Cells[2].Value == null ||
            dataGridView1.Rows[index].Cells[3].Value == null ||
            dataGridView1.Rows[index].Cells[4].Value == null ||
            dataGridView1.Rows[index].Cells[5].Value == null ||
            dataGridView1.Rows[index].Cells[6].Value == null ||
            dataGridView1.Rows[index].Cells[7].Value == null ||
            dataGridView1.Rows[index].Cells[8].Value == null 
            )
            {
                MessageBox.Show("Введены не все данные!");
                return;
            }

            string tab = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string numshop = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string surname = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string name = dataGridView1.Rows[index].Cells[3].Value.ToString();
            string otche = dataGridView1.Rows[index].Cells[4].Value.ToString();
            string adr = dataGridView1.Rows[index].Cells[5].Value.ToString();
            string sex = dataGridView1.Rows[index].Cells[6].Value.ToString();
            string bird = dataGridView1.Rows[index].Cells[7].Value.ToString();
            string family = dataGridView1.Rows[index].Cells[8].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "INSERT INTO СОТРУДНИК VALUES (" + tab + ", '" + numshop + "', '" + surname + "', '" + name + "' , '" + otche + "', '" + adr + "', '" + bird + "', '" + sex + "', '" + family  + "')";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();


        }

        private void ОбновитьСотрудникToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;
            

            string tab = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string numshop = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string surname = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string name = dataGridView1.Rows[index].Cells[3].Value.ToString();
            string otche = dataGridView1.Rows[index].Cells[4].Value.ToString();
            string adr = dataGridView1.Rows[index].Cells[5].Value.ToString();
            string sex = dataGridView1.Rows[index].Cells[6].Value.ToString();
            string bird = dataGridView1.Rows[index].Cells[7].Value.ToString();
            string family = dataGridView1.Rows[index].Cells[8].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "UPDATE [Сотрудник] SET [Номер магазина]='" + numshop + "',[Фамилия]= '" + surname + "',[Имя]= '" + name + "' ,[Отчество]= '" + otche + "',[Адрес]= '" + adr + "',[Дата рождения]= '" + bird + "',[Пол]= '" + sex + "',[Семейное положение]= '" + family + "' WHERE [Табельный номер сотрудника] = " + tab;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void УдалитьСотрудникToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;


            string tab = dataGridView1.Rows[index].Cells[0].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "DELETE FROM Сотрудник WHERE [Табельный номер сотрудника] = " + tab;
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void ЗагрузитьТоварToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();

            this.Column1.HeaderText = "Номер магазина";
            this.Column2.HeaderText = "Номер отдела";
            this.Column3.HeaderText = "Номер поставщика";
            this.Column4.HeaderText = "Идентификатор товара";
            this.Column5.HeaderText = "Цена";
            this.Column6.HeaderText = "Количество";
            this.Column7.HeaderText = "Срок годности";
            this.Column8.HeaderText = "Дата поставки";
            this.Column9.HeaderText = "";
            this.Column10.HeaderText = "";

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "SELECT * FROM Товар";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            OleDbDataReader dbReader = dbCommand.ExecuteReader();

            if (dbReader.HasRows == false)
            {
                MessageBox.Show("Данные не найдены");
            }
            else
            {

                while (dbReader.Read())
                {


                    dataGridView1.Rows.Add(dbReader["Номер магазина"], dbReader["Номер отдела"], dbReader["Номер поставщика"], dbReader["Идентификатор товара"], dbReader["Цена"], dbReader["Количество"], dbReader["Срок годности"], dbReader["Дата поставки"]);

                }
            }
            dbReader.Close();
            dbConnection.Close();

        }

        private void ДобавитьТоварToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;
            if (dataGridView1.Rows[index].Cells[0].Value == null ||
            dataGridView1.Rows[index].Cells[1].Value == null ||
            dataGridView1.Rows[index].Cells[2].Value == null ||
            dataGridView1.Rows[index].Cells[3].Value == null ||
            dataGridView1.Rows[index].Cells[4].Value == null ||
            dataGridView1.Rows[index].Cells[5].Value == null ||
            dataGridView1.Rows[index].Cells[6].Value == null ||
            dataGridView1.Rows[index].Cells[7].Value == null
            )
            {
                MessageBox.Show("Введены не все данные!");
                return;
            }

            string nomshop = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string nomotd = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string nompost = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string id = dataGridView1.Rows[index].Cells[3].Value.ToString();
            string cost = dataGridView1.Rows[index].Cells[4].Value.ToString();
            string quality = dataGridView1.Rows[index].Cells[5].Value.ToString();
            string srok = dataGridView1.Rows[index].Cells[6].Value.ToString();
            string data = dataGridView1.Rows[index].Cells[7].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "INSERT INTO Товар VALUES ('" + id + "', " + nompost + ", " + nomshop + ", " + nomotd + " , " + cost + ", " + quality + ", '" + srok + "', '" + data + "')";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);
            


            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();


        }

        private void ОбновитьТоварToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;
            

            string nomshop = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string nomotd = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string nompost = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string id = dataGridView1.Rows[index].Cells[3].Value.ToString();
            string cost = dataGridView1.Rows[index].Cells[4].Value.ToString();
            string quality = dataGridView1.Rows[index].Cells[5].Value.ToString();
            string srok = dataGridView1.Rows[index].Cells[6].Value.ToString();
            string data = dataGridView1.Rows[index].Cells[7].Value.ToString();

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "UPDATE [Товар] SET [Номер поставщика]= " + nompost + ",[Номер магазина]= " + nomshop + ",[Номер отдела]= " + nomotd + ",[Цена]= " + cost + ",[Количество]= " + quality + ",[Срок годности]= '" + srok + "',[Дата поставки]= '" + data + "' WHERE [Идентификатор товара] ='" + id + "'";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void УдалитьТоварToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!");
                return;
            }

            int index = dataGridView1.SelectedRows[0].Index;

            string id = dataGridView1.Rows[index].Cells[3].Value.ToString();


            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=lab3.accdb";
            OleDbConnection dbConnection = new OleDbConnection(connectionString);


            dbConnection.Open();
            string query = "DELETE FROM Товар WHERE [Идентификатор товара] ='" + id + "'";
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection);



            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса");
            else
                MessageBox.Show("Данные успешно изменены");

            dbConnection.Close();
        }

        private void ВыходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ПостащикToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Request request = new Request();
            request.Show();
        }
    }
}
