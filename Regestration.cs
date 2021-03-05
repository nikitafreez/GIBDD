using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using System.Text.RegularExpressions;

namespace GIBDD
{
    public partial class Regestration : Form
    {
        private string dbFileName = "GIBDD_DB.db";
        private SQLiteConnection connection;
        private SQLiteCommand command;
        public Regestration()
        {
            InitializeComponent();
            string conString = string.Format("Data Source={0};Version=3", dbFileName);
            connection = new SQLiteConnection(conString);
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            command = new SQLiteCommand();

            try
            {
                connection.Open();
                command.Connection = connection;
                string login = loginBox.Text;
                string password = passwordBox.Text;
                Regex emailRegex = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)$");
                Regex passwordRegex = new Regex(@"^([A-z0-9._!@%#№?$&*]){5,50}$");
                Match emailMatch = emailRegex.Match(login);
                Match passwordMatch = passwordRegex.Match(password);
                if (emailMatch.Success && passwordMatch.Success)
                {
                    command.CommandText = $"INSERT INTO Users(UserName, UserPassword) VALUES (@UserName, @UserPassword)";
                    command.Parameters.Add("@UserName", System.Data.DbType.String).Value = loginBox.Text;
                    command.Parameters.Add("@UserPassword", System.Data.DbType.String).Value = passwordBox.Text;

                    command.ExecuteNonQuery();

                    Authorization auth = new Authorization();
                    auth.Show();
                    Hide();
                }
                else
                {
                    MessageBox.Show("Некорректно заполнены данные. Повторите попытку");
                }
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
    }
}