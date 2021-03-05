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
using System.IO;

namespace GIBDD
{
    public partial class Authorization : Form
    {
        private string dbFileName = "GIBDD_DB.db";
        private SQLiteConnection connection;
        private SQLiteCommand command;
        private SQLiteDataReader reader;
        public Authorization()
        {
            InitializeComponent();
            string conString = string.Format("Data Source={0};Version=3", dbFileName);
            connection = new SQLiteConnection(conString);
        }
        private void Authorization_Load(object sender, EventArgs e)
        {
            if (!File.Exists(dbFileName))
            {
                try
                {
                    connection.Open();

                    command = new SQLiteCommand();
                    command.Connection = connection;
                    command.CommandText = "CREATE TABLE IF NOT EXISTS Users(" +
                        "UserID INTEGER PRIMARY KEY AUTOINCREMENT," +
                        "UserName VARCHAR," +
                        "UserPassword VARCHAR);";
                    command.ExecuteNonQuery();

                    command = new SQLiteCommand();
                    command.Connection = connection;
                    command.CommandText = "CREATE TABLE IF NOT EXISTS AutoType(" +
                        "AutoTypeID INTEGER PRIMARY KEY AUTOINCREMENT," +
                        "AutoTypeName VARCHAR," +
                        "TechOsmotrRatio INTEGER," +
                        "СharacteristicsList VARCHAR);";
                    command.ExecuteNonQuery();

                    command = new SQLiteCommand();
                    command.Connection = connection;
                    command.CommandText = "CREATE TABLE IF NOT EXISTS CarCaseType(" +
                        "CarCaseTypeID INTEGER PRIMARY KEY AUTOINCREMENT," +
                        "CarCaseTypeName VARCHAR);";
                    command.ExecuteNonQuery();

                    command = new SQLiteCommand();
                    command.Connection = connection;
                    command.CommandText = "CREATE TABLE IF NOT EXISTS ChassisType(" +
                        "ChassisTypeID INTEGER PRIMARY KEY AUTOINCREMENT," +
                        "ChassisTypeName VARCHAR);";
                    command.ExecuteNonQuery();

                    command = new SQLiteCommand();
                    command.Connection = connection;
                    command.CommandText = "CREATE TABLE IF NOT EXISTS ChaosType(" +
                        "ChaosTypeID INTEGER PRIMARY KEY AUTOINCREMENT," +
                        "ChaosTypeName VARCHAR);";
                    command.ExecuteNonQuery();

                    command = new SQLiteCommand();
                    command.Connection = connection;
                    command.CommandText = "CREATE TABLE IF NOT EXISTS CarNumbersDirectory(" +
                        "CarNumberID INTEGER PRIMARY KEY AUTOINCREMENT," +
                        "CarNumber VARCHAR," +
                        "MasterLastName VARCHAR," +
                        "MasterFirstName VARCHAR," +
                        "MasterMiddleName VARCHAR," +
                        "MasterAddress VARCHAR," +
                        "DopOrganization VARCHAR," +
                        "DopOrganizationAddress VARCHAR," +
                        "DopOrganizationBoss VARCHAR);";
                    command.ExecuteNonQuery();

                    command = new SQLiteCommand();
                    command.Connection = connection;
                    command.CommandText = "CREATE TABLE IF NOT EXISTS Cars(" +
                        "CarID INTEGER PRIMARY KEY AUTOINCREMENT," +
                        "CarNumberID INTEGER REFERENCES CarNumbersDirectory(CarNumberID)," +
                        "AutoModel VARCHAR," +
                        "EngineNum VARCHAR," +
                        "EngineVolume VARCHAR," +
                        "ReleaseDate VARCHAR," +
                        "AutoTypeID INTEGER REFERENCES AutoType(AutoTypeID)," +
                        "CarCaseTypeID INTEGER REFERENCES CarCaseType(CarCaseTypeID)," +
                        "ChassisTypeID INTEGER REFERENCES ChassisType(ChassisTypeID));";
                    command.ExecuteNonQuery();

                    command = new SQLiteCommand();
                    command.Connection = connection;
                    command.CommandText = "CREATE TABLE IF NOT EXISTS SearchingAuto(" +
                        "SearchingAutoID INTEGER PRIMARY KEY AUTOINCREMENT," +
                        "CarID INTEGER REFERENCES Cars(CarID)," +
                        "SearchingInfo VARCHAR," +
                        "SearchingStartDate VARCHAR," +
                        "SearchingStatus INTEGER);";
                    command.ExecuteNonQuery();

                    command = new SQLiteCommand();
                    command.Connection = connection;
                    command.CommandText = "CREATE TABLE IF NOT EXISTS TechOsmotr(" +
                        "TechOsmotrID INTEGER PRIMARY KEY AUTOINCREMENT," +
                        "CarID INTEGER REFERENCES Cars(CarID)," +
                        "QuitanceNum VARCHAR," +
                        "SumToPay DOUBLE);";
                    command.ExecuteNonQuery();

                    command = new SQLiteCommand();
                    command.Connection = connection;
                    command.CommandText = "CREATE TABLE IF NOT EXISTS DTP(" +
                        "DTPID INTEGER PRIMARY KEY AUTOINCREMENT," +
                        "ChaosTypeID INTEGER REFERENCES ChaosType(ChaosTypeID)," +
                        "SumOfDestruction DOUBLE," +
                        "NumOfVictims INTEGER," +
                        "PlaceOfDTP VARCHAR," +
                        "CarID INTEGER REFERENCES Cars(CarID)," +
                        "CauseOfDTP VARCHAR," +
                        "ShortExplanation VARCHAR," +
                        "DateOfDTP VARCHAR," +
                        "RoadConditions VARCHAR);";
                    command.ExecuteNonQuery();
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

        private void button3_Click(object sender, EventArgs e)
        {
            Regestration reg = new Regestration();
            reg.Show();
            Hide();
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            command = new SQLiteCommand();

            try
            {
                connection.Open();
                command.Connection = connection;

                command.CommandText = $"SELECT UserID FROM Users WHERE UserName=@UserName AND UserPassword=@UserPassword";
                command.Parameters.Add("@UserName", System.Data.DbType.String).Value = loginBox.Text;
                command.Parameters.Add("@UserPassword", System.Data.DbType.String).Value = PasswordBox.Text;
                reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        MainForm main = new MainForm();
                        main.Show();
                        Hide();
                    }
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