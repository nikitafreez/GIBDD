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
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GIBDD
{
    public partial class MainForm : Form
    {
        private string dbFileName = "GIBDD_DB.db";
        private SQLiteConnection connection;
        private SQLiteCommand command;
        private SQLiteDataReader reader;
        private SQLiteDataAdapter dataAdapter;
        private DataSet dataSet;
        public MainForm()
        {
            InitializeComponent();
            label1.Text = "Работа с таблицей ДТП";
            label12.Text = "Поиск по месту";
            panel1.Visible = true;
            panel10.Visible = true;
            DateOfDTPPicker.Visible = true;
            ExcelButton.Visible = true;
            StartSearchingDate.MaxDate = DateTime.Now;
            StartSearchingDate.MinDate = new DateTime(DateTime.Now.Year - 5, DateTime.Now.Month, DateTime.Now.Day);
            ReleaseDatePicker.MaxDate = DateTime.Now;
            ReleaseDatePicker.MinDate = new DateTime(DateTime.Now.Year - 60, DateTime.Now.Month, DateTime.Now.Day);
            string conString = string.Format("Data Source={0};Version=3", dbFileName);
            connection = new SQLiteConnection(conString);
            GetTables();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                label1.Text = "Работа с таблицей ДТП";
                label12.Text = "Поиск по месту";
                panel10.Visible = true;
                label10.Visible = true;
                DateOfDTPPicker.Visible = true;
                panel1.Visible = true;
                panel2.Visible = false;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = false;
                panel6.Visible = false;
                panel7.Visible = false;
                panel8.Visible = false;
                panel9.Visible = false;
                ExcelButton.Visible = true;
            }
            if (tabControl1.SelectedIndex == 1)
            {
                label1.Text = "Работа с таблицей автомобиль";
                label12.Text = "Поиск по марке";
                panel10.Visible = true;
                label10.Visible = false;
                DateOfDTPPicker.Visible = false;
                panel1.Visible = false;
                panel2.Visible = false;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = false;
                panel6.Visible = false;
                panel7.Visible = false;
                panel8.Visible = false;
                panel9.Visible = true;
                ExcelButton.Visible = false;
            }

            else if (tabControl1.SelectedIndex == 2)
            {
                label1.Text = "Работа с таблицей розыск авто";
                panel10.Visible = false;
                label10.Visible = false;
                DateOfDTPPicker.Visible = false;
                panel1.Visible = false;
                panel2.Visible = true;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = false;
                panel6.Visible = false;
                panel7.Visible = false;
                panel8.Visible = false;
                panel9.Visible = false;
                ExcelButton.Visible = true;
            }
            else if (tabControl1.SelectedIndex == 3)
            {
                label1.Text = "Работа с таблицей технический осмотр";
                panel10.Visible = false;
                label10.Visible = false;
                DateOfDTPPicker.Visible = false;
                panel1.Visible = false;
                panel2.Visible = false;
                panel3.Visible = true;
                panel4.Visible = false;
                panel5.Visible = false;
                panel6.Visible = false;
                panel7.Visible = false;
                panel8.Visible = false;
                panel9.Visible = false;
                ExcelButton.Visible = false;
            }
            else if (tabControl1.SelectedIndex == 4)
            {
                label1.Text = "Работа с таблицей справочник номеров";
                label12.Text = "Поиск по номеру";
                panel10.Visible = true;
                label10.Visible = false;
                DateOfDTPPicker.Visible = false;
                panel1.Visible = false;
                panel2.Visible = false;
                panel3.Visible = false;
                panel4.Visible = true;
                panel5.Visible = false;
                panel6.Visible = false;
                panel7.Visible = false;
                panel8.Visible = false;
                panel9.Visible = false;
                ExcelButton.Visible = false;
            }
            else if (tabControl1.SelectedIndex == 5)
            {
                label1.Text = "Работа с таблицей тип происшествия";
                panel10.Visible = false;
                label10.Visible = false;
                DateOfDTPPicker.Visible = false;
                panel1.Visible = false;
                panel2.Visible = false;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = true;
                panel6.Visible = false;
                panel7.Visible = false;
                panel8.Visible = false;
                panel9.Visible = false;
                ExcelButton.Visible = false;
            }
            else if (tabControl1.SelectedIndex == 6)
            {
                label1.Text = "Работа с таблицей тип кузова";
                panel10.Visible = false;
                label10.Visible = false;
                DateOfDTPPicker.Visible = false;
                panel1.Visible = false;
                panel2.Visible = false;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = false;
                panel6.Visible = true;
                panel7.Visible = false;
                panel8.Visible = false;
                panel9.Visible = false;
                ExcelButton.Visible = false;
            }
            else if (tabControl1.SelectedIndex == 7)
            {
                label1.Text = "Работа с таблицей тип шасси";
                panel10.Visible = false;
                label10.Visible = false;
                DateOfDTPPicker.Visible = false;
                panel1.Visible = false;
                panel2.Visible = false;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = false;
                panel6.Visible = false;
                panel7.Visible = true;
                panel8.Visible = false;
                panel9.Visible = false;
                ExcelButton.Visible = false;
            }
            else if (tabControl1.SelectedIndex == 8)
            {
                label1.Text = "Работа с таблицей тип автомобиля";
                panel10.Visible = false;
                label10.Visible = false;
                DateOfDTPPicker.Visible = false;
                panel1.Visible = false;
                panel2.Visible = false;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = false;
                panel6.Visible = false;
                panel7.Visible = false;
                panel8.Visible = true;
                panel9.Visible = false;
                ExcelButton.Visible = false;
            }

        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        DataTable dt = new DataTable();
        private void GetTables()
        {
            try
            {
                connection.Open();
                dataAdapter = new SQLiteDataAdapter("SELECT DTPID, ChaosType.ChaosTypeName, SumOfDestruction, NumOfVictims, PlaceOfDTP, CarNumbersDirectory.CarNumber, CauseOfDTP, ShortExplanation, DateOfDTP FROM DTP " +
                    "INNER JOIN CarNumbersDirectory ON DTP.CarID = CarNumbersDirectory.CarNumberID " +
                    "INNER JOIN ChaosType ON DTP.ChaosTypeID = ChaosType.ChaosTypeID", connection);
                dataSet = new DataSet();
                dt = dataSet.Tables.Add("DTP");
                dataAdapter.Fill(dt);
                DTP_dataGridView.DataSource = dataSet.Tables["DTP"];
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }

            try
            {
                connection.Open();
                dataAdapter = new SQLiteDataAdapter("SELECT CarID, CarNumbersDirectory.CarNumber, AutoModel, EngineNum, EngineVolume, " +
                    "ReleaseDate, AutoType.AutoTypeName, CarCaseType.CarCaseTypeName, ChassisType.ChassisTypeName FROM Cars " +
                    "INNER JOIN CarNumbersDirectory ON Cars.CarNumberID = CarNumbersDirectory.CarNumberID " +
                    "INNER JOIN AutoType ON Cars.AutoTypeID = AutoType.AutoTypeID " +
                    "INNER JOIN CarCaseType ON Cars.CarCaseTypeID = CarCaseType.CarCaseTypeID " +
                    "INNER JOIN ChassisType ON Cars.ChassisTypeID = ChassisType.ChassisTypeID", connection);
                dataSet = new DataSet();
                dt = dataSet.Tables.Add("Cars");
                dataAdapter.Fill(dt);
                Auto_dataGridView.DataSource = dataSet.Tables["Cars"];
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }

            try
            {
                connection.Open();
                dataAdapter = new SQLiteDataAdapter("SELECT SearchingAutoID, CarNumbersDirectory.CarNumber, SearchingInfo, SearchingStartDate, SearchingStatus FROM SearchingAuto INNER JOIN CarNumbersDirectory ON SearchingAuto.CarID = CarNumbersDirectory.CarNumberID", connection);
                dataSet = new DataSet();
                dt = dataSet.Tables.Add("SearchingAuto");
                dataAdapter.Fill(dt);
                Searching_dataGridView.DataSource = dataSet.Tables["SearchingAuto"];
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }

            try
            {
                connection.Open();
                dataAdapter = new SQLiteDataAdapter("SELECT * FROM CarNumbersDirectory", connection);
                dataSet = new DataSet();
                dt = dataSet.Tables.Add("CarNumbersDirectory");
                dataAdapter.Fill(dt);
                NumbersList_dataGridView.DataSource = dataSet.Tables["CarNumbersDirectory"];

                NumOfOsmotrCarComboBox.DataSource = dataSet.Tables["CarNumbersDirectory"];
                NumOfOsmotrCarComboBox.DisplayMember = "CarNumber";
                NumOfOsmotrCarComboBox.ValueMember = "CarNumberID";

                CarNumSearchingComboBox.DataSource = dataSet.Tables["CarNumbersDirectory"];
                CarNumSearchingComboBox.DisplayMember = "CarNumber";
                CarNumSearchingComboBox.ValueMember = "CarNumberID";

                NumOfCarComboBox.DataSource = dataSet.Tables["CarNumbersDirectory"];
                NumOfCarComboBox.DisplayMember = "CarNumber";
                NumOfCarComboBox.ValueMember = "CarNumberID";

                NumOfHurtCarComboBox.DataSource = dataSet.Tables["CarNumbersDirectory"];
                NumOfHurtCarComboBox.DisplayMember = "CarNumber";
                NumOfHurtCarComboBox.ValueMember = "CarNumberID";

            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }


            try
            {
                connection.Open();
                dataAdapter = new SQLiteDataAdapter("SELECT TechOsmotrID, CarNumbersDirectory.CarNumber, QuitanceNum, SumToPay FROM TechOsmotr INNER JOIN CarNumbersDirectory ON TechOsmotr.CarID = CarNumbersDirectory.CarNumberID", connection);
                dataSet = new DataSet();
                dt = dataSet.Tables.Add("TechOsmotr");
                dataAdapter.Fill(dt);
                Osmotr_dataGridView.DataSource = dataSet.Tables["TechOsmotr"];
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }

            try
            {
                connection.Open();
                dataAdapter = new SQLiteDataAdapter("SELECT * FROM ChaosType", connection);
                dataSet = new DataSet();
                dt = dataSet.Tables.Add("ChaosType");
                dataAdapter.Fill(dt);
                ChaosType_dataGridView.DataSource = dataSet.Tables["ChaosType"];

                ChaosTypeComboBox.DataSource = dataSet.Tables["ChaosType"];
                ChaosTypeComboBox.DisplayMember = "ChaosTypeName";
                ChaosTypeComboBox.ValueMember = "ChaosTypeID";
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }

            try
            {
                connection.Open();
                dataAdapter = new SQLiteDataAdapter("SELECT * FROM CarCaseType", connection);
                dataSet = new DataSet();
                dt = dataSet.Tables.Add("CarCaseType");
                dataAdapter.Fill(dt);
                CarCase_dataGridView.DataSource = dataSet.Tables["CarCaseType"];

                CarCaseTypeComboBox.DataSource = dataSet.Tables["CarCaseType"];
                CarCaseTypeComboBox.DisplayMember = "CarCaseTypeName";
                CarCaseTypeComboBox.ValueMember = "CarCaseTypeID";
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }

            try
            {
                connection.Open();
                dataAdapter = new SQLiteDataAdapter("SELECT * FROM ChassisType", connection);
                dataSet = new DataSet();
                dt = dataSet.Tables.Add("ChassisType");
                dataAdapter.Fill(dt);
                Chassis_dataGridView.DataSource = dataSet.Tables["ChassisType"];

                ChassiesTypeComboBox.DataSource = dataSet.Tables["ChassisType"];
                ChassiesTypeComboBox.DisplayMember = "ChassisTypeName";
                ChassiesTypeComboBox.ValueMember = "ChassisTypeID";
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }

            try
            {
                connection.Open();
                dataAdapter = new SQLiteDataAdapter("SELECT * FROM AutoType", connection);
                dataSet = new DataSet();
                dt = dataSet.Tables.Add("AutoType");
                dataAdapter.Fill(dt);
                AutoType_dataGridView.DataSource = dataSet.Tables["AutoType"];

                AutoTypeComboBox.DataSource = dataSet.Tables["AutoType"];
                AutoTypeComboBox.DisplayMember = "AutoTypeName";
                AutoTypeComboBox.ValueMember = "AutoTypeID";
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

        private void AddButton_Click(object sender, EventArgs e)
        {
            int selectedIndex = tabControl1.SelectedIndex;

            switch (selectedIndex)
            {
                case 0:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            int chaosType = Convert.ToInt32(ChaosTypeComboBox.SelectedValue);
                            string sumOfDestruction = SumOfDestructionBox.Text;
                            string peopleInDTP = HurtPeopleSumBox.Text;
                            string placeOfDTP = PlaceOfDTPBox.Text;
                            int numOfCar = Convert.ToInt32(NumOfCarComboBox.SelectedValue);
                            string whyDTP = WhyDTPBox.Text;
                            string shortExplanation = ShortExplanationBox.Text;
                            string roadCondition = RoadCondictionsBox.Text;
                            string dateOfDTP = DateOfDTPPicker.Value.ToString();

                            Regex sumOfDestructionRegex = new Regex(@"^[0-9]{2,8}$");
                            Regex peopleInDTPRegex = new Regex(@"^[0-9]{1,3}$");
                            Regex placeOfDTPRegex = new Regex(@"^[А-яA-zеЁ0-9(\s),.]{3,200}$");
                            Regex whyDTPRegex = new Regex(@"^[А-яA-zеЁ0-9(\s),.]{3,200}$");
                            Regex shortExplanationRegex = new Regex(@"^[А-яA-zеЁ0-9(\s),.]{3,200}$");
                            Regex roadConditionRegex = new Regex(@"^[А-яA-zеЁ0-9(\s),.]{3,200}$");

                            Match sumOfDestructioMatch = sumOfDestructionRegex.Match(sumOfDestruction);
                            Match peopleInDTPMatch = peopleInDTPRegex.Match(peopleInDTP);
                            Match placeOfDTPMatch = placeOfDTPRegex.Match(placeOfDTP);
                            Match whyDTPMatch = whyDTPRegex.Match(whyDTP);
                            Match shortExplanationMatch = shortExplanationRegex.Match(shortExplanation);
                            Match roadConditionMatch = roadConditionRegex.Match(roadCondition);

                            if (sumOfDestructioMatch.Success && peopleInDTPMatch.Success && placeOfDTPMatch.Success && whyDTPMatch.Success && shortExplanationMatch.Success && roadConditionMatch.Success)
                            {
                                command.CommandText = $"INSERT INTO DTP(ChaosTypeID, SumOfDestruction, NumOfVictims, PlaceOfDTP, CarID, CauseOfDTP, ShortExplanation, DateOfDTP) " +
                                    $"VALUES (@ChaosTypeID, @SumOfDestruction, @NumOfVictims, @PlaceOfDTP, @CarID, @CauseOfDTP, @ShortExplanation, @DateOfDTP)";
                                command.Parameters.Add("@ChaosTypeID", System.Data.DbType.String).Value = chaosType.ToString();
                                command.Parameters.Add("@SumOfDestruction", System.Data.DbType.String).Value = sumOfDestruction;
                                command.Parameters.Add("@NumOfVictims", System.Data.DbType.String).Value = peopleInDTP;
                                command.Parameters.Add("@PlaceOfDTP", System.Data.DbType.String).Value = placeOfDTP;
                                command.Parameters.Add("@CarID", System.Data.DbType.String).Value = numOfCar.ToString();
                                command.Parameters.Add("@CauseOfDTP", System.Data.DbType.String).Value = whyDTP;
                                command.Parameters.Add("@ShortExplanation", System.Data.DbType.String).Value = shortExplanation;
                                command.Parameters.Add("@DateOfDTP", System.Data.DbType.String).Value = dateOfDTP;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 1:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;

                            int numOfCar = Convert.ToInt32(NumOfCarComboBox.SelectedValue);
                            string markOfAuto = MarkOfAutoComboBox.Text;
                            string engineNum = EngineNumBox.Text;
                            string engimeCapacity = EngineCapacityBox.Text;
                            string releaseDate = ReleaseDatePicker.Value.ToString();
                            int autoType = Convert.ToInt32(AutoTypeComboBox.SelectedValue);
                            int carCaseType = Convert.ToInt32(CarCaseTypeComboBox.SelectedValue);
                            int chassiesType = Convert.ToInt32(ChassiesTypeComboBox.SelectedValue);

                            Regex markOfAutoRegex = new Regex(@"^[A-zА-я(\s)-]{1,20}$");
                            Regex engineNumRegex = new Regex(@"^[A-Z0-9]{17,17}$");

                            Match markOfAutoMatch = markOfAutoRegex.Match(markOfAuto);
                            Match engineNumMatch = engineNumRegex.Match(engineNum);

                            if (markOfAutoMatch.Success && engineNumMatch.Success)
                            {
                                command.CommandText = $"INSERT INTO Cars( CarNumberID, AutoModel, EngineNum, EngineVolume, ReleaseDate, AutoTypeID, CarCaseTypeID, ChassisTypeID) " +
                                    $"VALUES (@CarNumberID, @AutoModel, @EngineNum, @EngineVolume, @ReleaseDate, @AutoTypeID, @CarCaseTypeID, @ChassisTypeID)";
                                command.Parameters.Add("@CarNumberID", System.Data.DbType.String).Value = numOfCar.ToString();
                                command.Parameters.Add("@AutoModel", System.Data.DbType.String).Value = markOfAuto;
                                command.Parameters.Add("@EngineNum", System.Data.DbType.String).Value = engineNum;
                                command.Parameters.Add("@EngineVolume", System.Data.DbType.String).Value = engimeCapacity;
                                command.Parameters.Add("@ReleaseDate", System.Data.DbType.String).Value = releaseDate;
                                command.Parameters.Add("@AutoTypeID", System.Data.DbType.String).Value = autoType.ToString();
                                command.Parameters.Add("@CarCaseTypeID", System.Data.DbType.String).Value = carCaseType.ToString();
                                command.Parameters.Add("@ChassisTypeID", System.Data.DbType.String).Value = chassiesType.ToString();

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 2:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;

                            string searchingInfo = SearchingInfoBox.Text;
                            int numOfCar = Convert.ToInt32(NumOfOsmotrCarComboBox.SelectedValue);
                            string dateOfStartSearching = StartSearchingDate.Value.ToString();
                            int statusSearching;
                            if (SearchingStatusCheckBox.Checked)
                            {
                                statusSearching = 1;
                            }
                            else
                            {
                                statusSearching = 0;
                            }

                            Regex searchingInfoRegex = new Regex(@"^[A-zА-я0-9(\s).,]{1,200}$");

                            Match searchingInfoMatch = searchingInfoRegex.Match(searchingInfo);

                            if (searchingInfoMatch.Success)
                            {
                                command.CommandText = $"INSERT INTO SearchingAuto(CarID, SearchingInfo, SearchingStartDate, SearchingStatus) VALUES (@CarID, @SearchingInfo, @SearchingStartDate, @SearchingStatus)";
                                command.Parameters.Add("@CarID", System.Data.DbType.String).Value = numOfCar.ToString();
                                command.Parameters.Add("@SearchingInfo", System.Data.DbType.String).Value = searchingInfo;
                                command.Parameters.Add("@SearchingStartDate", System.Data.DbType.String).Value = dateOfStartSearching;
                                command.Parameters.Add("@SearchingStatus", System.Data.DbType.String).Value = statusSearching;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 3:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            int numOfCar = Convert.ToInt32(NumOfOsmotrCarComboBox.SelectedValue);
                            string quitanceNums = QuitanceNumBox.Text;
                            string sumOfOplata = SumOfOplataBox.Text;

                            Regex quitanceNumsRegex = new Regex(@"^[0-9]{10,15}$");
                            Regex sumOfOplataRegex = new Regex(@"^[0-9]{3,6}$");

                            Match quitanceNumsMatch = quitanceNumsRegex.Match(quitanceNums);
                            Match sumOfOplataMatch = sumOfOplataRegex.Match(sumOfOplata);

                            if (quitanceNumsMatch.Success && sumOfOplataMatch.Success)
                            {
                                command.CommandText = $"INSERT INTO TechOsmotr(CarID, QuitanceNum, SumToPay) VALUES (@CarID, @QuitanceNum, @SumToPay)";
                                command.Parameters.Add("@CarID", System.Data.DbType.String).Value = numOfCar.ToString();
                                command.Parameters.Add("@QuitanceNum", System.Data.DbType.String).Value = quitanceNums;
                                command.Parameters.Add("@SumToPay", System.Data.DbType.String).Value = sumOfOplata;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 4:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            string carNumber = CarNumberBox.Text;
                            string firstName = FirstNameBox.Text;
                            string lastName = LastNameBox.Text;
                            string middleName = MiddleNameBox.Text;
                            string address = AddressBox.Text;

                            string organizationName = OrganizationBox.Text;
                            string FIOMaster = OrganizationMasterFIOBox.Text;
                            string addressOrganization = OrganizationAddressBox.Text;

                            Regex carNumberRegex = new Regex(@"^[АВЕКМНОРСТУХ]{1}[0-9]{3}[АВЕКМНОРСТУХ]{2}$");
                            Regex firstNameRegex = new Regex(@"^[А-я]{1,30}");
                            Regex lastNameRegex = new Regex(@"^[А-я]{1,30}");
                            Regex middleNameRegex = new Regex(@"^[А-яA-z]{0,30}");
                            Regex addressRegex = new Regex(@"^[А-яA-zеЁ0-9(\s),.]{3,200}$");

                            Regex dopFIORegex = new Regex(@"^[А-я\s]{0,150}$");
                            Regex dopAddressRegex = new Regex(@"^[А-яA-z0-9,.]{0,200}$");
                            Regex organizationNameRegex = new Regex(@"^[А-яA-z]{0,30}$");

                            Match carNumberMatch = carNumberRegex.Match(carNumber);
                            Match firstNameMatch = firstNameRegex.Match(firstName);
                            Match lastNameMatch = lastNameRegex.Match(lastName);
                            Match middleNameMatch = middleNameRegex.Match(middleName);
                            Match addressMatch = addressRegex.Match(address);

                            Match organizationNameMatch = organizationNameRegex.Match(organizationName);
                            Match FIOMasterMatch = dopFIORegex.Match(FIOMaster);
                            Match addressOrganizationMatch = dopAddressRegex.Match(addressOrganization);

                            if (carNumberMatch.Success && firstNameMatch.Success && lastNameMatch.Success && middleNameMatch.Success && addressMatch.Success && organizationNameMatch.Success && FIOMasterMatch.Success && addressOrganizationMatch.Success)
                            {
                                command.CommandText = $"INSERT INTO CarNumbersDirectory(CarNumber, MasterLastName, MasterFirstName, MasterMiddleName, MasterAddress, DopOrganization, DopOrganizationAddress, DopOrganizationBoss)" +
                                    $" VALUES (@CarNumber, @MasterLastName, @MasterFirstName, @MasterMiddleName, @MasterAddress, @DopOrganization, @DopOrganizationAddress, @DopOrganizationBoss)";
                                command.Parameters.Add("@CarNumber", System.Data.DbType.String).Value = carNumber;
                                command.Parameters.Add("@MasterLastName", System.Data.DbType.String).Value = lastName;
                                command.Parameters.Add("@MasterFirstName", System.Data.DbType.String).Value = firstName;
                                command.Parameters.Add("@MasterMiddleName", System.Data.DbType.String).Value = middleName;
                                command.Parameters.Add("@MasterAddress", System.Data.DbType.String).Value = address;
                                command.Parameters.Add("@DopOrganization", System.Data.DbType.String).Value = organizationName;
                                command.Parameters.Add("@DopOrganizationAddress", System.Data.DbType.String).Value = addressOrganization;
                                command.Parameters.Add("@DopOrganizationBoss", System.Data.DbType.String).Value = FIOMaster;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 5:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            string chaosName = ChaosNameBox.Text;
                            Regex chaosRegex = new Regex(@"^([A-zА-я ]){2,}$");
                            Match chaosMatch = chaosRegex.Match(chaosName);
                            if (chaosMatch.Success)
                            {
                                command.CommandText = $"INSERT INTO ChaosType(ChaosTypeName) VALUES (@ChaosTypeName)";
                                command.Parameters.Add("@ChaosTypeName", System.Data.DbType.String).Value = chaosName;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 6:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            string kuzovTypeName = KuzovTypeBox.Text;
                            Regex kuzovNameRegex = new Regex(@"^([A-zА-я]){3,20}$");
                            Match kuzovNameMatch = kuzovNameRegex.Match(kuzovTypeName);
                            if (kuzovNameMatch.Success)
                            {
                                command.CommandText = $"INSERT INTO CarCaseType(CarCaseTypeName) VALUES (@CarCaseTypeName)";
                                command.Parameters.Add("@CarCaseTypeName", System.Data.DbType.String).Value = kuzovTypeName;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 7:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            string chassisName = ChassisTypeNameBox.Text;
                            Regex chassisNameRegex = new Regex(@"^([A-zА-я]){1,20}$");
                            Match chassisNameMatch = chassisNameRegex.Match(chassisName);
                            if (chassisNameMatch.Success)
                            {
                                command.CommandText = $"INSERT INTO ChassisType(ChassisTypeName) VALUES (@ChassisTypeName)";
                                command.Parameters.Add("@ChassisTypeName", System.Data.DbType.String).Value = chassisName;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 8:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            string autoType = AutoTypeNameBox.Text;
                            string pereodichnost = PereodichnostTechBox.Text;
                            string characteristics = CharacteristicsBox.Text;
                            Regex autoTypeRegex = new Regex(@"^([A-zА-я\s]){1,20}$");
                            Regex pereodichnostRegex = new Regex(@"^([0-9]){1,2}$");
                            Regex characteristicsRegex = new Regex(@"^([А-я,]){1,}$");
                            Match autoTypeMatch = autoTypeRegex.Match(autoType);
                            Match pereodichnostMatch = pereodichnostRegex.Match(pereodichnost.ToString());
                            Match characteristicsMatch = characteristicsRegex.Match(characteristics);
                            if (autoTypeMatch.Success && pereodichnostMatch.Success && characteristicsMatch.Success)
                            {
                                command.CommandText = $"INSERT INTO AutoType(AutoTypeName, TechOsmotrRatio, СharacteristicsList) VALUES (@AutoTypeName, @TechOsmotrRatio, @СharacteristicsList)";
                                command.Parameters.Add("@AutoTypeName", System.Data.DbType.String).Value = autoType;
                                command.Parameters.Add("@TechOsmotrRatio", System.Data.DbType.String).Value = pereodichnost;
                                command.Parameters.Add("@СharacteristicsList", System.Data.DbType.String).Value = characteristics;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
            }
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            int selectedIndex = tabControl1.SelectedIndex;

            switch (selectedIndex)
            {
                case 0:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            command.CommandText = $"DELETE FROM DTP WHERE DTPID = {DTP_dataGridView.SelectedRows[0].Cells[0].Value}";
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
                        GetTables();
                        break;
                    }
                case 1:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            command.CommandText = $"DELETE FROM Cars WHERE CarID = {Auto_dataGridView.SelectedRows[0].Cells[0].Value}";
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
                        GetTables();
                        break;
                    }
                case 2:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            command.CommandText = $"DELETE FROM SearchingAuto WHERE SearchingAutoID = {Searching_dataGridView.SelectedRows[0].Cells[0].Value}";
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
                        GetTables();
                        break;
                    }
                case 3:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            command.CommandText = $"DELETE FROM TechOsmotr WHERE TechOsmotrID = {Osmotr_dataGridView.SelectedRows[0].Cells[0].Value}";
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
                        GetTables();
                        break;
                    }
                case 4:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            command.CommandText = $"DELETE FROM CarNumbersDirectory WHERE CarNumberID = {NumbersList_dataGridView.SelectedRows[0].Cells[0].Value}";
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
                        GetTables();
                        break;
                    }
                case 5:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            command.CommandText = $"DELETE FROM ChaosType WHERE ChaosTypeID = {ChaosType_dataGridView.SelectedRows[0].Cells[0].Value}";
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
                        GetTables();
                        break;
                    }
                case 6:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            command.CommandText = $"DELETE FROM CarCaseType WHERE CarCaseTypeID = {CarCase_dataGridView.SelectedRows[0].Cells[0].Value}";
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
                        GetTables();
                        break;
                    }
                case 7:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            command.CommandText = $"DELETE FROM ChassisType WHERE ChassisTypeID = {Chassis_dataGridView.SelectedRows[0].Cells[0].Value}";
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
                        GetTables();
                        break;
                    }
                case 8:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            command.CommandText = $"DELETE FROM AutoType WHERE AutoTypeID = {AutoType_dataGridView.SelectedRows[0].Cells[0].Value}";
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
                        GetTables();
                        break;
                    }
            }
        }

        private void UpdateButton_Click(object sender, EventArgs e)
        {
            int selectedIndex = tabControl1.SelectedIndex;

            switch (selectedIndex)
            {
                case 0:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            int chaosType = Convert.ToInt32(ChaosTypeComboBox.SelectedValue);
                            string sumOfDestruction = SumOfDestructionBox.Text;
                            string peopleInDTP = HurtPeopleSumBox.Text;
                            string placeOfDTP = PlaceOfDTPBox.Text;
                            int numOfCar = Convert.ToInt32(NumOfCarComboBox.SelectedValue);
                            string whyDTP = WhyDTPBox.Text;
                            string shortExplanation = ShortExplanationBox.Text;
                            string roadCondition = RoadCondictionsBox.Text;
                            string dateOfDTP = DateOfDTPPicker.Value.ToString();

                            Regex sumOfDestructionRegex = new Regex(@"^[0-9]{2,8}$");
                            Regex peopleInDTPRegex = new Regex(@"^[0-9]{1,3}$");
                            Regex placeOfDTPRegex = new Regex(@"^[А-яA-zеЁ0-9(\s),.]{3,200}$");
                            Regex whyDTPRegex = new Regex(@"^[А-яA-zеЁ0-9(\s),.]{3,200}$");
                            Regex shortExplanationRegex = new Regex(@"^[А-яA-zеЁ0-9(\s),.]{3,200}$");
                            Regex roadConditionRegex = new Regex(@"^[А-яA-zеЁ0-9(\s),.]{3,200}$");

                            Match sumOfDestructioMatch = sumOfDestructionRegex.Match(sumOfDestruction);
                            Match peopleInDTPMatch = peopleInDTPRegex.Match(peopleInDTP);
                            Match placeOfDTPMatch = placeOfDTPRegex.Match(placeOfDTP);
                            Match whyDTPMatch = whyDTPRegex.Match(whyDTP);
                            Match shortExplanationMatch = shortExplanationRegex.Match(shortExplanation);
                            Match roadConditionMatch = roadConditionRegex.Match(roadCondition);

                            if (sumOfDestructioMatch.Success && peopleInDTPMatch.Success && placeOfDTPMatch.Success && whyDTPMatch.Success && shortExplanationMatch.Success && roadConditionMatch.Success)
                            {
                                command.CommandText = $"UPDATE DTP SET ChaosTypeID = @ChaosTypeID, SumOfDestruction = @SumOfDestruction, NumOfVictims = @NumOfVictims, PlaceOfDTP = @PlaceOfDTP, CarID = @CarID, CauseOfDTP = @CauseOfDTP, ShortExplanation = @ShortExplanation, DateOfDTP = @DateOfDTP " +
                                    $"WHERE DTPID = {DTP_dataGridView.SelectedRows[0].Cells[0].Value}";
                                command.Parameters.Add("@ChaosTypeID", System.Data.DbType.String).Value = chaosType.ToString();
                                command.Parameters.Add("@SumOfDestruction", System.Data.DbType.String).Value = sumOfDestruction;
                                command.Parameters.Add("@NumOfVictims", System.Data.DbType.String).Value = peopleInDTP;
                                command.Parameters.Add("@PlaceOfDTP", System.Data.DbType.String).Value = placeOfDTP;
                                command.Parameters.Add("@CarID", System.Data.DbType.String).Value = numOfCar.ToString();
                                command.Parameters.Add("@CauseOfDTP", System.Data.DbType.String).Value = whyDTP;
                                command.Parameters.Add("@ShortExplanation", System.Data.DbType.String).Value = shortExplanation;
                                command.Parameters.Add("@DateOfDTP", System.Data.DbType.String).Value = dateOfDTP;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 1:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;

                            int numOfCar = Convert.ToInt32(NumOfCarComboBox.SelectedValue);
                            string markOfAuto = MarkOfAutoComboBox.Text;
                            string engineNum = EngineNumBox.Text;
                            string engimeCapacity = EngineCapacityBox.Text;
                            string releaseDate = ReleaseDatePicker.Value.ToString();
                            int autoType = Convert.ToInt32(AutoTypeComboBox.SelectedValue);
                            int carCaseType = Convert.ToInt32(CarCaseTypeComboBox.SelectedValue);
                            int chassiesType = Convert.ToInt32(ChassiesTypeComboBox.SelectedValue);

                            Regex markOfAutoRegex = new Regex(@"^[A-zА-я(\s)-]{1,20}$");
                            Regex engineNumRegex = new Regex(@"^[A-Z0-9]{17,17}$");

                            Match markOfAutoMatch = markOfAutoRegex.Match(markOfAuto);
                            Match engineNumMatch = engineNumRegex.Match(engineNum);

                            if (markOfAutoMatch.Success && engineNumMatch.Success)
                            {
                                command.CommandText = $"UPDATE Cars SET CarNumberID = @CarNumberID, AutoModel = @AutoModel, EngineNum = @EngineNum, EngineVolume = @EngineVolume, ReleaseDate = @ReleaseDate, AutoTypeID =  @AutoTypeID, CarCaseTypeID = @CarCaseTypeID, ChassisTypeID = @ChassisTypeID" +
                                    $"WHERE CarID = {Auto_dataGridView.SelectedRows[0].Cells[0].Value}";
                                command.Parameters.Add("@CarNumberID", System.Data.DbType.String).Value = numOfCar.ToString();
                                command.Parameters.Add("@AutoModel", System.Data.DbType.String).Value = markOfAuto;
                                command.Parameters.Add("@EngineNum", System.Data.DbType.String).Value = engineNum;
                                command.Parameters.Add("@EngineVolume", System.Data.DbType.String).Value = engimeCapacity;
                                command.Parameters.Add("@ReleaseDate", System.Data.DbType.String).Value = releaseDate;
                                command.Parameters.Add("@AutoTypeID", System.Data.DbType.String).Value = autoType.ToString();
                                command.Parameters.Add("@CarCaseTypeID", System.Data.DbType.String).Value = carCaseType.ToString();
                                command.Parameters.Add("@ChassisTypeID", System.Data.DbType.String).Value = chassiesType.ToString();

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 2:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;

                            string searchingInfo = SearchingInfoBox.Text;
                            int numOfCar = Convert.ToInt32(NumOfOsmotrCarComboBox.SelectedValue);
                            string dateOfStartSearching = StartSearchingDate.Value.ToString();
                            int statusSearching;
                            if (SearchingStatusCheckBox.Checked)
                            {
                                statusSearching = 1;
                            }
                            else
                            {
                                statusSearching = 0;
                            }

                            Regex searchingInfoRegex = new Regex(@"^[A-zА-я0-9(\s).,]{1,200}$");

                            Match searchingInfoMatch = searchingInfoRegex.Match(searchingInfo);

                            if (searchingInfoMatch.Success)
                            {
                                command.CommandText = $"UPDATE SearchingAuto SET CarID = @CarID, SearchingInfo = @SearchingInfo, SearchingStartDate = @SearchingStartDate, SearchingStatus = @SearchingStatus WHERE SearchingAutoID = {Searching_dataGridView.SelectedRows[0].Cells[0].Value}";
                                command.Parameters.Add("@CarID", System.Data.DbType.String).Value = numOfCar.ToString();
                                command.Parameters.Add("@SearchingInfo", System.Data.DbType.String).Value = searchingInfo;
                                command.Parameters.Add("@SearchingStartDate", System.Data.DbType.String).Value = dateOfStartSearching;
                                command.Parameters.Add("@SearchingStatus", System.Data.DbType.String).Value = statusSearching;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 3:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            int numOfCar = Convert.ToInt32(NumOfOsmotrCarComboBox.SelectedValue);
                            string quitanceNums = QuitanceNumBox.Text;
                            string sumOfOplata = SumOfOplataBox.Text;

                            Regex quitanceNumsRegex = new Regex(@"^[0-9]{10,15}$");
                            Regex sumOfOplataRegex = new Regex(@"^[0-9]{3,6}$");

                            Match quitanceNumsMatch = quitanceNumsRegex.Match(quitanceNums);
                            Match sumOfOplataMatch = sumOfOplataRegex.Match(sumOfOplata);

                            if (quitanceNumsMatch.Success && sumOfOplataMatch.Success)
                            {
                                command.CommandText = $"UPDATE TechOsmotr SET CarID = @CarID, QuitanceNum = @QuitanceNum, SumToPay = @SumToPay WHERE TechOsmotrID = {Osmotr_dataGridView.SelectedRows[0].Cells[0].Value}";
                                command.Parameters.Add("@CarID", System.Data.DbType.String).Value = numOfCar.ToString();
                                command.Parameters.Add("@QuitanceNum", System.Data.DbType.String).Value = quitanceNums;
                                command.Parameters.Add("@SumToPay", System.Data.DbType.String).Value = sumOfOplata;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 4:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            string carNumber = CarNumberBox.Text;
                            string firstName = FirstNameBox.Text;
                            string lastName = LastNameBox.Text;
                            string middleName = MiddleNameBox.Text;
                            string address = AddressBox.Text;

                            string organizationName = OrganizationBox.Text;
                            string FIOMaster = OrganizationMasterFIOBox.Text;
                            string addressOrganization = OrganizationAddressBox.Text;

                            Regex carNumberRegex = new Regex(@"^[АВЕКМНОРСТУХ]{1}[0-9]{3}[АВЕКМНОРСТУХ]{2}$");
                            Regex firstNameRegex = new Regex(@"^[А-я]{1,30}");
                            Regex lastNameRegex = new Regex(@"^[А-я]{1,30}");
                            Regex middleNameRegex = new Regex(@"^[А-яA-z]{0,30}");
                            Regex addressRegex = new Regex(@"^[А-яA-zеЁ0-9(\s),.]{3,200}$");

                            Regex dopFIORegex = new Regex(@"^[А-я\s]{0,150}$");
                            Regex dopAddressRegex = new Regex(@"^[А-яA-z0-9,.]{0,200}$");
                            Regex organizationNameRegex = new Regex(@"^[А-яA-z]{0,30}$");

                            Match carNumberMatch = carNumberRegex.Match(carNumber);
                            Match firstNameMatch = firstNameRegex.Match(firstName);
                            Match lastNameMatch = lastNameRegex.Match(lastName);
                            Match middleNameMatch = middleNameRegex.Match(middleName);
                            Match addressMatch = addressRegex.Match(address);

                            Match organizationNameMatch = organizationNameRegex.Match(organizationName);
                            Match FIOMasterMatch = dopFIORegex.Match(FIOMaster);
                            Match addressOrganizationMatch = dopAddressRegex.Match(addressOrganization);

                            if (carNumberMatch.Success && firstNameMatch.Success && lastNameMatch.Success && middleNameMatch.Success && addressMatch.Success && organizationNameMatch.Success && FIOMasterMatch.Success && addressOrganizationMatch.Success)
                            {
                                command.CommandText = $"UPDATE CarNumbersDirectory SET CarNumber = @CarNumber, MasterLastName = @MasterLastName, MasterFirstName = @MasterFirstName, MasterMiddleName = @MasterMiddleName, MasterAddress = @MasterAddress, DopOrganization = @DopOrganization, DopOrganizationAddress = @DopOrganizationAddress, DopOrganizationBoss = @DopOrganizationBoss" +
                                    $" WHERE CarNumberID = {NumbersList_dataGridView.SelectedRows[0].Cells[0].Value}";
                                command.Parameters.Add("@CarNumber", System.Data.DbType.String).Value = carNumber;
                                command.Parameters.Add("@MasterLastName", System.Data.DbType.String).Value = lastName;
                                command.Parameters.Add("@MasterFirstName", System.Data.DbType.String).Value = firstName;
                                command.Parameters.Add("@MasterMiddleName", System.Data.DbType.String).Value = middleName;
                                command.Parameters.Add("@MasterAddress", System.Data.DbType.String).Value = address;
                                command.Parameters.Add("@DopOrganization", System.Data.DbType.String).Value = organizationName;
                                command.Parameters.Add("@DopOrganizationAddress", System.Data.DbType.String).Value = addressOrganization;
                                command.Parameters.Add("@DopOrganizationBoss", System.Data.DbType.String).Value = FIOMaster;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 5:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            string chaosName = ChaosNameBox.Text;
                            Regex chaosRegex = new Regex(@"^([A-zА-я ]){2,}$");
                            Match chaosMatch = chaosRegex.Match(chaosName);
                            if (chaosMatch.Success)
                            {
                                command.CommandText = $"UPDATE ChaosType SET ChaosTypeName = @ChaosTypeName WHERE ChaosTypeID = {ChaosType_dataGridView.SelectedRows[0].Cells[0].Value}";
                                command.Parameters.Add("@ChaosTypeName", System.Data.DbType.String).Value = chaosName;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 6:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            string kuzovTypeName = KuzovTypeBox.Text;
                            Regex kuzovNameRegex = new Regex(@"^([A-zА-я]){3,20}$");
                            Match kuzovNameMatch = kuzovNameRegex.Match(kuzovTypeName);
                            if (kuzovNameMatch.Success)
                            {
                                command.CommandText = $"UPDATE CarCaseType SET CarCaseTypeName = @CarCaseTypeName WHERE CarCaseTypeID = {CarCase_dataGridView.SelectedRows[0].Cells[0].Value}";
                                command.Parameters.Add("@CarCaseTypeName", System.Data.DbType.String).Value = kuzovTypeName;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 7:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            string chassisName = ChassisTypeNameBox.Text;
                            Regex chassisNameRegex = new Regex(@"^([A-zА-я]){1,20}$");
                            Match chassisNameMatch = chassisNameRegex.Match(chassisName);
                            if (chassisNameMatch.Success)
                            {
                                command.CommandText = $"Update ChassisType SET ChassisTypeName = @ChassisTypeName WHERE ChassisTypeID ={Chassis_dataGridView.SelectedRows[0].Cells[0].Value}";
                                command.Parameters.Add("@ChassisTypeName", System.Data.DbType.String).Value = chassisName;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
                case 8:
                    {
                        try
                        {
                            connection.Open();
                            command = new SQLiteCommand();
                            command.Connection = connection;
                            string autoType = AutoTypeNameBox.Text;
                            string pereodichnost = PereodichnostTechBox.Text;
                            string characteristics = CharacteristicsBox.Text;
                            Regex autoTypeRegex = new Regex(@"^([A-zА-я\s]){1,20}$");
                            Regex pereodichnostRegex = new Regex(@"^([0-9]){1,2}$");
                            Regex characteristicsRegex = new Regex(@"^([А-я,]){1,}$");
                            Match autoTypeMatch = autoTypeRegex.Match(autoType);
                            Match pereodichnostMatch = pereodichnostRegex.Match(pereodichnost.ToString());
                            Match characteristicsMatch = characteristicsRegex.Match(characteristics);
                            if (autoTypeMatch.Success && pereodichnostMatch.Success && characteristicsMatch.Success)
                            {
                                command.CommandText = $"Update AutoType SET AutoTypeName = @AutoTypeName, TechOsmotrRatio = @TechOsmotrRatio, СharacteristicsList = @СharacteristicsList WHERE AutoTypeID={AutoType_dataGridView.SelectedRows[0].Cells[0].Value}";
                                command.Parameters.Add("@AutoTypeName", System.Data.DbType.String).Value = autoType;
                                command.Parameters.Add("@TechOsmotrRatio", System.Data.DbType.String).Value = pereodichnost;
                                command.Parameters.Add("@СharacteristicsList", System.Data.DbType.String).Value = characteristics;

                                command.ExecuteNonQuery();
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
                        GetTables();
                        break;
                    }
            }
        }

        public class DTP
        {
            public string ChaosTypeName { get; set; }
            public string DeteOfDTP { get; set; }
            public string PlaceOfDTP { get; set; }
            public string CarNumber { get; set; }
        }
        static void DisplayInExcelDTP(IEnumerable<DTP> dtps)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;

            excelApp.Workbooks.Add();

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = "Происшествие";
            workSheet.Cells[1, "B"] = "Дата происшествия";
            workSheet.Cells[1, "C"] = "Место происшествия";
            workSheet.Cells[1, "D"] = "Номер машины";

            var row = 1;
            foreach (var dtp in dtps)
            {
                row++;
                workSheet.Cells[row, "A"] = dtp.ChaosTypeName;
                workSheet.Cells[row, "B"] = dtp.DeteOfDTP;
                workSheet.Cells[row, "C"] = dtp.PlaceOfDTP;
                workSheet.Cells[row, "D"] = dtp.CarNumber;
            }

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
            workSheet.Columns[4].AutoFit();
        }

        public class SearchCar
        {
            public string CarNuber { get; set; }
            public string CarMarka { get; set; }
            public string DateOfSearching { get; set; }
            public string InfoAboutSearching { get; set; }
            public string SearchingStatus { get; set; }
        }
        static void DisplayInExcelSerching(IEnumerable<SearchCar> cars)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;

            excelApp.Workbooks.Add();

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = "Номер автомобиля";
            workSheet.Cells[1, "B"] = "Дата розыска";
            workSheet.Cells[1, "C"] = "Информация о розыске";
            workSheet.Cells[1, "D"] = "Статус розыска";

            var row = 1;
            foreach (var car in cars)
            {
                row++;
                workSheet.Cells[row, "A"] = car.CarNuber;
                workSheet.Cells[row, "B"] = car.DateOfSearching;
                workSheet.Cells[row, "C"] = car.InfoAboutSearching;
                workSheet.Cells[row, "D"] = car.SearchingStatus;
            }

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
            workSheet.Columns[4].AutoFit();
        }
        private void ExcelButton_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                int totalRows = DTP_dataGridView.RowCount;
                var DTPList = new List<DTP>();
                for (int i = 0; i < totalRows; i++)
                {
                    DTPList.Add(new DTP
                    {
                        ChaosTypeName = DTP_dataGridView.Rows[i].Cells[1].Value.ToString(),
                        DeteOfDTP = DTP_dataGridView.Rows[i].Cells[8].Value.ToString(),
                        PlaceOfDTP = DTP_dataGridView.Rows[i].Cells[4].Value.ToString(),
                        CarNumber = DTP_dataGridView.Rows[i].Cells[5].Value.ToString(),
                    });
                }
                DisplayInExcelDTP(DTPList);
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                int totalRows = Searching_dataGridView.RowCount;
                var CarsList = new List<SearchCar>();
                for (int i = 0; i < totalRows; i++)
                {
                    CarsList.Add(new SearchCar
                    {
                        CarNuber = Searching_dataGridView.Rows[i].Cells[1].Value.ToString(),
                        DateOfSearching = Searching_dataGridView.Rows[i].Cells[3].Value.ToString(),
                        InfoAboutSearching = Searching_dataGridView.Rows[i].Cells[2].Value.ToString(),
                        SearchingStatus = Searching_dataGridView.Rows[i].Cells[4].Value.ToString()
                    });
                }
                DisplayInExcelSerching(CarsList);
            }
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            string searchValue = SearchBox.Text;

            int selectedIndex = tabControl1.SelectedIndex;

            switch (selectedIndex)
            {
                case 0:
                    {
                        try
                        {
                            DataSet ds = new DataSet();
                            connection.Open();
                            dataAdapter = new SQLiteDataAdapter($"SELECT * FROM DTP WHERE PlaceOfDTP LIKE '%{SearchBox.Text}%'", connection);
                            DataTable dt = ds.Tables.Add("DTP");

                            dataAdapter.Fill(dt);
                            DTP_dataGridView.DataSource = ds.Tables["DTP"];
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                            connection.Close();
                        }
                        break;
                    }
                case 1:
                    {
                        try
                        {
                            DataSet ds = new DataSet();
                            connection.Open();
                            dataAdapter = new SQLiteDataAdapter($"SELECT CarID, CarNumbersDirectory.CarNumber, AutoModel, EngineNum, EngineVolume, " +
                    $"ReleaseDate, AutoType.AutoTypeName, CarCaseType.CarCaseTypeName, ChassisType.ChassisTypeName FROM Cars " +
                    "INNER JOIN CarNumbersDirectory ON Cars.CarNumberID = CarNumbersDirectory.CarNumberID " +
                    "INNER JOIN AutoType ON Cars.AutoTypeID = AutoType.AutoTypeID " +
                    "INNER JOIN CarCaseType ON Cars.CarCaseTypeID = CarCaseType.CarCaseTypeID " +
                    "INNER JOIN ChassisType ON Cars.ChassisTypeID = ChassisType.ChassisTypeID " +
                    $"WHERE AutoModel LIKE '%{SearchBox.Text}%'", connection);
                            DataTable dt = ds.Tables.Add("Cars");

                            dataAdapter.Fill(dt);
                            Auto_dataGridView.DataSource = ds.Tables["Cars"];
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                            connection.Close();
                        }
                        break;
                    }
                case 4:
                    {
                        try
                        {
                            DataSet ds = new DataSet();
                            connection.Open();
                            dataAdapter = new SQLiteDataAdapter($"SELECT * FROM CarNumbersDirectory WHERE CarNumber LIKE '%{SearchBox.Text}%'", connection);
                            DataTable dt = ds.Tables.Add("CarNumbersDirectory");

                            dataAdapter.Fill(dt);
                            NumbersList_dataGridView.DataSource = ds.Tables["CarNumbersDirectory"];
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                            connection.Close();
                        }
                        break;
                    }

            }
        }
    }
}