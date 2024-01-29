using Guna.UI2.WinForms;
using Microsoft.VisualBasic;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Timers;
using System.Windows.Forms;
using TheArtOfDevHtmlRenderer.Adapters.Entities;
using static WinFormsApp1.Form1;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        private DateTime lastInteractionTime;
        public int speed { get; private set; }
        List<string> productIds = new List<string>();
        List<string> colors = new List<string>();
        List<string> types = new List<string>();
        public List<string> names = new List<string>();
        List<string> columnNames = new List<string>();
        List<int> indexOfNames = new List<int>();

        Dictionary<string, Color> colorCategories = new Dictionary<string, Color>
            {
                { "XL-F" ,Color.FromArgb(244, 176, 132) },
                { "L-F",Color.FromArgb(155, 194, 230) },
                { "L-FP",Color.FromArgb(142, 169, 219) },
                { "M-F",Color.FromArgb(189, 215, 238) },
                { "M-FP",Color.FromArgb(189, 215, 238) },
                { "S-F",Color.FromArgb(221, 235, 247) },
                { "L-B",Color.FromArgb(255, 217, 102) },
                { "M-B",Color.FromArgb(255, 230, 153) },
                { "S-B",Color.FromArgb(255, 242, 204) },
                { "C&M",Color.FromArgb(144,238,144) },
                { "S-P",Color.FromArgb(248, 203, 173) }
            };

        string[] itemsToAdd = { "XL-F", "L-F", "L-FP", "M-F", "M-FP", "S-F", "L-B", "M-B", "S-B", "S-P", "C&M" };

        DataTable table = new DataTable();
        int columnCount;
        private int rowIndexFromMouseDown;
        private DataGridViewRow movingRow;

        DataGridViewCheckBoxColumn check = new DataGridViewCheckBoxColumn();
        DataGridViewTextBoxColumn textColumn = new DataGridViewTextBoxColumn();
        DataGridViewTextBoxColumn textColumn1 = new DataGridViewTextBoxColumn();

        //SqlConnection con = new SqlConnection("Data Source=DESKTOP-T79U63J\SQLEXPRESS;Initial Catalog=Test;Integrated Security=True"); <--------- should use specific Server Name and Database 
        string query1 = "SELECT TOP 1 Kiln_Car_number FROM GKloading ORDER BY ID DESC";
        string query2 = "SELECT color FROM colors";
        string query3 = "SELECT items FROM item";
        string query4 = "select Name from employee";
        string query5 = "select category from categories";
        string query11 = "select top 1 speed from GKloading order by DateandTime desc";


        List<string> Itemslist = new List<string>();
        List<string> Colorslist = new List<string>();

        private bool isDragging = false;
        private Point offset;

        private System.Timers.Timer dailyTimer;
        private DateTime desiredExecutionTime = DateTime.Today.Add(new TimeSpan(6, 30, 0));

        private bool mouseDown;
        private Point lastLocation;

        public Form1()
        {

            InitializeComponent();
            initialize();

            using (SqlConnection connection = new SqlConnection(Properties.Settings.Default.ConnectionString))
            {
                using (SqlCommand command = new SqlCommand(query11, connection))
                {
                    connection.Open();

                    using (SqlDataReader reader11 = command.ExecuteReader())
                    {
                        if (reader11.HasRows)
                        {
                            while (reader11.Read())
                            {
                                string s = reader11["speed"].ToString();
                                int n = int.Parse(s);
                                speed = n;
                            }
                        }
                        else
                        {
                            speed = 138;
                        }
                    }
                }
            }

            guna2TextBox2.TextAlign = HorizontalAlignment.Center;
            guna2TextBox1.TextAlign = HorizontalAlignment.Center;
            guna2TextBox1.Visible = false;
            guna2TextBox2.Visible = false;
            timer1.Start();

            SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);
            DateTime currentDate = DateTime.Now;


            //get colors to the combo box from the colors table 
            SqlCommand cmd2 = new SqlCommand(query2, con);
            con.Open();
            SqlDataReader reader = cmd2.ExecuteReader();
            while (reader.Read())
            {
                string data = reader.GetString(0);
                comboBox14.Items.Add(data);
            }
            con.Close();

            //get item numbers to the combo box from the item table  
            SqlCommand cmd3 = new SqlCommand(query3, con);
            con.Open();
            SqlDataReader reader1 = cmd3.ExecuteReader();
            while (reader1.Read())
            {
                string data1 = reader1.GetString(0);
                comboBox11.Items.Add(data1);
            }
            con.Close();


            //get emp numbers to the combo box from the employee table  
            SqlCommand cmd4 = new SqlCommand(query4, con);
            con.Open();
            SqlDataReader reader2 = cmd4.ExecuteReader();
            while (reader2.Read())
            {
                string data2 = reader2.GetString(0);
                //comboBox8.Items.Add(data2);
            }
            con.Close();


            //get kiln car numbers data to show the next kiln car number 
            SqlCommand cmd = new SqlCommand(query1, con);
            con.Open();
            var car_number = cmd.ExecuteScalar();
            int x = Convert.ToInt32(car_number);
            con.Close();

            //get_carnumber();
            settimer();

        }


        //connection to database 
        private void initialize()
        {
            bool islook = true;
            while (islook)
            {
                if (Properties.Settings.Default.ConnectionString == "")
                {
                    string input = Interaction.InputBox("Enter the Connection String:", "Connecting Database");
                    if (input == "")
                    {
                        MessageBox.Show("Enter valid Connection String", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.Close();
                    }
                    else
                    {
                        try
                        {
                            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(input);
                            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                            {
                                Properties.Settings.Default.ConnectionString = input;
                                MessageBox.Show("Connection string is valid!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                Properties.Settings.Default.Save();
                            }
                            islook = false;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Invalid connection string: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    islook = false;
                }
            }
        }


        //setting the timer to execute
        private void settimer()
        {
            // Calculate the time until the first desired execution
            TimeSpan timeUntilExecution = desiredExecutionTime - DateTime.Now;
            if (timeUntilExecution.TotalMilliseconds < 0)
            {
                // If the desired time has already passed for today, set it for tomorrow
                desiredExecutionTime = desiredExecutionTime.AddDays(1);
                timeUntilExecution = desiredExecutionTime - DateTime.Now;
            }

            dailyTimer = new System.Timers.Timer();
            dailyTimer.Interval = timeUntilExecution.TotalMilliseconds;
            dailyTimer.Elapsed += DailyTimer_Elapsed;
            dailyTimer.AutoReset = false; // Set this to false so that the timer only triggers once
            dailyTimer.Start();
        }




        //insert new color into combo box,colors table
        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);
            string input = Interaction.InputBox("නව වර්ණය ඇතුළත් කරන්න:", "නව වර්ණය");
            if (!string.IsNullOrEmpty(input))
            {
                SqlCommand cmd2 = new SqlCommand("INSERT INTO [dbo].[colors]\r\n           ([color])\r\n     VALUES\r\n           ('" + input + "')", con);
                con.Open();
                cmd2.ExecuteNonQuery();
                con.Close();
                comboBox14.Items.Add(input);
            }
        }


        //insert new item number into combo box,item table
        private void button5_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);
            string item = "";
            string category = "";
            ShowInput(ref item, ref category, 300, 200);
            if (!string.IsNullOrEmpty(item) && (!string.IsNullOrEmpty(category)))
            {
                SqlCommand cmd2 = new SqlCommand("INSERT INTO [dbo].[item]([items],[category])VALUES('" + item + "','" + category + "')", con);
                con.Open();
                cmd2.ExecuteNonQuery();
                con.Close();
                comboBox11.Items.Add(item);
                MessageBox.Show("අයිතමය ඇතුළත් කිරීම සාර්ථකයි", "සාර්ථකයි", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("හිස් ක්ෂේත්‍ර හෝ සංවාද කොටුව වසා ඇත", "විස්තරය", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Dialog Box to Input New Employee 
        private static DialogResult ShowInput(ref string item, ref string category, int width = 300, int height = 200)
        {
            Size size = new Size(width, height);
            Form inputBox = new Form();
            inputBox.Location = new Point(0, 0);
            inputBox.MaximizeBox = false;
            inputBox.FormBorderStyle = FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "නව අයිතම තොරතුරු";

            Label label = new Label();
            label.Text = "අයිතම කේතය ඇතුල් කරන්න:";
            label.Location = new Point(5, 5);
            label.Width = size.Width - 10;
            inputBox.Controls.Add(label);

            TextBox textBox = new TextBox();
            textBox.Size = new Size(size.Width - 10, 23);
            textBox.Location = new Point(5, label.Location.Y + 20);
            inputBox.Controls.Add(textBox);

            Label label1 = new Label();
            label1.Text = "කාණ්ඩය ඇතුළු කරන්න:";
            label1.Location = new Point(5, textBox.Location.Y + 30);
            label1.Width = size.Width - 10;
            inputBox.Controls.Add(label1);

            ComboBox comboBox = new ComboBox();
            comboBox.Size = new Size(size.Width - 25, 23);
            comboBox.Location = new Point(5, label1.Location.Y + 20);
            comboBox.Items.Add("XL-F");
            comboBox.Items.Add("L-F");
            comboBox.Items.Add("L-FP");
            comboBox.Items.Add("M-F");
            comboBox.Items.Add("M-FP");
            comboBox.Items.Add("S-F");
            comboBox.Items.Add("L-B");
            comboBox.Items.Add("M-B");
            comboBox.Items.Add("S-B");
            comboBox.Items.Add("S-P");
            comboBox.Items.Add("C&M");
            inputBox.Controls.Add(comboBox);

            Button okButton = new Button();
            okButton.DialogResult = DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new Point(size.Width - 80 - 80, size.Height - 30);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new Point(size.Width - 80, size.Height - 30);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;
            DialogResult result = inputBox.ShowDialog();
            item = textBox.Text;
            category = comboBox.Text;
            return result;
        }


        //get data from database and save in a excel file in desktop as 'GK Loading.xlsx'
        private void excelfilesaver()
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filename = "Glost Kiln Loading.xlsx";
                string filepath = Path.Combine(desktopPath, filename);

                DateTime Date = DateTime.Now.Date.AddHours(6);
                DateTime startDate = Date.AddDays(-1);
                DateTime endDate = DateTime.Now.Date.AddHours(6);
                string connectionString = Properties.Settings.Default.ConnectionString;
                string query = "SELECT Kiln_Car_number,Type,ProductID,Color,Quantity,Loader,DateandTime FROM GKloading WHERE DateandTime > @startDate AND DateandTime <= @endDate order by DateandTime";
                string query1 = "SELECT ProductID,Color,sum(Quantity) As Total FROM GKloading where Type IS NULL and DateandTime > @startDate AND DateandTime <= @endDate group by ProductID,Color";
                string query8 = "SELECT ProductID,Color,sum(Quantity) As Total FROM GKloading where Type = 'R' and DateandTime > @startDate AND DateandTime <= @endDate group by ProductID,Color";
                string query2 = "select distinct shift, loader from Loader where DateTime > @startDate and DateTime <= @endDate";
                string query5 = "select kiln_car_number,sum(quantity) as Total from GKloading where DateandTime > @startDate and DateandTime <= @endDate group by Kiln_Car_number,DateandTime";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(query, connection);
                    command.Parameters.AddWithValue("@startDate", startDate);
                    command.Parameters.AddWithValue("@endDate", endDate);
                    SqlDataReader reader = command.ExecuteReader();
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(new FileInfo(filepath)))
                    {
                        DateTime date = DateTime.Now;
                        string currentDate1 = DateTime.Now.AddDays(-1).ToString("MMM dd");
                        var worksheet = package.Workbook.Worksheets.Add(currentDate1);
                        // Add headers
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            worksheet.Cells[1, i + 1].Value = reader.GetName(i);
                            worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                            worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            System.Drawing.Color customColor1 = System.Drawing.Color.FromArgb(30, 144, 255);
                            worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(customColor1);
                            worksheet.Cells[1, i + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[1, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        }
                        // Add rows
                        int row = 2;
                        while (reader.Read())
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                if (i == 6)
                                {
                                    DateTime dateTimeValue = (DateTime)reader[i];
                                    worksheet.Cells[row, i + 1].Value = dateTimeValue;
                                    worksheet.Cells[row, i + 1].Style.Numberformat.Format = "yyyy/MM/dd HH:mm";
                                    worksheet.Cells[row, i + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[row, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                }
                                else
                                {
                                    worksheet.Cells[row, i + 1].Value = reader[i];
                                    worksheet.Cells[row, i + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[row, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                }
                            }
                            row++;
                        }
                        connection.Close();
                        SqlCommand command1 = new SqlCommand(query1, connection);
                        connection.Open();
                        command1.Parameters.AddWithValue("@startDate", startDate);
                        command1.Parameters.AddWithValue("@endDate", endDate);
                        SqlDataReader reader1 = command1.ExecuteReader();
                        // Add headers
                        worksheet.Cells[1, 9, 1, 11].Merge = true;
                        worksheet.Cells[1, 9].Value = "Glazeware";
                        worksheet.Cells[2, 9].Value = "ProductID";
                        worksheet.Cells[2, 10].Value = "Color";
                        worksheet.Cells[2, 11].Value = "Total";
                        worksheet.Cells[2, 9].Style.Font.Bold = true;
                        worksheet.Cells[2, 10].Style.Font.Bold = true;
                        worksheet.Cells[2, 11].Style.Font.Bold = true;
                        System.Drawing.Color customColor = System.Drawing.Color.FromArgb(30, 144, 255);
                        System.Drawing.Color customColor2 = System.Drawing.Color.FromArgb(255, 255, 102);
                        worksheet.Cells[1, 9, 1, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[2, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[2, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[2, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[2, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[2, 10].Style.Fill.BackgroundColor.SetColor(customColor);
                        worksheet.Cells[1, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[1, 9].Style.Fill.BackgroundColor.SetColor(customColor2);
                        worksheet.Cells[2, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[2, 9].Style.Fill.BackgroundColor.SetColor(customColor);
                        worksheet.Cells[2, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[2, 11].Style.Fill.BackgroundColor.SetColor(customColor);
                        // Add rows
                        int r = 3;
                        int Total = 0;
                        while (reader1.Read())
                        {
                            worksheet.Cells[r, 9].Value = reader1["ProductID"];
                            worksheet.Cells[r, 10].Value = reader1["Color"];
                            worksheet.Cells[r, 11].Value = reader1["Total"];
                            worksheet.Cells[r, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                            int x = (int)worksheet.Cells[r, 11].Value;
                            Total = Total + x;
                            r++;
                        }
                        if (Total != 0)
                        {
                            System.Drawing.Color customColor3 = System.Drawing.Color.FromArgb(255, 215, 0);
                            worksheet.Cells[r + 1, 10].Value = "Total";
                            worksheet.Cells[r + 1, 10].Style.Font.Bold = true;
                            worksheet.Cells[r + 1, 11].Value = Total;
                            worksheet.Cells[r + 1, 11].Style.Font.Bold = true;
                            worksheet.Cells[r + 1, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r + 1, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r + 1, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r + 1, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r + 1, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[r + 1, 10].Style.Fill.BackgroundColor.SetColor(customColor3);
                            worksheet.Cells[r + 1, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[r + 1, 11].Style.Fill.BackgroundColor.SetColor(customColor3);
                        }
                        connection.Close();
                        SqlCommand command8 = new SqlCommand(query8, connection);
                        connection.Open();
                        command8.Parameters.AddWithValue("@startDate", startDate);
                        command8.Parameters.AddWithValue("@endDate", endDate);
                        SqlDataReader reader8 = command8.ExecuteReader();
                        // Add headers
                        worksheet.Cells[1, 13, 1, 15].Merge = true;
                        worksheet.Cells[1, 13].Value = "Repair";
                        worksheet.Cells[2, 13].Value = "ProductID";
                        worksheet.Cells[2, 14].Value = "Color";
                        worksheet.Cells[2, 15].Value = "Total";
                        worksheet.Cells[2, 13].Style.Font.Bold = true;
                        worksheet.Cells[2, 14].Style.Font.Bold = true;
                        worksheet.Cells[2, 15].Style.Font.Bold = true;
                        System.Drawing.Color customColor6 = System.Drawing.Color.FromArgb(30, 144, 255);
                        System.Drawing.Color customColor4 = System.Drawing.Color.FromArgb(255, 255, 102);
                        worksheet.Cells[1, 13, 1, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[2, 13].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[2, 14].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[2, 15].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[2, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[2, 14].Style.Fill.BackgroundColor.SetColor(customColor6);
                        worksheet.Cells[1, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[1, 13].Style.Fill.BackgroundColor.SetColor(customColor4);
                        worksheet.Cells[2, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[2, 13].Style.Fill.BackgroundColor.SetColor(customColor6);
                        worksheet.Cells[2, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[2, 15].Style.Fill.BackgroundColor.SetColor(customColor6);
                        // Add rows
                        int r3 = 3;
                        int Total1 = 0;
                        while (reader8.Read())
                        {
                            worksheet.Cells[r3, 13].Value = reader8["ProductID"];
                            worksheet.Cells[r3, 14].Value = reader8["Color"];
                            worksheet.Cells[r3, 15].Value = reader8["Total"];
                            worksheet.Cells[r3, 13].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r3, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r3, 14].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r3, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r3, 15].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r3, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                            int x = (int)worksheet.Cells[r3, 15].Value;
                            Total1 = Total1 + x;
                            r3++;
                        }
                        if (Total1 != 0)
                        {
                            System.Drawing.Color customColor5 = System.Drawing.Color.FromArgb(255, 215, 0);
                            worksheet.Cells[r3 + 1, 14].Value = "Total";
                            worksheet.Cells[r3 + 1, 14].Style.Font.Bold = true;
                            worksheet.Cells[r3 + 1, 15].Value = Total1;
                            worksheet.Cells[r3 + 1, 15].Style.Font.Bold = true;
                            worksheet.Cells[r3 + 1, 14].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r3 + 1, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r3 + 1, 15].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r3 + 1, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r3 + 1, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[r3 + 1, 14].Style.Fill.BackgroundColor.SetColor(customColor5);
                            worksheet.Cells[r3 + 1, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[r3 + 1, 15].Style.Fill.BackgroundColor.SetColor(customColor5);
                        }
                        connection.Close();

                        SqlCommand command3 = new SqlCommand(query5, connection);
                        connection.Open();
                        command3.Parameters.AddWithValue("@startDate", startDate);
                        command3.Parameters.AddWithValue("@endDate", endDate);
                        SqlDataReader reader3 = command3.ExecuteReader();
                        // Add headers                
                        worksheet.Cells[1, 17].Value = "Kiln Car Number";
                        worksheet.Cells[1, 18].Value = "Total";
                        worksheet.Cells[1, 17].Style.Font.Bold = true;
                        worksheet.Cells[1, 18].Style.Font.Bold = true;
                        worksheet.Cells[1, 17].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 17].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[1, 18].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 18].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[1, 18].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[1, 18].Style.Fill.BackgroundColor.SetColor(customColor);
                        worksheet.Cells[1, 17].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[1, 17].Style.Fill.BackgroundColor.SetColor(customColor);
                        // Add rows
                        int r1 = 2;
                        while (reader3.Read())
                        {
                            worksheet.Cells[r1, 17].Value = reader3["Kiln_Car_number"];
                            worksheet.Cells[r1, 18].Value = reader3["Total"];
                            worksheet.Cells[r1, 17].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r1, 17].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r1, 18].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r1, 18].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            r1++;
                        }
                        connection.Close();

                        SqlCommand command2 = new SqlCommand(query2, connection);
                        connection.Open();
                        command2.Parameters.AddWithValue("@startDate", startDate);
                        command2.Parameters.AddWithValue("@endDate", endDate);
                        SqlDataReader reader2 = command2.ExecuteReader();
                        // Add headers                
                        worksheet.Cells[1, 20].Value = "Shift";
                        worksheet.Cells[1, 21].Value = "Loader";
                        worksheet.Cells[1, 20].Style.Font.Bold = true;
                        worksheet.Cells[1, 21].Style.Font.Bold = true;
                        worksheet.Cells[1, 20].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 20].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[1, 21].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 21].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[1, 21].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[1, 21].Style.Fill.BackgroundColor.SetColor(customColor);
                        worksheet.Cells[1, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[1, 20].Style.Fill.BackgroundColor.SetColor(customColor);
                        // Add rows
                        int r2 = 2;
                        while (reader2.Read())
                        {
                            worksheet.Cells[r2, 20].Value = reader2["Shift"];
                            worksheet.Cells[r2, 21].Value = reader2["loader"];
                            worksheet.Cells[r2, 20].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r2, 20].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r2, 21].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r2, 21].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            r2++;
                        }
                        connection.Close();
                        worksheet.Cells.AutoFitColumns();
                        worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        // Save the Excel file
                        package.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("යම් දෝෂයක් ඇත..ස්වයංක්‍රීයව සුරැකීමට නොහැකි!", "දෝෂය", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        //setting tmer to execute 
        private void DailyTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            // This method will be executed once a day at the desired time        
            excelfilesaver();

            // Calculate the time until the next desired execution
            desiredExecutionTime = desiredExecutionTime.AddDays(1);
            TimeSpan timeUntilExecution = desiredExecutionTime - DateTime.Now;

            // Reset and restart the timer for the next day's execution
            dailyTimer.Interval = timeUntilExecution.TotalMilliseconds;
            dailyTimer.Start();
        }




        // excel file save for save button
        private void bunifuFlatButton3_Click(object sender, EventArgs e)
        {
            export();
        }


        // Export excel file for selected date 
        private void export()
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                DateTime StartDate = guna2DateTimePicker1.Value.Date.AddHours(6);
                DateTime EndDate = StartDate.AddDays(1);
                DateTime enter = guna2DateTimePicker1.Value;
                DateTime now1 = DateTime.Now;

                if (enter > now1)
                {
                    MessageBox.Show("කරුණාකර වලංගු දිනයක් තෝරන්න", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    using (SqlConnection connection = new SqlConnection(Properties.Settings.Default.ConnectionString))
                    {
                        connection.Open();
                        string query = "SELECT Kiln_Car_number,Type,ProductID,Color,Quantity,Loader,DateandTime FROM GKloading WHERE DateandTime > @startDate AND DateandTime <= @endDate order by DateandTime";
                        string query1 = "SELECT ProductID,Color,sum(Quantity) As Total FROM GKloading where Type IS NULL and DateandTime > @startDate AND DateandTime <= @endDate group by ProductID,Color";
                        string query8 = "SELECT ProductID,Color,sum(Quantity) As Total FROM GKloading where Type = 'R' and DateandTime > @startDate AND DateandTime <= @endDate group by ProductID,Color";
                        string query2 = "select distinct shift, loader from Loader where DateTime > @startDate and DateTime <= @endDate";
                        string query5 = "select kiln_car_number,sum(quantity) as Total from GKloading where DateandTime > @startDate and DateandTime <= @endDate group by Kiln_Car_number,DateandTime";
                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@startDate", StartDate);
                            command.Parameters.AddWithValue("@endDate", EndDate);
                            SqlDataReader reader = command.ExecuteReader();
                            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                            using (ExcelPackage package = new ExcelPackage())
                            {
                                string Date = guna2DateTimePicker1.Value.ToString("MMM dd");
                                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(Date);
                                // Add headers
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    worksheet.Cells[1, i + 1].Value = reader.GetName(i);
                                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                                    worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    System.Drawing.Color customColor1 = System.Drawing.Color.FromArgb(30, 144, 255);
                                    worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(customColor1);
                                    worksheet.Cells[1, i + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[1, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                }
                                // Add rows
                                int row = 2;
                                while (reader.Read())
                                {
                                    for (int i = 0; i < reader.FieldCount; i++)
                                    {
                                        if (i == 6)
                                        {
                                            DateTime dateTimeValue = (DateTime)reader[i];
                                            worksheet.Cells[row, i + 1].Value = dateTimeValue;
                                            worksheet.Cells[row, i + 1].Style.Numberformat.Format = "yyyy/MM/dd HH:mm";
                                            worksheet.Cells[row, i + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                            worksheet.Cells[row, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                        }
                                        else
                                        {
                                            worksheet.Cells[row, i + 1].Value = reader[i];
                                            worksheet.Cells[row, i + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                            worksheet.Cells[row, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                        }
                                    }
                                    row++;
                                }
                                connection.Close();
                                SqlCommand command1 = new SqlCommand(query1, connection);
                                connection.Open();
                                command1.Parameters.AddWithValue("@startDate", StartDate);
                                command1.Parameters.AddWithValue("@endDate", EndDate);
                                SqlDataReader reader1 = command1.ExecuteReader();
                                // Add headers
                                worksheet.Cells[1, 9, 1, 11].Merge = true;
                                worksheet.Cells[1, 9].Value = "Glazeware";
                                worksheet.Cells[2, 9].Value = "ProductID";
                                worksheet.Cells[2, 10].Value = "Color";
                                worksheet.Cells[2, 11].Value = "Total";
                                worksheet.Cells[2, 9].Style.Font.Bold = true;
                                worksheet.Cells[2, 10].Style.Font.Bold = true;
                                worksheet.Cells[2, 11].Style.Font.Bold = true;
                                System.Drawing.Color customColor = System.Drawing.Color.FromArgb(30, 144, 255);
                                System.Drawing.Color customColor2 = System.Drawing.Color.FromArgb(255, 255, 102);
                                worksheet.Cells[1, 9, 1, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[2, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[2, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[2, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[2, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[2, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[2, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[2, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[2, 10].Style.Fill.BackgroundColor.SetColor(customColor);
                                worksheet.Cells[1, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[1, 9].Style.Fill.BackgroundColor.SetColor(customColor2);
                                worksheet.Cells[2, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[2, 9].Style.Fill.BackgroundColor.SetColor(customColor);
                                worksheet.Cells[2, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[2, 11].Style.Fill.BackgroundColor.SetColor(customColor);
                                // Add rows
                                int r = 3;
                                int Total = 0;
                                while (reader1.Read())
                                {
                                    worksheet.Cells[r, 9].Value = reader1["ProductID"];
                                    worksheet.Cells[r, 10].Value = reader1["Color"];
                                    worksheet.Cells[r, 11].Value = reader1["Total"];
                                    worksheet.Cells[r, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    worksheet.Cells[r, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    worksheet.Cells[r, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                                    int x = (int)worksheet.Cells[r, 11].Value;
                                    Total = Total + x;
                                    r++;
                                }
                                if (Total != 0)
                                {
                                    System.Drawing.Color customColor3 = System.Drawing.Color.FromArgb(255, 215, 0);
                                    worksheet.Cells[r + 1, 10].Value = "Total";
                                    worksheet.Cells[r + 1, 10].Style.Font.Bold = true;
                                    worksheet.Cells[r + 1, 11].Value = Total;
                                    worksheet.Cells[r + 1, 11].Style.Font.Bold = true;
                                    worksheet.Cells[r + 1, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r + 1, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    worksheet.Cells[r + 1, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r + 1, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    worksheet.Cells[r + 1, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    worksheet.Cells[r + 1, 10].Style.Fill.BackgroundColor.SetColor(customColor3);
                                    worksheet.Cells[r + 1, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    worksheet.Cells[r + 1, 11].Style.Fill.BackgroundColor.SetColor(customColor3);
                                }
                                connection.Close();
                                SqlCommand command8 = new SqlCommand(query8, connection);
                                connection.Open();
                                command8.Parameters.AddWithValue("@startDate", StartDate);
                                command8.Parameters.AddWithValue("@endDate", EndDate);
                                SqlDataReader reader8 = command8.ExecuteReader();
                                // Add headers
                                worksheet.Cells[1, 13, 1, 15].Merge = true;
                                worksheet.Cells[1, 13].Value = "Repair";
                                worksheet.Cells[2, 13].Value = "ProductID";
                                worksheet.Cells[2, 14].Value = "Color";
                                worksheet.Cells[2, 15].Value = "Total";
                                worksheet.Cells[2, 13].Style.Font.Bold = true;
                                worksheet.Cells[2, 14].Style.Font.Bold = true;
                                worksheet.Cells[2, 15].Style.Font.Bold = true;
                                System.Drawing.Color customColor6 = System.Drawing.Color.FromArgb(30, 144, 255);
                                System.Drawing.Color customColor4 = System.Drawing.Color.FromArgb(255, 255, 102);
                                worksheet.Cells[1, 13, 1, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[2, 13].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[2, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[2, 14].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[2, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[2, 15].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[2, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[2, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[2, 14].Style.Fill.BackgroundColor.SetColor(customColor6);
                                worksheet.Cells[1, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[1, 13].Style.Fill.BackgroundColor.SetColor(customColor4);
                                worksheet.Cells[2, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[2, 13].Style.Fill.BackgroundColor.SetColor(customColor6);
                                worksheet.Cells[2, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[2, 15].Style.Fill.BackgroundColor.SetColor(customColor6);
                                // Add rows
                                int r3 = 3;
                                int Total1 = 0;
                                while (reader8.Read())
                                {
                                    worksheet.Cells[r3, 13].Value = reader8["ProductID"];
                                    worksheet.Cells[r3, 14].Value = reader8["Color"];
                                    worksheet.Cells[r3, 15].Value = reader8["Total"];
                                    worksheet.Cells[r3, 13].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r3, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    worksheet.Cells[r3, 14].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r3, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    worksheet.Cells[r3, 15].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r3, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                                    int x = (int)worksheet.Cells[r3, 15].Value;
                                    Total1 = Total1 + x;
                                    r3++;
                                }
                                if (Total1 != 0)
                                {
                                    System.Drawing.Color customColor5 = System.Drawing.Color.FromArgb(255, 215, 0);
                                    worksheet.Cells[r3 + 1, 14].Value = "Total";
                                    worksheet.Cells[r3 + 1, 14].Style.Font.Bold = true;
                                    worksheet.Cells[r3 + 1, 15].Value = Total1;
                                    worksheet.Cells[r3 + 1, 15].Style.Font.Bold = true;
                                    worksheet.Cells[r3 + 1, 14].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r3 + 1, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    worksheet.Cells[r3 + 1, 15].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r3 + 1, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    worksheet.Cells[r3 + 1, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    worksheet.Cells[r3 + 1, 14].Style.Fill.BackgroundColor.SetColor(customColor5);
                                    worksheet.Cells[r3 + 1, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    worksheet.Cells[r3 + 1, 15].Style.Fill.BackgroundColor.SetColor(customColor5);
                                }
                                connection.Close();

                                SqlCommand command3 = new SqlCommand(query5, connection);
                                connection.Open();
                                command3.Parameters.AddWithValue("@startDate", StartDate);
                                command3.Parameters.AddWithValue("@endDate", EndDate);
                                SqlDataReader reader3 = command3.ExecuteReader();
                                // Add headers                
                                worksheet.Cells[1, 17].Value = "Kiln Car Number";
                                worksheet.Cells[1, 18].Value = "Total";
                                worksheet.Cells[1, 17].Style.Font.Bold = true;
                                worksheet.Cells[1, 18].Style.Font.Bold = true;
                                worksheet.Cells[1, 17].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[1, 17].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[1, 18].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[1, 18].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[1, 18].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[1, 18].Style.Fill.BackgroundColor.SetColor(customColor);
                                worksheet.Cells[1, 17].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[1, 17].Style.Fill.BackgroundColor.SetColor(customColor);
                                // Add rows
                                int r1 = 2;
                                while (reader3.Read())
                                {
                                    worksheet.Cells[r1, 17].Value = reader3["Kiln_Car_number"];
                                    worksheet.Cells[r1, 18].Value = reader3["Total"];
                                    worksheet.Cells[r1, 17].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r1, 17].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    worksheet.Cells[r1, 18].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r1, 18].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    r1++;
                                }
                                connection.Close();

                                SqlCommand command2 = new SqlCommand(query2, connection);
                                connection.Open();
                                command2.Parameters.AddWithValue("@startDate", StartDate);
                                command2.Parameters.AddWithValue("@endDate", EndDate);
                                SqlDataReader reader2 = command2.ExecuteReader();
                                // Add headers                
                                worksheet.Cells[1, 20].Value = "Shift";
                                worksheet.Cells[1, 21].Value = "Loader";
                                worksheet.Cells[1, 20].Style.Font.Bold = true;
                                worksheet.Cells[1, 21].Style.Font.Bold = true;
                                worksheet.Cells[1, 20].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[1, 20].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[1, 21].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[1, 21].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                worksheet.Cells[1, 21].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[1, 21].Style.Fill.BackgroundColor.SetColor(customColor);
                                worksheet.Cells[1, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[1, 20].Style.Fill.BackgroundColor.SetColor(customColor);
                                // Add rows
                                int r2 = 2;
                                while (reader2.Read())
                                {
                                    worksheet.Cells[r2, 20].Value = reader2["Shift"];
                                    worksheet.Cells[r2, 21].Value = reader2["loader"];
                                    worksheet.Cells[r2, 20].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r2, 20].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    worksheet.Cells[r2, 21].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[r2, 21].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                    r2++;
                                }
                                connection.Close();
                                worksheet.Cells.AutoFitColumns();
                                worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                // FileInfo file = new FileInfo($"C:\\Users\\damih\\Desktop\\{Date} GK Loading.xlsx");
                                string filename = $"{Date} GK Loading.xlsx";
                                string filepath = Path.Combine(desktopPath, filename);
                                FileInfo file = new FileInfo(filepath);
                                try
                                {
                                    package.SaveAs(file);
                                    MessageBox.Show("දත්ත සාර්ථකව සුරකින ලදී", "සාර්ථකයි", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("ගොනුව දැනට විවෘත කර ඇති නිසා එය සුරැකිය නොහැක.", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("දත්ත සුරැකීම සාර්ථක නොවේ", "දෝෂය", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // Save Excel File for the Total Loadings(Item and Color wise) for the current month
        private void bunifuFlatButton11_Click(object sender, EventArgs e)
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                DateTime currentDate = DateTime.Now;
                string formattedDate = currentDate.ToString("MMM");
                DateTime startDate = new DateTime(currentDate.Year, currentDate.Month, 1, 6, 0, 0);
                // Iterate through the dates of the current month
                int i = 1;
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                ExcelPackage package = new ExcelPackage();
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"{currentDate.ToString("MMM")}-S");
                for (DateTime StartDate = startDate; StartDate <= currentDate; StartDate = StartDate.AddDays(1))
                {
                    using (SqlConnection connection = new SqlConnection(Properties.Settings.Default.ConnectionString))
                    {
                        string query1 = "SELECT ProductID,Color,sum(Quantity) As Total FROM GKloading where DateandTime > @startDate AND DateandTime <= @endDate group by ProductID,Color";
                        DateTime EndDate = StartDate.AddDays(1);
                        SqlCommand command1 = new SqlCommand(query1, connection);
                        connection.Open();
                        command1.Parameters.AddWithValue("@startDate", StartDate);
                        command1.Parameters.AddWithValue("@endDate", EndDate);
                        SqlDataReader reader1 = command1.ExecuteReader();
                        // Add headers
                        worksheet.Cells[1, i, 1, i + 2].Merge = true;
                        worksheet.Cells[1, i].Value = StartDate.ToString("MMM-dd");
                        worksheet.Cells[2, i].Value = "ProductID";
                        worksheet.Cells[2, i + 1].Value = "Color";
                        worksheet.Cells[2, i + 2].Value = "Total";
                        worksheet.Cells[1, i].Style.Font.Bold = true;
                        worksheet.Cells[2, i].Style.Font.Bold = true;
                        worksheet.Cells[1, i].Style.Font.Size = 14;
                        worksheet.Cells[2, i].Style.Font.Size = 14;
                        worksheet.Cells[2, i + 1].Style.Font.Size = 14;
                        worksheet.Cells[2, i + 2].Style.Font.Size = 14;
                        worksheet.Cells[2, i + 1].Style.Font.Bold = true;
                        worksheet.Cells[2, i + 2].Style.Font.Bold = true;
                        System.Drawing.Color customColor = System.Drawing.Color.FromArgb(30, 144, 255);
                        worksheet.Cells[1, i, 1, i + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[2, i].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[2, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[2, i + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        worksheet.Cells[1, i, 1, i + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[1, i, 1, i + 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 181, 86));
                        worksheet.Cells[2, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[2, i + 1].Style.Fill.BackgroundColor.SetColor(customColor);
                        worksheet.Cells[2, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[2, i].Style.Fill.BackgroundColor.SetColor(customColor);
                        worksheet.Cells[2, i + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[2, i + 2].Style.Fill.BackgroundColor.SetColor(customColor);
                        // Add rows
                        int r = 3;
                        int Total = 0;
                        while (reader1.Read())
                        {
                            worksheet.Cells[r, i].Value = reader1["ProductID"];
                            worksheet.Cells[r, i + 1].Value = reader1["Color"];
                            worksheet.Cells[r, i + 2].Value = reader1["Total"];
                            worksheet.Cells[r, i].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r, i].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r, i + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r, i + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r, i + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            int x = (int)worksheet.Cells[r, i + 2].Value;
                            Total = Total + x;
                            r++;
                        }
                        if (Total != 0)
                        {
                            System.Drawing.Color customColor3 = System.Drawing.Color.FromArgb(255, 215, 0);
                            worksheet.Cells[r + 1, i + 1].Value = "Total";
                            worksheet.Cells[r + 1, i + 1].Style.Font.Bold = true;
                            worksheet.Cells[r + 1, i + 2].Value = Total;
                            worksheet.Cells[r + 1, i + 2].Style.Font.Bold = true;
                            worksheet.Cells[r + 1, i + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r + 1, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r + 1, i + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[r + 1, i + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            worksheet.Cells[r + 1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[r + 1, i + 1].Style.Fill.BackgroundColor.SetColor(customColor3);
                            worksheet.Cells[r + 1, i + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[r + 1, i + 2].Style.Fill.BackgroundColor.SetColor(customColor3);
                        }
                        worksheet.Cells.AutoFitColumns();
                        worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }
                    i += 4;
                }
                worksheet.View.FreezePanes(3, 1);
                // FileInfo file = new FileInfo($"C:\\Users\\damih\\Desktop\\{Date} GK Loading.xlsx");
                string filename = $"GK Loading {currentDate.ToString("MMM")}-S.xlsx";
                string filepath = Path.Combine(desktopPath, filename);
                FileInfo file = new FileInfo(filepath);
                try
                {
                    package.SaveAs(file);
                    MessageBox.Show("දත්ත සාර්ථකව සුරකින ලදී", "සාර්ථකයි", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ගොනුව දැනට විවෘත කර ඇති නිසා එය සුරැකිය නොහැක.", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("දත්ත සුරැකීම සාර්ථක නොවේ", "දෝෂය", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // check whether input is not null and whether it is letter or symbol 
        private bool ContainsLettersOrSymbols(string input)
        {
            if (!(string.IsNullOrEmpty(input)))
            {
                foreach (char c in input)
                {
                    if (char.IsLetter(c) || !char.IsLetterOrDigit(c))
                    {
                        return false;
                    }
                }
                return true;
            }
            else
            {
                return false;
            }
        }



        //insert new member into combo box, employee table
        private void button8_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);
            string name = "";
            ShowInputDialogBox(ref name, 300, 200);
            if (!string.IsNullOrEmpty(name))
            {
                SqlCommand cmd2 = new SqlCommand("INSERT INTO [dbo].[employee]([Name])VALUES('" + name + "')", con);
                con.Open();
                cmd2.ExecuteNonQuery();
                con.Close();
                //comboBox8.Items.Add(name);
                MessageBox.Show("සේවක නම ඇතුළත් කිරීම සාර්ථකයි", "සාර්ථකයි", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("හිස් ක්ෂේත්‍ර හෝ සංවාද කොටුව වසා ඇත", "විස්තරය", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Dialog Box to Input New Employee 
        private static DialogResult ShowInputDialogBox(ref string name, int width = 300, int height = 150)
        {
            Size size = new Size(width, height);
            Form inputBox = new Form();
            inputBox.Location = new Point(0, 0);
            inputBox.MaximizeBox = false;
            inputBox.FormBorderStyle = FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "නව සේවක තොරතුරු";

            Label label = new Label();
            label.Text = "සාමාජිකයාගේ නම ඇතුලත් කරන්න:";
            label.Location = new Point(5, 30);
            label.Width = size.Width - 10;
            inputBox.Controls.Add(label);

            TextBox textBox = new TextBox();
            textBox.Size = new Size(size.Width - 10, 23);
            textBox.Location = new Point(5, label.Location.Y + 20);
            inputBox.Controls.Add(textBox);

            Button okButton = new Button();
            okButton.DialogResult = DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new Point(size.Width - 80 - 80, size.Height - 30);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new Point(size.Width - 80, size.Height - 30);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;
            DialogResult result = inputBox.ShowDialog();
            name = textBox.Text;
            return result;
        }
        /*
        // Save Loader Details in Database
        private void bunifuFlatButton12_Click(object sender, EventArgs e)
        {
            DataTable dataTable = new DataTable();

            if ((checkBox1.Checked && !checkBox2.Checked) || (!checkBox1.Checked && checkBox2.Checked))
            {
                if (!string.IsNullOrEmpty(comboBox7.Text) && !string.IsNullOrEmpty(comboBox8.Text))
                {
                    if (names.Contains(comboBox8.Text))
                    {
                        MessageBox.Show("දත්ත දැනටමත් ඇතුළත් කර ඇත", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        DateTime currentDate = DateTime.Now;
                        SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);
                        SqlCommand cmd9 = new SqlCommand("INSERT INTO [dbo].[Loader]\r\n           ([DateTime]\r\n           ,[Shift]\r\n           ,[loader])\r\n     VALUES\r\n           ('" + currentDate + "','" + comboBox7.Text + "','" + comboBox8.Text + "')", con);
                        con.Open();
                        cmd9.ExecuteNonQuery();
                        con.Close();
                        names.Add(comboBox8.Text);
                        dataTable.Columns.Add("names", typeof(string));
                        foreach (string Item in names)
                        {
                            dataTable.Rows.Add(Item);
                        }
                        guna2DataGridView2.DataSource = dataTable;
                        guna2DataGridView2.AllowUserToAddRows = false;
                        MessageBox.Show("දත්ත සුරැකීම සාර්ථකයි", "සාර්ථකයි", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        comboBox8.Text = "";
                    }
                }
                else
                {
                    MessageBox.Show("කරුණාකර සියලුම ක්ෂේත්‍ර නිවැරදිව පුරවන්න", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("කරුණාකර සියලුම ක්ෂේත්‍ර නිවැරදිව පුරවන්න", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        */



        private void bunifuFlatButton15_Click(object sender, EventArgs e)
        {
            string item = comboBox11.Text;
            string color = comboBox14.Text;
            Itemslist.Add(item);
            Colorslist.Add(color);
        }

        private void bunifuFlatButton11_Click_1(object sender, EventArgs e)
        {
            foreach (string name in Itemslist)
            {
                DataGridViewTextBoxColumn text = new DataGridViewTextBoxColumn();
                text.HeaderText = name;
                guna2DataGridView1.Columns.Add(text);
            }
            guna2DataGridView1.Show();
        }


        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        /*
        // Add new Columns and Add Columns for existing Table 
        private void button1_Click(object sender, EventArgs e)
        {
            List<string> categories = new List<string>();

            categories.AddRange(itemsToAdd);

            string query = "SELECT category FROM item WHERE items = @item";
            SqlConnection connection = new SqlConnection(Properties.Settings.Default.ConnectionString);
            bool Isok = true;
            bool isDone = true;
            bool columnnotExists = true;
            foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
            {
                if (column.HeaderText == "Kiln Car")
                {
                    columnnotExists = false;
                    break;
                }
            }
            if (columnnotExists)
            {
                if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && guna2RadioButton1.Checked)
                {
                    string item = comboBox11.Text;
                    string color = comboBox14.Text;
                    if (string.IsNullOrEmpty(color))
                    {
                        color = "White";
                    }
                    string enteredColor = color;
                    string nickname = GetNicknameForColor(enteredColor);
                    productIds.Add(item);
                    colors.Add(color);
                    comboBox11.Text = "";
                    comboBox14.Text = "";
                    guna2RadioButton1.Checked = false;

                    textColumn1.HeaderText = "Kiln Car";
                    textColumn1.Name = "Kiln Car";
                    textColumn1.Width = 35;
                    columnNames.Add(textColumn1.Name);
                    guna2DataGridView1.Columns.Add(textColumn1);
                    int i = 0;
                    foreach (string Item in productIds)
                    {
                        DataGridViewTextBoxColumn Text = new DataGridViewTextBoxColumn();
                        Text.HeaderText = $"{Item} {nickname} (R)";
                        Text.Name = "Repair";
                        string itemNumber = Text.HeaderText.Split(' ')[0];
                        Text.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                        guna2DataGridView1.Columns.Add(Text);
                        columnNames.Add(Text.Name);
                        i++;
                    }
                    check.Name = "check";
                    check.ValueType = typeof(bool);
                    check.HeaderText = "Finished";
                    guna2DataGridView1.Columns.Add(check);
                    textColumn.Name = "Total";
                    textColumn.Width = 60;
                    textColumn.HeaderText = "Total";
                    columnNames.Add(check.Name);
                    columnNames.Add(textColumn.Name);
                    guna2DataGridView1.Columns.Add(textColumn);
                    guna2DataGridView1.DataSource = table;
                    SetRowCount(25);
                    guna2DataGridView1.AllowUserToAddRows = false;
                    guna2DataGridView1.Columns[0].Frozen = true;
                    guna2DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                    System.Drawing.Color desiredColor = System.Drawing.Color.PeachPuff;
                    guna2DataGridView1.Columns[0].DefaultCellStyle.BackColor = desiredColor;
                    guna2DataGridView1.Columns[0].HeaderCell.Style.BackColor = System.Drawing.Color.PaleVioletRed;
                    guna2DataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.FromArgb(244, 187, 255);
                    guna2DataGridView1.Columns[1].HeaderCell.Style.BackColor = Color.FromArgb(241, 167, 254);
                    guna2DataGridView1.AlternatingRowsDefaultCellStyle = null;
                    comboBox11.Focus();
                    Isok = false;

                }
                else if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && comboBox11.Text != "Test" && comboBox11.Text != "Sample" && Isok)
                {
                    bool Isok2 = false;
                    string item = comboBox11.Text;
                    string color = comboBox14.Text;
                    if (string.IsNullOrEmpty(color))
                    {
                        color = "White";
                    }
                    string enteredColor = color;
                    string nickname = GetNicknameForColor(enteredColor);

                    connection.Open();
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@item", item);

                        // Execute the SQL query
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int index = categories.IndexOf(category);
                                productIds.Add(item);
                                colors.Add(color);
                                textColumn1.HeaderText = "Kiln Car";
                                textColumn1.Name = "Kiln Car";
                                textColumn1.Width = 35;
                                columnNames.Add(textColumn1.Name);
                                guna2DataGridView1.Columns.Add(textColumn1);
                                int i = 0;
                                foreach (string Item in productIds)
                                {
                                    DataGridViewTextBoxColumn Text = new DataGridViewTextBoxColumn();
                                    Text.HeaderText = $"{Item} {nickname}";
                                    Text.Name = category;
                                    string itemNumber = Text.HeaderText.Split(' ')[0];
                                    Text.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                    columnNames.Add(Text.Name);
                                    indexOfNames.Add(index);
                                    guna2DataGridView1.Columns.Add(Text);
                                    Text.DefaultCellStyle.BackColor = colorCategories[category];
                                    Text.HeaderCell.Style.BackColor = colorCategories[category];
                                    i++;
                                }
                                check.Name = "check";
                                check.ValueType = typeof(bool);
                                check.HeaderText = "Finished";
                                guna2DataGridView1.Columns.Add(check);
                                textColumn.Name = "Total";
                                textColumn.HeaderText = "Total";
                                textColumn.Width = 60;
                                columnNames.Add(check.Name);
                                columnNames.Add(textColumn.Name);
                                guna2DataGridView1.Columns.Add(textColumn);
                                guna2DataGridView1.DataSource = table;
                                SetRowCount(25);

                                guna2DataGridView1.AllowUserToAddRows = false;
                                guna2DataGridView1.Columns[0].Frozen = true;
                                guna2DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                                System.Drawing.Color desiredColor = System.Drawing.Color.PeachPuff;
                                guna2DataGridView1.Columns[0].DefaultCellStyle.BackColor = desiredColor;
                                guna2DataGridView1.Columns[0].HeaderCell.Style.BackColor = System.Drawing.Color.PaleVioletRed;
                                guna2DataGridView1.AlternatingRowsDefaultCellStyle = null;
                                comboBox11.Text = "";
                                comboBox14.Text = "";
                                comboBox11.Focus();
                            }
                            else
                            {
                                Isok2 = false;
                            }
                        }
                    }
                    connection.Close();
                }
                else if ((comboBox11.Text == "Test" || comboBox11.Text == "Sample") && string.IsNullOrEmpty(comboBox14.Text) && !guna2RadioButton1.Checked)
                {
                    string item = comboBox11.Text;
                    productIds.Add(item);
                    colors.Add("");
                    comboBox11.Text = "";
                    comboBox14.Text = "";

                    textColumn1.HeaderText = "Kiln Car";
                    textColumn1.Name = "Kiln Car";
                    textColumn1.Width = 35;
                    columnNames.Add(textColumn1.Name);
                    guna2DataGridView1.Columns.Add(textColumn1);
                    int i = 0;
                    foreach (string Item in productIds)
                    {
                        DataGridViewTextBoxColumn Text = new DataGridViewTextBoxColumn();
                        Text.HeaderText = $"{Item} {colors[i]}";
                        Text.Name = "Test";
                        string itemNumber = Text.HeaderText.Split(' ')[0];
                        Text.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                        guna2DataGridView1.Columns.Add(Text);
                        columnNames.Add(Text.Name);
                        i++;
                    }
                    check.Name = "check";
                    check.ValueType = typeof(bool);
                    check.HeaderText = "Finished";
                    guna2DataGridView1.Columns.Add(check);
                    textColumn.Name = "Total";
                    textColumn.HeaderText = "Total";
                    textColumn.Width = 60;
                    columnNames.Add(check.Name);
                    columnNames.Add(textColumn.Name);
                    guna2DataGridView1.Columns.Add(textColumn);
                    guna2DataGridView1.DataSource = table;
                    SetRowCount(25);
                    guna2DataGridView1.AllowUserToAddRows = false;
                    guna2DataGridView1.Columns[0].Frozen = true;
                    guna2DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                    System.Drawing.Color desiredColor = System.Drawing.Color.PeachPuff;
                    guna2DataGridView1.Columns[0].DefaultCellStyle.BackColor = desiredColor;
                    guna2DataGridView1.Columns[0].HeaderCell.Style.BackColor = System.Drawing.Color.PaleVioletRed;
                    guna2DataGridView1.AlternatingRowsDefaultCellStyle = null;
                    comboBox11.Focus();
                }
                else
                {
                    MessageBox.Show("ක්ෂේත්‍ර නිවැරදිව පුරවන්න", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && guna2RadioButton1.Checked)
                {
                    bool columnNotExists = true;
                    string item = comboBox11.Text;
                    string color = comboBox14.Text;
                    if (string.IsNullOrEmpty(color))
                    {
                        color = "White";
                    }
                    string enteredColor = color;
                    bool isColorAvailable = GetAvailableColor(enteredColor);
                    if (isColorAvailable)
                    {
                        string nickname = GetNicknameForColor(enteredColor);
                        foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                        {
                            if (column.HeaderText == $"{item} {nickname} (R)")
                            {
                                if (color == colors[column.Index - 1])
                                {
                                    columnNotExists = false;
                                    break;
                                }
                            }
                        }
                        if (columnNotExists)
                        {
                            productIds.Add(item);
                            colors.Add(color);

                            DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                            newColumn.HeaderText = $"{item} {nickname} (R)";
                            newColumn.Name = "Repair";
                            string itemNumber = newColumn.HeaderText.Split(' ')[0];
                            newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                            columnNames.Insert(guna2DataGridView1.Columns.Count - 2, newColumn.Name);

                            int insertionIndex = guna2DataGridView1.Columns.Count - 2;
                            guna2DataGridView1.Columns.Insert(insertionIndex, newColumn);
                            guna2DataGridView1.Columns[guna2DataGridView1.Columns.Count - 3].DefaultCellStyle.BackColor = Color.FromArgb(244, 187, 255);
                            guna2DataGridView1.Columns[guna2DataGridView1.Columns.Count - 3].HeaderCell.Style.BackColor = Color.FromArgb(241, 167, 254);

                            int existingColumnIndexToCheck = 0;
                            int newColumnIndex = guna2DataGridView1.Columns.Count - 3;

                            foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                            {
                                bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                if (isRowDisabled)
                                {
                                    row.Cells[newColumnIndex].ReadOnly = true;
                                }
                            }
                            comboBox11.Text = "";
                            comboBox14.Text = "";
                            comboBox11.Focus();
                            guna2RadioButton1.Checked = false;
                            isDone = false;
                        }
                        else
                        {
                            MessageBox.Show("දැනටමත් අයිතම අංකය පවතිනවා", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("වර්ණය වලංගු නොවේ", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && comboBox11.Text != "Test" && comboBox11.Text != "Sample" && isDone)
                {
                    bool columnNotExists1 = true;
                    string item = comboBox11.Text;
                    string color = comboBox14.Text;
                    if (string.IsNullOrEmpty(color))
                    {
                        color = "White";
                    }
                    string enteredColor = color;
                    string nickname = GetNicknameForColor(enteredColor);

                    foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                    {
                        if (column.HeaderText == $"{item} {nickname}")
                        {
                            if (color == colors[column.Index - 1])
                            {
                                columnNotExists1 = false;
                                break;
                            }
                        }
                    }
                    if (columnNotExists1)
                    {
                        bool sameColumnContain = false;
                        bool NotcontainsRepairColumn = true;

                        foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                        {
                            if (column.Name == "repair")
                            {
                                NotcontainsRepairColumn = false;
                                break;
                            }
                        }
                        connection.Open();
                        SqlCommand command = new SqlCommand(query, connection);
                        command.Parameters.AddWithValue("@item", item);
                        // Execute the SQL query
                        SqlDataReader reader = command.ExecuteReader();

                        if (reader.Read())
                        {
                            string category = reader["category"].ToString();
                            int index = categories.IndexOf(category);
                            bool IsnotDone = true;
                            foreach (string name in columnNames)
                            {
                                if (category == name)
                                {
                                    sameColumnContain = true;
                                }
                            }

                            if (sameColumnContain)
                            {
                                int i = 0;
                                foreach (string name in columnNames)
                                {
                                    if (category == name)
                                    {
                                        sameColumnContain = true;
                                        break;
                                    }
                                    i++;
                                }
                                productIds.Insert(i - 1, item);
                                colors.Insert(i - 1, color);
                                DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                newColumn.HeaderText = $"{item} {nickname}";
                                newColumn.Name = category;
                                string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                columnNames.Insert(i, newColumn.Name);
                                indexOfNames.Add(index);
                                guna2DataGridView1.Columns.Insert(i, newColumn);
                                newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                int existingColumnIndexToCheck = 0;
                                int newColumnIndex = i;

                                foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                {
                                    bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                    if (isRowDisabled)
                                    {
                                        row.Cells[newColumnIndex].ReadOnly = true;
                                    }
                                }
                                comboBox11.Text = "";
                                comboBox14.Text = "";
                                comboBox11.Focus();
                                IsnotDone = false;
                            }
                            if (IsnotDone && indexOfNames.Count != 0 && !sameColumnContain)
                            {
                                int result = 0;
                                int maxNumber = indexOfNames.Max();

                                if (index > maxNumber)
                                {
                                    int r = 0;
                                    string categoryOfLook = categories[maxNumber];
                                    foreach (string column in columnNames)
                                    {
                                        if (categoryOfLook == column)
                                        {
                                            result = r;
                                        }
                                        r++;
                                    }
                                    productIds.Insert(result < productIds.Count ? result : productIds.Count, item);
                                    colors.Insert(result < colors.Count ? result : colors.Count, color);

                                    DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                    newColumn.HeaderText = $"{item} {nickname}";
                                    newColumn.Name = category.ToString();
                                    string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                    newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                    columnNames.Insert(result + 1, newColumn.Name);
                                    indexOfNames.Add(index);
                                    guna2DataGridView1.Columns.Insert(result + 1, newColumn);
                                    newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                    newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                    int existingColumnIndexToCheck = 0;
                                    int newColumnIndex = result + 1;

                                    foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                    {
                                        bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                        if (isRowDisabled)
                                        {
                                            row.Cells[newColumnIndex].ReadOnly = true;
                                        }
                                    }
                                    comboBox11.Text = "";
                                    comboBox14.Text = "";
                                    comboBox11.Focus();
                                    IsnotDone = false;

                                }
                                else
                                {
                                    int i = -1;
                                    int result1 = 0;
                                    foreach (int number in indexOfNames)
                                    {
                                        if (index > number)
                                        {
                                            if (number > i)
                                            {
                                                i = number;
                                            }
                                        }
                                    }

                                    if (i != -1)
                                    {
                                        int s = 0;
                                        string categoryOfLook = categories[i];
                                        foreach (string column in columnNames)
                                        {
                                            if (categoryOfLook == column)
                                            {
                                                result1 = s;
                                            }
                                            s++;
                                        }
                                        productIds.Insert(result1 < productIds.Count ? result1 : productIds.Count, item);
                                        colors.Insert(result1 < colors.Count ? result1 : colors.Count, color);
                                        DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                        newColumn.HeaderText = $"{item} {nickname}";
                                        newColumn.Name = category.ToString();
                                        string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                        newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                        columnNames.Insert(result1 + 1, newColumn.Name);
                                        indexOfNames.Add(index);
                                        guna2DataGridView1.Columns.Insert(result1 + 1, newColumn);
                                        newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                        newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                        int existingColumnIndexToCheck = 0;
                                        int newColumnIndex = result1;

                                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                        {
                                            bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                            if (isRowDisabled)
                                            {
                                                row.Cells[newColumnIndex].ReadOnly = true;
                                            }
                                        }
                                        comboBox11.Text = "";
                                        comboBox14.Text = "";
                                        comboBox11.Focus();
                                        IsnotDone = false;
                                    }
                                    else
                                    {
                                        int minNumber = indexOfNames.Min();
                                        string categoryOfLook = categories[minNumber];
                                        foreach (string column in columnNames)
                                        {
                                            if (categoryOfLook == column)
                                            {
                                                result1 = guna2DataGridView1.Columns[categoryOfLook].Index;
                                                break;
                                            }
                                        }
                                        productIds.Insert(result1 - 1, item);
                                        colors.Insert(result1 - 1, color);
                                        DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                        newColumn.HeaderText = $"{item} {nickname}";
                                        newColumn.Name = category.ToString();
                                        string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                        newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                        columnNames.Insert(result1, newColumn.Name);
                                        indexOfNames.Add(index);
                                        guna2DataGridView1.Columns.Insert(result1, newColumn);
                                        newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                        newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                        int existingColumnIndexToCheck = 0;
                                        int newColumnIndex = result1;

                                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                        {
                                            bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                            if (isRowDisabled)
                                            {
                                                row.Cells[newColumnIndex].ReadOnly = true;
                                            }
                                        }
                                        comboBox11.Text = "";
                                        comboBox14.Text = "";
                                        comboBox11.Focus();
                                        IsnotDone = false;
                                    }
                                }
                            }
                            else if (IsnotDone)
                            {
                                bool havingTestColumn = false;
                                foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                                {
                                    if (column.Name == "Test" || column.Name == "Sample")
                                    {
                                        havingTestColumn = true;
                                    }
                                }

                                if (havingTestColumn)
                                {
                                    productIds.Insert(1, item);
                                    colors.Insert(1, color);
                                    DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                    newColumn.HeaderText = $"{item} {nickname}";
                                    newColumn.Name = category;
                                    string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                    newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                    columnNames.Insert(2, newColumn.Name);
                                    indexOfNames.Add(index);
                                    guna2DataGridView1.Columns.Insert(2, newColumn);
                                    newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                    newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                    int existingColumnIndexToCheck = 0;
                                    int newColumnIndex = 2;

                                    foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                    {
                                        bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                        if (isRowDisabled)
                                        {
                                            row.Cells[newColumnIndex].ReadOnly = true;
                                        }
                                    }
                                    comboBox11.Text = "";
                                    comboBox14.Text = "";
                                    comboBox11.Focus();
                                }
                                else
                                {
                                    productIds.Insert(0, item);
                                    colors.Insert(0, color);
                                    DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                    newColumn.HeaderText = $"{item} {nickname}";
                                    newColumn.Name = category;
                                    string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                    newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                    columnNames.Insert(1, newColumn.Name);
                                    indexOfNames.Add(index);
                                    guna2DataGridView1.Columns.Insert(1, newColumn);
                                    newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                    newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                    int existingColumnIndexToCheck = 0;
                                    int newColumnIndex = 1;

                                    foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                    {
                                        bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                        if (isRowDisabled)
                                        {
                                            row.Cells[newColumnIndex].ReadOnly = true;
                                        }
                                    }
                                    comboBox11.Text = "";
                                    comboBox14.Text = "";
                                    comboBox11.Focus();
                                }
                            }
                        }
                        connection.Close();
                    }
                    else
                    {
                        MessageBox.Show("දැනටමත් අයිතම අංකය පවතිනවා", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                else if ((comboBox11.Text == "Test") || (comboBox11.Text == "Sample") && string.IsNullOrEmpty(comboBox14.Text) && !guna2RadioButton1.Checked)
                {
                    bool columnNotExists2 = true;
                    string item = comboBox11.Text;

                    foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                    {
                        if (column.Name == "Test")
                        {
                            columnNotExists2 = false;
                            break;
                        }
                    }
                    if (columnNotExists2)
                    {
                        productIds.Insert(0, item);
                        colors.Insert(0, "");
                        DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                        newColumn.HeaderText = $"{item} {""}";
                        newColumn.Name = "Test";
                        string itemNumber = newColumn.HeaderText.Split(' ')[0];
                        newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                        columnNames.Insert(1, newColumn.Name);
                        guna2DataGridView1.Columns.Insert(1, newColumn);

                        int existingColumnIndexToCheck = 0;
                        int newColumnIndex = 1;

                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                        {
                            bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                            if (isRowDisabled)
                            {
                                row.Cells[newColumnIndex].ReadOnly = true;
                            }
                        }
                        comboBox11.Text = "";
                        comboBox14.Text = "";
                        comboBox11.Focus();
                    }
                    else
                    {
                        MessageBox.Show("දැනටමත් අයිතමය පවතිනවා", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("ක්ෂේත්‍ර නිවැරදිව පුරවන්න", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }*/


        public bool GetAvailableColor(string enteredColor)
        {
            string color = null;
            bool IsColorAvailable = true;
            using (SqlConnection connection = new SqlConnection(Properties.Settings.Default.ConnectionString))
            {
                connection.Open();
                string query1 = "SELECT color FROM colors WHERE color = @EnteredColor";
                using (SqlCommand command1 = new SqlCommand(query1, connection))
                {
                    command1.Parameters.AddWithValue("@EnteredColor", enteredColor);
                    using (SqlDataReader reader1 = command1.ExecuteReader())
                    {
                        if (reader1.Read())
                        {
                            color = reader1["color"].ToString();
                            if (string.IsNullOrEmpty(color))
                            {
                                IsColorAvailable = false;
                            }
                        }
                    }
                }
            }
            return IsColorAvailable;
        }


        public string GetNicknameForColor(string enteredColor)
        {
            string nickname = null;
            using (SqlConnection connection = new SqlConnection(Properties.Settings.Default.ConnectionString))
            {
                connection.Open();
                string query = "SELECT NickName FROM colors WHERE color = @EnteredColor";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@EnteredColor", enteredColor);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            nickname = reader["Nickname"].ToString();
                            if (string.IsNullOrEmpty(nickname))
                            {
                                nickname = enteredColor;
                            }
                        }
                    }
                }
            }
            return nickname;
        }


        private void SetRowCount(int speed)
        {

            int rowCount = speed / 6;
            // Ensure the DataTable has the required number of rows
            while (table.Rows.Count < rowCount)
            {

                table.Rows.Add(table.NewRow());
            }

        }


        // Save data to Database
        private void guna2DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            bool isSaved = true;
            int checkindex = guna2DataGridView1.Columns["check"].Index;
            if (e.RowIndex >= 0 && e.ColumnIndex == checkindex)
            {
                if (names.Count == 0)
                {
                    MessageBox.Show("සාමාජිකයින්ගේ නම් ඇතුළත් කර නැත", "දෝෂය", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                    {
                        if (row.Index < e.RowIndex)
                        {
                            if (row.Cells["Kiln Car"].ReadOnly == false)
                            {
                                isSaved = false;
                                break;

                            }
                        }
                    }
                    if (isSaved)
                    {
                        int RowIndex = e.RowIndex;
                        int listSize = names.Count;
                        int indexOftheName = RowIndex % listSize;
                        DataGridViewCheckBoxCell ch1 = new DataGridViewCheckBoxCell();
                        ch1 = (DataGridViewCheckBoxCell)guna2DataGridView1.Rows[guna2DataGridView1.CurrentRow.Index].Cells["check"];
                        if (ch1.Value == null)
                            ch1.Value = false;
                        switch (ch1.Value.ToString())
                        {
                            case "True":
                                ch1.Value = false;
                                break;
                            case "False":
                                ch1.Value = true;
                                break;
                        }
                        bool checkOk = (bool)ch1.Value;
                        if (checkOk)
                        {
                            bool numberIsOk = false;
                            int total = 0;
                            int rowIndex = e.RowIndex;
                            int columnIndex = e.ColumnIndex;
                            string columnName = guna2DataGridView1.Columns[e.ColumnIndex].Name;

                            DataGridViewCell carnumberCell = guna2DataGridView1.Rows[rowIndex].Cells[0];

                            // Check if the Kiln Car Number Empty
                            if (string.IsNullOrEmpty(carnumberCell.Value?.ToString()))
                            {
                                MessageBox.Show("කාර් අංකය හිස්", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                guna2DataGridView1.Rows[rowIndex].Cells["check"].Value = null;
                            }
                            else
                            {
                                if (int.TryParse(carnumberCell.Value?.ToString(), out int parsedValue))
                                {
                                    if (parsedValue >= 1 && parsedValue <= 66)
                                    {
                                        numberIsOk = true;
                                    }
                                    else
                                    {
                                        MessageBox.Show("කාර් අංකය වලංගු පරාසයක නැත", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        numberIsOk = false;
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("කාර් අංකය වලංගු නැත", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    numberIsOk = false;
                                }
                            }
                            // Check whether Row is already Disabled
                            DataGridViewCell targetCell = guna2DataGridView1.Rows[rowIndex].Cells[0];
                            if (targetCell.ReadOnly)
                            {
                                IsAccessible = false;
                            }
                            else
                            {
                                IsAccessible = true;
                            }
                            // Check if the cell is in the checkbox column
                            if (IsAccessible && numberIsOk)

                            {
                                List<string> car_number = new List<string>();
                                List<string> color = new List<string>();
                                List<string> item = new List<string>();
                                List<int> quantity = new List<int>();
                                List<string> RepairColor = new List<string>();
                                List<string> RepairItem = new List<string>();
                                List<int> RepairQuantity = new List<int>();
                                bool isok = true;
                                DataGridViewRow row = guna2DataGridView1.Rows[rowIndex];
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    string columnNameRepair = guna2DataGridView1.Columns[cell.ColumnIndex].Name;
                                    if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()) && cell.ColumnIndex != columnIndex)
                                    {
                                        if (int.TryParse(cell.Value.ToString(), out int num1))
                                        {
                                            if (cell.ColumnIndex != 0 && columnNameRepair != "Repair" && cell.ColumnIndex != guna2DataGridView1.ColumnCount - 1)
                                            {
                                                total += num1;
                                                item.Add(productIds[cell.ColumnIndex - 1]);
                                                color.Add(colors[cell.ColumnIndex - 1]);
                                                quantity.Add(num1);
                                            }
                                            else if (columnNameRepair == "Repair")
                                            {
                                                total += num1;
                                                RepairItem.Add(productIds[cell.ColumnIndex - 1]);
                                                RepairColor.Add(colors[cell.ColumnIndex - 1]);
                                                RepairQuantity.Add(num1);
                                            }
                                            else
                                            {
                                                car_number.Add((string)row.Cells[0].Value);
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show($"'{cell.OwningColumn.HeaderText}' තීරු අගය අංකයක් නොවේ", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            isok = false;
                                        }
                                    }
                                }
                                if (isok)
                                {
                                    DateTime currentDateTime = DateTime.Now;
                                    SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);
                                    for (int i = 0; i < item.Count; i++)
                                    {
                                        SqlCommand cmd2 = new SqlCommand("INSERT INTO [dbo].[GKloading]\r\n           ([Kiln_Car_number]\r\n           ,[ProductID]\r\n           ,[Color]\r\n           ,[Quantity]\r\n           ,[DateandTime]\r\n           ,[Loader]\r\n           ,[row]\r\n           ,[Shift]\r\n           ,[speed]\r\n           ,[sheet])     VALUES\r\n           ('" + car_number[0] + "','" + item[i] + "','" + color[i] + "','" + quantity[i] + "','" + currentDateTime + "','" + names[indexOftheName] + "','" + rowIndex + "','" + guna2TextBox2.Text + "','" + speed + "','" + guna2TextBox1.Text + "')", con);
                                        con.Open();
                                        cmd2.ExecuteNonQuery();
                                        con.Close();
                                    }
                                    if (RepairItem.Count != 0)
                                    {
                                        for (int i = 0; i < RepairItem.Count; i++)
                                        {
                                            SqlCommand cmd3 = new SqlCommand("INSERT INTO [dbo].[GKloading]\r\n           ([Kiln_Car_number]\r\n           ,[ProductID]\r\n           ,[Color]\r\n           ,[Quantity]\r\n           ,[DateandTime]\r\n           ,[Loader]\r\n           ,[Type]\r\n           ,[row]\r\n           ,[Shift]\r\n           ,[speed]\r\n           ,[sheet])     VALUES\r\n           ('" + car_number[0] + "','" + RepairItem[i] + "','" + RepairColor[i] + "','" + RepairQuantity[i] + "','" + currentDateTime + "','" + names[indexOftheName] + "','" + "R" + "','" + rowIndex + "','" + guna2TextBox2.Text + "','" + speed + "','" + guna2TextBox1.Text + "')", con);
                                            con.Open();
                                            cmd3.ExecuteNonQuery();
                                            con.Close();
                                        }
                                    }
                                    guna2DataGridView1.Rows[e.RowIndex].Cells["Total"].Value = total;
                                    total = 0;
                                    foreach (DataGridViewCell cell in guna2DataGridView1.Rows[rowIndex].Cells)
                                    {
                                        cell.ReadOnly = true;
                                        Font italicFont = new Font(cell.InheritedStyle.Font, FontStyle.Italic);
                                        cell.Style.Font = italicFont;
                                        cell.Style.ForeColor = Color.FromArgb(80, 78, 72, 1);
                                    }
                                    item.Clear();
                                    color.Clear();
                                    quantity.Clear();
                                    car_number.Clear();
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("කරුණාකර පෙර කාර් දත්ත මුලින් අවසන් කරන්න", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        // Check whether the Kiln_Car_number is Empty or not
        private void guna2DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string columnName = guna2DataGridView1.Columns[e.ColumnIndex].Name;
            if (e.ColumnIndex != 0 && e.RowIndex >= 0 && columnName != "check")
            {
                object firstCellValue = guna2DataGridView1.Rows[e.RowIndex].Cells[0].Value;
                if (firstCellValue != null && int.TryParse(firstCellValue.ToString(), out _))
                { }
                else
                {
                    MessageBox.Show("කාර් අංකය හිස් හෝ වලංගු නොවේ", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    guna2DataGridView1.CurrentCell = guna2DataGridView1.Rows[e.RowIndex].Cells[0];
                    guna2DataGridView1.BeginEdit(true);
                }
            }

            /*int rowCount = e.RowIndex + 1;
            // Assuming you want to make the cells in the next row read-only after 20 rows
            if (rowCount >= 5 && rowCount < guna2DataGridView1.Rows.Count)
            {
                Console.WriteLine(rowCount);
                DataGridViewRow nextRow = guna2DataGridView1.Rows[rowCount];
                foreach (DataGridViewCell cell in nextRow.Cells)
                {
                    cell.ReadOnly = true;
                }
            }*/
        }

        // Validate close button
        private void guna2ControlBox1_Click(object sender, EventArgs e)
        {
            bool IsUnsavedData = false;
            bool dataGridViewExists = Controls.OfType<DataGridView>().Any();

            if (dataGridViewExists)
            {
                foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                {
                    // Assuming your DataGridView has at least one column
                    DataGridViewCell cell = row.Cells[0]; // Access the first column cell

                    if (cell.Value != null)
                    {
                        bool isCellReadOnly = cell.ReadOnly;
                        if (!isCellReadOnly)
                        {
                            IsUnsavedData = true;
                            break;
                        }
                    }
                }
                if (IsUnsavedData)
                {
                    MessageBox.Show("ගබඩා නොකළ දත්ත පවතී", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    DialogResult Result = MessageBox.Show("ඔබට වසා දැමීමට අවශ්‍ය බව විශ්වාසද?", "වසා දැමීම", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (Result == DialogResult.OK)
                    {
                        this.Close();
                    }
                }
            }
            else
            {
                DialogResult Result = MessageBox.Show("ඔබට වසා දැමීමට අවශ්‍ය බව විශ්වාසද?", "වසා දැමීම", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (Result == DialogResult.OK)
                {
                    this.Close();
                }
            }
        }


        // Deleting the Existing DataGridView 
        private void guna2Button7_Click(object sender, EventArgs e)
        {
            bool IsUnsavedData = false;

            if (guna2DataGridView1.RowCount > 0)
            {
                foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                {
                    // Assuming your DataGridView has at least one column
                    DataGridViewCell cell = row.Cells[0]; // Access the first column cell
                    bool IsSaved = cell.ReadOnly;
                    if (cell.Value != null && IsSaved == false)
                    {
                        IsUnsavedData = true;
                        break;
                    }
                }
                if (IsUnsavedData)
                {
                    MessageBox.Show("ගබඩා නොකළ දත්ත පවතී", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    DialogResult Result = MessageBox.Show("ඔබට දත්ත වගුව මැකීමට අවශ්‍ය බව විශ්වාසද?", "මකා දමන්න", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (Result == DialogResult.OK)
                    {
                        guna2DataGridView1.DataSource = null;
                        guna2DataGridView1.Rows.Clear();
                        guna2DataGridView1.Columns.Clear();
                        guna2DataGridView2.DataSource = null;
                        guna2DataGridView2.Rows.Clear();
                        guna2DataGridView2.Columns.Clear();
                        productIds.Clear();
                        columnNames.Clear();
                        indexOfNames.Clear();
                        colors.Clear();
                        names.Clear();
                        guna2TextBox1.Text = "";
                        guna2TextBox2.Text = "";
                        guna2TextBox2.Visible = false;
                        guna2TextBox1.Visible = false;
                        speed = 0;
                        label5.Visible = true;
                    }
                }
            }
            else
            {
                MessageBox.Show("දත්ත වගුවක් නොමැත", "මකා දමන්න", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }


        // Resetting the Table
        private void guna2Button1_Click(object sender, EventArgs e)
        {
            List<string> removingNames = new List<string>();
            int lastReadOnlyRow = -1;
            bool IsUnsavedData = false;

            if (guna2DataGridView1.RowCount > 0)
            {
                foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                {
                    // Assuming your DataGridView has at least one column
                    DataGridViewCell cell = row.Cells[0]; // Access the first column cell
                    bool IsSaved = cell.ReadOnly;
                    if (cell.Value != null && IsSaved == false)
                    {
                        IsUnsavedData = true;
                        break;
                    }
                }
                if (IsUnsavedData)
                {
                    MessageBox.Show("ගබඩා නොකළ දත්ත පවතී", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    DialogResult Result = MessageBox.Show("ඔබට දත්ත වගුව මැකීමට අවශ්‍ය බව විශ්වාසද?", "මකා දමන්න", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (Result == DialogResult.OK)
                    {

                        for (int i = guna2DataGridView1.Rows.Count - 1; i >= 0; i--)
                        {
                            DataGridViewCell firstCell = guna2DataGridView1.Rows[i].Cells[0];

                            if (firstCell.ReadOnly)
                            {
                                lastReadOnlyRow = i;
                                break; // Exit the loop once the last read-only row is found
                            }
                        }

                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                        {
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                cell.Value = null;
                                cell.ReadOnly = false;
                                Font boldFont = new Font(cell.InheritedStyle.Font, FontStyle.Bold);
                                cell.Style.Font = boldFont;
                                cell.Style.ForeColor = System.Drawing.Color.Black;
                            }
                        }
                        // Assuming dataGridView is the name of your DataGridView control

                        if (lastReadOnlyRow != -1)
                        {
                            int listSize = names.Count;
                            int indexOftheName = lastReadOnlyRow % listSize;
                            if (indexOftheName < listSize - 1)
                            {
                                for (int i = 0; i <= indexOftheName; i++)
                                {
                                    removingNames.Add(names[i]);
                                }
                                names.RemoveRange(0, removingNames.Count);
                                names.AddRange(removingNames);
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("දත්ත වගුවක් නොමැත", "මකා දමන්න", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }


        private int CalculateWidthBasedOnItemNumber(string itemNumber)
        {
            int calculatedWidth = 53;
            return calculatedWidth;
        }



        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            /*
            if (e.KeyCode == Keys.Delete)
            {
                // Check if a column header is selected
                if (guna2DataGridView1.SelectedCells.Count > 0 && guna2DataGridView1.SelectedCells[0].OwningColumn != null)
                {
                    DataGridViewColumn selectedColumn = guna2DataGridView1.SelectedCells[0].OwningColumn;

                    // Remove the selected column
                    guna2DataGridView1.Columns.Remove(selectedColumn);
                }
            }
            */
        }

        private void guna2RadioButton1_MouseClick(object sender, MouseEventArgs e)
        {
            // Toggle the Checked state
            guna2RadioButton1.Checked = !guna2RadioButton1.Checked;
        }

        private void comboBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
            }

            if (e.KeyChar == (char)Keys.Space)
            {
                comboBox14.Focus();
                e.Handled = true;
            }
        }


        ////////////////////////////////////////////////////
        private void guna2DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > 0 && e.ColumnIndex < guna2DataGridView1.Columns.Count - 2)
            {
                UpdateTotalForRow(e.RowIndex);
            }
        }
        private void UpdateTotalForRow(int rowIndex)
        {
            int total = 0;
            DataGridViewRow row = guna2DataGridView1.Rows[rowIndex];
            foreach (DataGridViewCell cell in row.Cells)
            {
                if (cell.ColumnIndex != 0 && cell.ColumnIndex != guna2DataGridView1.Columns.Count - 1 && cell.ColumnIndex != guna2DataGridView1.Columns.Count - 2)
                {
                    int cellValue = ParseCellValue(cell.Value);
                    total += cellValue;
                }
            }

            // Update the total column for the specified row
            guna2DataGridView1.Rows[rowIndex].Cells["Total"].Style.Font = new Font("Yu Gothic UI", 9, FontStyle.Regular);
            guna2DataGridView1.Rows[rowIndex].Cells["Total"].Value = total;
        }

        private int ParseCellValue(object value)
        {
            int parsedValue;
            return int.TryParse(value?.ToString(), out parsedValue) ? parsedValue : 0;
        }


        private void guna2DataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);
            if (e.RowIndex >= 0)
            {
                DataGridViewCell clickedCell = guna2DataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (clickedCell.ReadOnly)
                {
                    // Clicked on a column header cell
                    string columnName = guna2DataGridView1.Columns[e.ColumnIndex].Name;
                    if (columnName == "Kiln Car")
                    {
                        List<string> listofDate1 = new List<string>();
                        string kilnCarnumber1 = (string)guna2DataGridView1.Rows[e.RowIndex].Cells[0].Value;
                        string kilnCarnumberOld = (string)guna2DataGridView1.Rows[e.RowIndex].Cells[0].Value;
                        object cellValue = clickedCell.Value;
                        DialogResult result = ShowcarUpdate(ref kilnCarnumber1, 300, 200);
                        if (result == DialogResult.OK)
                        {
                            guna2DataGridView1.Rows[e.RowIndex].Cells[0].Value = kilnCarnumber1;
                            guna2DataGridView1.Rows[e.RowIndex].Cells["check"].Value = false;
                            guna2DataGridView1.Rows[e.RowIndex].ReadOnly = false;
                            DataGridViewRow row = guna2DataGridView1.Rows[e.RowIndex];
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                if (cell.OwningColumn.Name == "Total")
                                {
                                    cell.Style.Font = new Font("Yu Gothic UI", 9, FontStyle.Regular);
                                }
                                else
                                {
                                    cell.Style.Font = new Font("Yu Gothic UI", 10, FontStyle.Bold);
                                    cell.Style.ForeColor = Color.Black;
                                }
                            }
                            string query = "SELECT TOP 1 DateandTime FROM GKloading WHERE Kiln_Car_number = @KilnCarNumber ORDER BY DateandTime DESC";
                            using (SqlConnection connection = new SqlConnection(Properties.Settings.Default.ConnectionString))
                            {
                                SqlCommand command = new SqlCommand(query, connection);
                                command.Parameters.AddWithValue("@KilnCarNumber", kilnCarnumberOld);
                                connection.Open();
                                SqlDataReader reader = command.ExecuteReader();
                                while (reader.Read())
                                {
                                    DateTime date = (DateTime)reader["DateandTime"];
                                    listofDate1.Add(date.ToString());
                                }
                                reader.Close();
                                connection.Close();
                                string deleteQuery = "delete from GKloading where Kiln_Car_number = @KilnCarNumber and DateandTime = @Date1";
                                SqlCommand command1 = new SqlCommand(deleteQuery, connection);
                                command1.Parameters.AddWithValue("@KilnCarNumber", kilnCarnumberOld);
                                command1.Parameters.AddWithValue("@Date1", listofDate1[0]);
                                connection.Open();
                                int rowsAffected = command1.ExecuteNonQuery();
                                connection.Close();
                            }
                        }
                    }

                    else if (columnName != "Kiln Car" && columnName != "check" && columnName != "Total")
                    {
                        List<string> listofDate = new List<string>();
                        string item = productIds[e.ColumnIndex - 1];
                        string color = colors[e.ColumnIndex - 1];
                        string kilnCarnumber = (string)guna2DataGridView1.Rows[e.RowIndex].Cells[0].Value;
                        string quantity = (string)guna2DataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                        object cellValue = clickedCell.Value;
                        DialogResult result = ShowUpdate(ref item, ref kilnCarnumber, ref color, ref quantity, 300, 200);
                        if (result == DialogResult.OK)
                        {
                            guna2DataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = quantity;
                            guna2DataGridView1.Rows[e.RowIndex].Cells["check"].Value = false;
                            guna2DataGridView1.Rows[e.RowIndex].ReadOnly = false;
                            DataGridViewRow row = guna2DataGridView1.Rows[e.RowIndex];
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                if (cell.OwningColumn.Name == "Total")
                                {
                                    cell.Style.Font = new Font("Yu Gothic UI", 9, FontStyle.Regular);
                                }
                                else
                                {
                                    cell.Style.Font = new Font("Yu Gothic UI", 10, FontStyle.Bold);
                                    cell.Style.ForeColor = Color.Black;
                                }
                            }
                            string query = "SELECT TOP 1 DateandTime FROM GKloading WHERE Kiln_Car_number = @KilnCarNumber ORDER BY DateandTime DESC";
                            using (SqlConnection connection = new SqlConnection(Properties.Settings.Default.ConnectionString))
                            {
                                SqlCommand command = new SqlCommand(query, connection);
                                command.Parameters.AddWithValue("@KilnCarNumber", kilnCarnumber);
                                connection.Open();
                                SqlDataReader reader = command.ExecuteReader();
                                while (reader.Read())
                                {
                                    DateTime date = (DateTime)reader["DateandTime"];
                                    listofDate.Add(date.ToString());
                                }
                                reader.Close();
                                connection.Close();
                                string deleteQuery = "delete from GKloading where Kiln_Car_number = @KilnCarNumber and DateandTime = @Date1";
                                SqlCommand command1 = new SqlCommand(deleteQuery, connection);
                                command1.Parameters.AddWithValue("@KilnCarNumber", kilnCarnumber);
                                command1.Parameters.AddWithValue("@Date1", listofDate[0]);
                                connection.Open();
                                int rowsAffected = command1.ExecuteNonQuery();
                                connection.Close();
                            }
                        }
                    }
                }
            }
        }
        private static DialogResult ShowUpdate(ref string item, ref string kilnCarnumber, ref string color, ref string quantity, int width = 300, int height = 200)
        {
            Size size = new Size(width, height);
            Form inputBox = new Form();
            inputBox.StartPosition = FormStartPosition.CenterScreen;
            inputBox.MaximizeBox = false;
            inputBox.FormBorderStyle = FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "යාවත්කාලීන කිරීම";

            Label label = new Label();
            label.Text = $"කාර් අංකය: {kilnCarnumber}";
            label.Location = new Point(5, 20);
            label.Width = size.Width - 10;
            inputBox.Controls.Add(label);

            Label label1 = new Label();
            label1.Text = $"අයිතමය: {item} {color}";
            label1.Location = new Point(5, label.Location.Y + 20);
            label1.Width = size.Width - 10;
            inputBox.Controls.Add(label1);

            Label label2 = new Label();
            label2.Text = "ප්‍රමාණය:";
            label2.Location = new Point(5, label1.Location.Y + 30);
            label2.Width = size.Width - 10;
            inputBox.Controls.Add(label2);

            TextBox textBox = new TextBox();
            textBox.Size = new Size(size.Width - 80, 23);
            textBox.Location = new Point(6, label2.Location.Y + 20);
            textBox.Text = quantity;
            inputBox.Controls.Add(textBox);

            Button okButton = new Button();
            okButton.DialogResult = DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new Point(size.Width - 80 - 80, size.Height - 30);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new Point(size.Width - 80, size.Height - 30);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;
            DialogResult result = inputBox.ShowDialog();
            quantity = textBox.Text;
            return result;
        }

        private static DialogResult ShowcarUpdate(ref string kilnCarnumber1, int width = 300, int height = 200)
        {
            Size size = new Size(width, height);
            Form inputBox = new Form();
            inputBox.StartPosition = FormStartPosition.CenterScreen;
            inputBox.MaximizeBox = false;
            inputBox.FormBorderStyle = FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "යාවත්කාලීන කිරීම";

            Label label = new Label();
            label.Text = $"කාර් අංකය: {kilnCarnumber1}";
            label.Location = new Point(5, 20);
            label.Width = size.Width - 10;
            inputBox.Controls.Add(label);

            TextBox textBox = new TextBox();
            textBox.Size = new Size(size.Width - 80, 23);
            textBox.Location = new Point(6, label.Location.Y + 20);
            textBox.Text = kilnCarnumber1;
            inputBox.Controls.Add(textBox);

            Button okButton = new Button();
            okButton.DialogResult = DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new Point(size.Width - 80 - 80, size.Height - 30);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new Point(size.Width - 80, size.Height - 30);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;
            DialogResult result = inputBox.ShowDialog();
            kilnCarnumber1 = textBox.Text;
            return result;
        }


        public class RowData
        {
            public int RowNumber { get; set; }
            public object Value { get; set; }
        }
        private void guna2DataGridView1_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            List<RowData> rowSavedDataList = new List<RowData>();
            List<RowData> rowInsertDataList = new List<RowData>();
            SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);

            // Check if the target column index is within a valid range
            if (e.ColumnIndex > 0 && e.ColumnIndex < guna2DataGridView1.Columns.Count - 1)
            {
                int targetColumnIndex = e.ColumnIndex - 1; // Adjusted the target index

                if (targetColumnIndex >= 0 && targetColumnIndex < productIds.Count && targetColumnIndex < colors.Count)
                {
                    string item = productIds[targetColumnIndex];
                    string color = colors[targetColumnIndex];
                    string itemOld = productIds[targetColumnIndex];
                    string colorOld = colors[targetColumnIndex];
                    string columnName = guna2DataGridView1.Columns[targetColumnIndex + 1].Name;

                    DialogResult result = Changecolumn(ref item, ref color, 300, 200);
                    if (result == DialogResult.OK)
                    {
                        // Store the data from the old column (value, isReadOnly)
                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                        {
                            bool notDone = true;
                            if (row.Cells[0].ReadOnly && row.Cells[targetColumnIndex + 1].Value != null && !string.IsNullOrEmpty(row.Cells[targetColumnIndex + 1].Value.ToString()))
                            {
                                rowSavedDataList.Add(new RowData { RowNumber = row.Index, Value = row.Cells[targetColumnIndex + 1].Value.ToString() });
                                notDone = false;
                            }
                            else if (row.Cells[targetColumnIndex + 1].Value != null && !string.IsNullOrEmpty(row.Cells[targetColumnIndex + 1].Value.ToString()) && notDone)
                            {
                                rowInsertDataList.Add(new RowData { RowNumber = row.Index, Value = row.Cells[targetColumnIndex + 1].Value.ToString() });

                            }
                        }

                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                        {
                            if (row.Cells[0].ReadOnly)
                            {
                                // Check if the target column index is valid
                                if (row.Cells[targetColumnIndex + 1].Value != null && !string.IsNullOrEmpty(row.Cells[targetColumnIndex + 1].Value.ToString()))
                                {
                                    if (columnName == "Repair")
                                    {
                                        string car_number = row.Cells[0].Value.ToString();
                                        string query = $"update GKloading set ProductID = @item, Color = @color where Kiln_Car_number = @KilnCarNumber and ProductID = @Item0 and Color = @Color0 AND Type = 'R' AND DateandTime = (SELECT MAX(DateandTime) FROM GKloading WHERE kiln_car_number = @KilnCarNumber)";
                                        SqlCommand cmd0 = new SqlCommand(query, con);
                                        cmd0.Parameters.AddWithValue("@item", item);
                                        cmd0.Parameters.AddWithValue("@color", color);
                                        cmd0.Parameters.AddWithValue("@KilnCarNumber", car_number);
                                        cmd0.Parameters.AddWithValue("@Item0", itemOld);
                                        cmd0.Parameters.AddWithValue("@Color0", colorOld);
                                        try
                                        {
                                            con.Open();
                                            int rowsAffected = cmd0.ExecuteNonQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                        con.Close();
                                    }
                                    else
                                    {
                                        string car_number = row.Cells[0].Value.ToString();
                                        string query = $"update GKloading set ProductID = @item, Color = @color where Kiln_Car_number = @KilnCarNumber and ProductID = @Item0 and Color = @Color0 AND DateandTime = (SELECT MAX(DateandTime) FROM GKloading WHERE kiln_car_number = @KilnCarNumber)";
                                        SqlCommand cmd0 = new SqlCommand(query, con);
                                        cmd0.Parameters.AddWithValue("@item", item);
                                        cmd0.Parameters.AddWithValue("@color", color);
                                        cmd0.Parameters.AddWithValue("@KilnCarNumber", car_number);
                                        cmd0.Parameters.AddWithValue("@Item0", itemOld);
                                        cmd0.Parameters.AddWithValue("@Color0", colorOld);
                                        try
                                        {
                                            con.Open();
                                            int rowsAffected = cmd0.ExecuteNonQuery();
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                        con.Close();
                                    }
                                }
                            }
                        }

                        productIds.RemoveAt(targetColumnIndex);
                        colors.RemoveAt(targetColumnIndex);
                        productIds.Insert(targetColumnIndex, item);
                        colors.Insert(targetColumnIndex, color);

                        DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                        if (columnName == "Repair")
                        {
                            string headerText = $"{item} {color} (R)";
                            newColumn.HeaderText = headerText;
                            newColumn.Width = 53;
                        }
                        else
                        {
                            string headerText = $"{item} {color}";
                            newColumn.HeaderText = headerText;
                            newColumn.Width = 53;
                        }

                        newColumn.Name = columnName;
                        guna2DataGridView1.Columns.RemoveAt(targetColumnIndex + 1);
                        guna2DataGridView1.Columns.Insert(targetColumnIndex + 1, newColumn);
                        if (columnName == "Repair")
                        {
                            newColumn.DefaultCellStyle.BackColor = Color.FromArgb(204, 207, 255);
                            newColumn.HeaderCell.Style.BackColor = System.Drawing.Color.MediumPurple;
                        }
                        else
                        {
                            newColumn.DefaultCellStyle.BackColor = colorCategories[columnName];
                            newColumn.HeaderCell.Style.BackColor = colorCategories[columnName];
                        }
                        // Populate the new column with the stored data
                        foreach (var rowData in rowSavedDataList)
                        {
                            guna2DataGridView1.Rows[rowData.RowNumber].Cells[targetColumnIndex + 1].Value = rowData.Value;
                            guna2DataGridView1.Rows[rowData.RowNumber].Cells[targetColumnIndex + 1].ReadOnly = true;
                            Font italicFont = new Font(guna2DataGridView1.Rows[rowData.RowNumber].Cells[targetColumnIndex + 1].InheritedStyle.Font, FontStyle.Italic);
                            guna2DataGridView1.Rows[rowData.RowNumber].Cells[targetColumnIndex + 1].Style.Font = italicFont;
                            guna2DataGridView1.Rows[rowData.RowNumber].Cells[targetColumnIndex + 1].Style.ForeColor = Color.FromArgb(80, 78, 72, 1);

                        }
                        foreach (var rowData in rowInsertDataList)
                        {
                            guna2DataGridView1.Rows[rowData.RowNumber].Cells[targetColumnIndex + 1].Value = rowData.Value;
                        }
                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                        {
                            if (row.Cells[0].ReadOnly)
                            {
                                row.Cells[targetColumnIndex + 1].ReadOnly = true;
                            }
                        }
                    }
                }
            }
        }


        private static DialogResult Changecolumn(ref string item, ref string color, int width = 300, int height = 200)
        {
            SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);
            string query2 = "SELECT color FROM colors";
            string query3 = "SELECT items FROM item";

            Size size = new Size(width, height);
            Form inputBox = new Form();
            inputBox.StartPosition = FormStartPosition.CenterScreen;
            inputBox.MaximizeBox = false;
            inputBox.FormBorderStyle = FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "යාවත්කාලීන කිරීම";

            Label label1 = new Label();
            label1.Text = $"අයිතමය:";
            label1.Location = new Point(5, 20);
            label1.Width = size.Width - 10;
            inputBox.Controls.Add(label1);

            ComboBox comboBox1 = new ComboBox();
            comboBox1.Text = $"{item}";
            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.MaxDropDownItems = 5;
            comboBox1.Location = new Point(5, label1.Location.Y + 20);
            comboBox1.Width = size.Width - 10;
            inputBox.Controls.Add(comboBox1);

            Label label2 = new Label();
            label2.Text = $"වර්ණය:";
            label2.Location = new Point(5, comboBox1.Location.Y + 30);
            label2.Width = size.Width - 10;
            inputBox.Controls.Add(label2);

            ComboBox comboBox = new ComboBox();
            comboBox.Text = $"{color}";
            comboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox.MaxDropDownItems = 5;
            comboBox.Location = new Point(5, label2.Location.Y + 20);
            comboBox.Width = size.Width - 10;
            inputBox.Controls.Add(comboBox);

            //get colors to the combo box from the colors table 
            SqlCommand cmd2 = new SqlCommand(query2, con);
            con.Open();
            SqlDataReader reader = cmd2.ExecuteReader();
            while (reader.Read())
            {
                string data = reader.GetString(0);
                comboBox.Items.Add(data);
            }
            con.Close();

            //get item numbers to the combo box from the item table  
            SqlCommand cmd3 = new SqlCommand(query3, con);
            con.Open();
            SqlDataReader reader1 = cmd3.ExecuteReader();
            while (reader1.Read())
            {
                string data1 = reader1.GetString(0);
                comboBox1.Items.Add(data1);
            }
            con.Close();

            Button okButton = new Button();
            okButton.DialogResult = DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new Point(size.Width - 80 - 80, size.Height - 30);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new Point(size.Width - 80, size.Height - 30);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;
            DialogResult result = inputBox.ShowDialog();
            item = comboBox1.Text;
            color = comboBox.Text;
            return result;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (panel1.Visible)
            {
                panel1.Hide();
            }
            else
            {
                panel1.Show();
            }
        }

        private void Panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = true;
                offset = new Point(e.X, e.Y);
            }
        }

        private void Panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Point newLocation = this.PointToScreen(new Point(e.X, e.Y));
                newLocation.Offset(-offset.X, -offset.Y);
                this.Location = newLocation;
            }
        }

        private void Panel1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = false;
            }
        }

        private void guna2DataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            lastInteractionTime = DateTime.Now;
        }

        private void CheckIdleTime()
        {
            // Set your idle time threshold (e.g., 5 minutes)
            TimeSpan idleThreshold = TimeSpan.FromSeconds(30);

            // Check if the user is idle
            if (DateTime.Now - lastInteractionTime > idleThreshold)
            {
                // Select the first cell in the current row
                if (guna2DataGridView1.CurrentRow != null)
                {
                    guna2DataGridView1.Rows[guna2DataGridView1.CurrentRow.Index].Cells[0].Selected = true;
                }
            }
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
            {
                if (column.Name == "Repair")
                {
                    column.Visible = false;
                }
                else
                {
                    column.Visible = true;
                }
            }
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
            {
                column.Visible = true;
            }
        }

        private void guna2Button8_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
            {
                if (column.Name != "Repair" && column.Name != "check" && column.Name != "Total" && column.Name != "Kiln Car")
                {
                    column.Visible = false;
                }
                else
                {
                    column.Visible = true;
                }
            }
        }

        private void toolStripContainer1_ContentPanel_Load(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            CheckIdleTime();
        }

        private void guna2Button3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Space)
            {
                e.Handled = true;
                List<string> categories = new List<string>();

                categories.AddRange(itemsToAdd);

                string query = "SELECT category FROM item WHERE items = @item";
                SqlConnection connection = new SqlConnection(Properties.Settings.Default.ConnectionString);
                bool Isok = true;
                bool isDone = true;
                bool columnnotExists = true;
                foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                {
                    if (column.HeaderText == "Kiln Car")
                    {
                        columnnotExists = false;
                        break;
                    }
                }
                if (columnnotExists)
                {
                    if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && guna2RadioButton1.Checked)
                    {
                        string item = comboBox11.Text;
                        string color = comboBox14.Text;
                        if (string.IsNullOrEmpty(color))
                        {
                            color = "White";
                        }
                        string enteredColor = color;
                        string nickname = GetNicknameForColor(enteredColor);
                        productIds.Add(item);
                        colors.Add(color);
                        comboBox11.Text = "";
                        comboBox14.Text = "";
                        guna2RadioButton1.Checked = false;

                        textColumn1.HeaderText = "Kiln Car";
                        textColumn1.Name = "Kiln Car";
                        textColumn1.Width = 35;
                        columnNames.Add(textColumn1.Name);
                        guna2DataGridView1.Columns.Add(textColumn1);
                        int i = 0;
                        foreach (string Item in productIds)
                        {
                            DataGridViewTextBoxColumn Text = new DataGridViewTextBoxColumn();
                            Text.HeaderText = $"{Item} {nickname} (R)";
                            Text.Name = "Repair";
                            string itemNumber = Text.HeaderText.Split(' ')[0];
                            Text.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                            guna2DataGridView1.Columns.Add(Text);
                            columnNames.Add(Text.Name);
                            i++;
                        }
                        check.Name = "check";
                        check.ValueType = typeof(bool);
                        check.HeaderText = "Finished";
                        guna2DataGridView1.Columns.Add(check);
                        textColumn.Name = "Total";
                        textColumn.Width = 60;
                        textColumn.HeaderText = "Total";
                        columnNames.Add(check.Name);
                        columnNames.Add(textColumn.Name);
                        guna2DataGridView1.Columns.Add(textColumn);
                        guna2DataGridView1.DataSource = table;
                        SetRowCount(speed);
                        guna2DataGridView1.AllowUserToAddRows = false;
                        guna2DataGridView1.Columns[0].Frozen = true;
                        guna2DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                        System.Drawing.Color desiredColor = System.Drawing.Color.PeachPuff;
                        guna2DataGridView1.Columns[0].DefaultCellStyle.BackColor = desiredColor;
                        guna2DataGridView1.Columns[0].HeaderCell.Style.BackColor = System.Drawing.Color.PaleVioletRed;
                        guna2DataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.FromArgb(244, 187, 255);
                        guna2DataGridView1.Columns[1].HeaderCell.Style.BackColor = Color.FromArgb(241, 167, 254);
                        guna2DataGridView1.AlternatingRowsDefaultCellStyle = null;
                        comboBox11.Focus();
                        Isok = false;

                    }
                    else if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && comboBox11.Text != "Test" && comboBox11.Text != "Sample" && Isok)
                    {
                        bool Isok2 = false;
                        string item = comboBox11.Text;
                        string color = comboBox14.Text;
                        if (string.IsNullOrEmpty(color))
                        {
                            color = "White";
                        }
                        string enteredColor = color;
                        string nickname = GetNicknameForColor(enteredColor);

                        connection.Open();
                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@item", item);

                            // Execute the SQL query
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    string category = reader["category"].ToString();
                                    int index = categories.IndexOf(category);
                                    productIds.Add(item);
                                    colors.Add(color);
                                    textColumn1.HeaderText = "Kiln Car";
                                    textColumn1.Name = "Kiln Car";
                                    textColumn1.Width = 35;
                                    columnNames.Add(textColumn1.Name);
                                    guna2DataGridView1.Columns.Add(textColumn1);
                                    int i = 0;
                                    foreach (string Item in productIds)
                                    {
                                        DataGridViewTextBoxColumn Text = new DataGridViewTextBoxColumn();
                                        Text.HeaderText = $"{Item} {nickname}";
                                        Text.Name = category;
                                        string itemNumber = Text.HeaderText.Split(' ')[0];
                                        Text.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                        columnNames.Add(Text.Name);
                                        indexOfNames.Add(index);
                                        guna2DataGridView1.Columns.Add(Text);
                                        Text.DefaultCellStyle.BackColor = colorCategories[category];
                                        Text.HeaderCell.Style.BackColor = colorCategories[category];
                                        i++;
                                    }
                                    check.Name = "check";
                                    check.ValueType = typeof(bool);
                                    check.HeaderText = "Finished";
                                    guna2DataGridView1.Columns.Add(check);
                                    textColumn.Name = "Total";
                                    textColumn.HeaderText = "Total";
                                    textColumn.Width = 60;
                                    columnNames.Add(check.Name);
                                    columnNames.Add(textColumn.Name);
                                    guna2DataGridView1.Columns.Add(textColumn);
                                    guna2DataGridView1.DataSource = table;
                                    SetRowCount(speed);

                                    guna2DataGridView1.AllowUserToAddRows = false;
                                    guna2DataGridView1.Columns[0].Frozen = true;
                                    guna2DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                                    System.Drawing.Color desiredColor = System.Drawing.Color.PeachPuff;
                                    guna2DataGridView1.Columns[0].DefaultCellStyle.BackColor = desiredColor;
                                    guna2DataGridView1.Columns[0].HeaderCell.Style.BackColor = System.Drawing.Color.PaleVioletRed;
                                    guna2DataGridView1.AlternatingRowsDefaultCellStyle = null;
                                    comboBox11.Text = "";
                                    comboBox14.Text = "";
                                    comboBox11.Focus();
                                }
                                else
                                {
                                    Isok2 = false;
                                }
                            }
                        }
                        connection.Close();
                    }
                    else if ((comboBox11.Text == "Test" || comboBox11.Text == "Sample") && string.IsNullOrEmpty(comboBox14.Text) && !guna2RadioButton1.Checked)
                    {
                        string item = comboBox11.Text;
                        productIds.Add(item);
                        colors.Add("");
                        comboBox11.Text = "";
                        comboBox14.Text = "";

                        textColumn1.HeaderText = "Kiln Car";
                        textColumn1.Name = "Kiln Car";
                        textColumn1.Width = 35;
                        columnNames.Add(textColumn1.Name);
                        guna2DataGridView1.Columns.Add(textColumn1);
                        int i = 0;
                        foreach (string Item in productIds)
                        {
                            DataGridViewTextBoxColumn Text = new DataGridViewTextBoxColumn();
                            Text.HeaderText = $"{Item} {colors[i]}";
                            Text.Name = "Test";
                            string itemNumber = Text.HeaderText.Split(' ')[0];
                            Text.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                            guna2DataGridView1.Columns.Add(Text);
                            columnNames.Add(Text.Name);
                            i++;
                        }
                        check.Name = "check";
                        check.ValueType = typeof(bool);
                        check.HeaderText = "Finished";
                        guna2DataGridView1.Columns.Add(check);
                        textColumn.Name = "Total";
                        textColumn.HeaderText = "Total";
                        textColumn.Width = 60;
                        columnNames.Add(check.Name);
                        columnNames.Add(textColumn.Name);
                        guna2DataGridView1.Columns.Add(textColumn);
                        guna2DataGridView1.DataSource = table;
                        SetRowCount(speed);
                        guna2DataGridView1.AllowUserToAddRows = false;
                        guna2DataGridView1.Columns[0].Frozen = true;
                        guna2DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                        System.Drawing.Color desiredColor = System.Drawing.Color.PeachPuff;
                        guna2DataGridView1.Columns[0].DefaultCellStyle.BackColor = desiredColor;
                        guna2DataGridView1.Columns[0].HeaderCell.Style.BackColor = System.Drawing.Color.PaleVioletRed;
                        guna2DataGridView1.AlternatingRowsDefaultCellStyle = null;
                        comboBox11.Focus();
                    }
                    else
                    {
                        MessageBox.Show("ක්ෂේත්‍ර නිවැරදිව පුරවන්න", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && guna2RadioButton1.Checked)
                    {
                        bool columnNotExists = true;
                        string item = comboBox11.Text;
                        string color = comboBox14.Text;
                        if (string.IsNullOrEmpty(color))
                        {
                            color = "White";
                        }
                        string enteredColor = color;
                        bool isColorAvailable = GetAvailableColor(enteredColor);
                        if (isColorAvailable)
                        {
                            string nickname = GetNicknameForColor(enteredColor);
                            foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                            {
                                if (column.HeaderText == $"{item} {nickname} (R)")
                                {
                                    if (color == colors[column.Index - 1])
                                    {
                                        columnNotExists = false;
                                        break;
                                    }
                                }
                            }
                            if (columnNotExists)
                            {
                                productIds.Add(item);
                                colors.Add(color);

                                DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                newColumn.HeaderText = $"{item} {nickname} (R)";
                                newColumn.Name = "Repair";
                                string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                columnNames.Insert(guna2DataGridView1.Columns.Count - 2, newColumn.Name);

                                int insertionIndex = guna2DataGridView1.Columns.Count - 2;
                                guna2DataGridView1.Columns.Insert(insertionIndex, newColumn);
                                guna2DataGridView1.Columns[guna2DataGridView1.Columns.Count - 3].DefaultCellStyle.BackColor = Color.FromArgb(244, 187, 255);
                                guna2DataGridView1.Columns[guna2DataGridView1.Columns.Count - 3].HeaderCell.Style.BackColor = Color.FromArgb(241, 167, 254);

                                int existingColumnIndexToCheck = 0;
                                int newColumnIndex = guna2DataGridView1.Columns.Count - 3;

                                foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                {
                                    bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                    if (isRowDisabled)
                                    {
                                        row.Cells[newColumnIndex].ReadOnly = true;
                                    }
                                }
                                comboBox11.Text = "";
                                comboBox14.Text = "";
                                comboBox11.Focus();
                                guna2RadioButton1.Checked = false;
                                isDone = false;
                            }
                            else
                            {
                                MessageBox.Show("දැනටමත් අයිතම අංකය පවතිනවා", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("වර්ණය වලංගු නොවේ", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && comboBox11.Text != "Test" && comboBox11.Text != "Sample" && isDone)
                    {
                        bool columnNotExists1 = true;
                        string item = comboBox11.Text;
                        string color = comboBox14.Text;
                        if (string.IsNullOrEmpty(color))
                        {
                            color = "White";
                        }
                        string enteredColor = color;
                        string nickname = GetNicknameForColor(enteredColor);

                        foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                        {
                            if (column.HeaderText == $"{item} {nickname}")
                            {
                                if (color == colors[column.Index - 1])
                                {
                                    columnNotExists1 = false;
                                    break;
                                }
                            }
                        }
                        if (columnNotExists1)
                        {
                            bool sameColumnContain = false;
                            bool NotcontainsRepairColumn = true;

                            foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                            {
                                if (column.Name == "repair")
                                {
                                    NotcontainsRepairColumn = false;
                                    break;
                                }
                            }
                            connection.Open();
                            SqlCommand command = new SqlCommand(query, connection);
                            command.Parameters.AddWithValue("@item", item);
                            // Execute the SQL query
                            SqlDataReader reader = command.ExecuteReader();

                            if (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int index = categories.IndexOf(category);
                                bool IsnotDone = true;
                                foreach (string name in columnNames)
                                {
                                    if (category == name)
                                    {
                                        sameColumnContain = true;
                                    }
                                }

                                if (sameColumnContain)
                                {
                                    int i = 0;
                                    foreach (string name in columnNames)
                                    {
                                        if (category == name)
                                        {
                                            sameColumnContain = true;
                                            break;
                                        }
                                        i++;
                                    }
                                    productIds.Insert(i - 1, item);
                                    colors.Insert(i - 1, color);
                                    DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                    newColumn.HeaderText = $"{item} {nickname}";
                                    newColumn.Name = category;
                                    string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                    newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                    columnNames.Insert(i, newColumn.Name);
                                    indexOfNames.Add(index);
                                    guna2DataGridView1.Columns.Insert(i, newColumn);
                                    newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                    newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                    int existingColumnIndexToCheck = 0;
                                    int newColumnIndex = i;

                                    foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                    {
                                        bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                        if (isRowDisabled)
                                        {
                                            row.Cells[newColumnIndex].ReadOnly = true;
                                        }
                                    }
                                    comboBox11.Text = "";
                                    comboBox14.Text = "";
                                    comboBox11.Focus();
                                    IsnotDone = false;
                                }
                                if (IsnotDone && indexOfNames.Count != 0 && !sameColumnContain)
                                {
                                    int result = 0;
                                    int maxNumber = indexOfNames.Max();

                                    if (index > maxNumber)
                                    {
                                        int r = 0;
                                        string categoryOfLook = categories[maxNumber];
                                        foreach (string column in columnNames)
                                        {
                                            if (categoryOfLook == column)
                                            {
                                                result = r;
                                            }
                                            r++;
                                        }
                                        productIds.Insert(result < productIds.Count ? result : productIds.Count, item);
                                        colors.Insert(result < colors.Count ? result : colors.Count, color);

                                        DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                        newColumn.HeaderText = $"{item} {nickname}";
                                        newColumn.Name = category.ToString();
                                        string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                        newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                        columnNames.Insert(result + 1, newColumn.Name);
                                        indexOfNames.Add(index);
                                        guna2DataGridView1.Columns.Insert(result + 1, newColumn);
                                        newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                        newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                        int existingColumnIndexToCheck = 0;
                                        int newColumnIndex = result + 1;

                                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                        {
                                            bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                            if (isRowDisabled)
                                            {
                                                row.Cells[newColumnIndex].ReadOnly = true;
                                            }
                                        }
                                        comboBox11.Text = "";
                                        comboBox14.Text = "";
                                        comboBox11.Focus();
                                        IsnotDone = false;

                                    }
                                    else
                                    {
                                        int i = -1;
                                        int result1 = 0;
                                        foreach (int number in indexOfNames)
                                        {
                                            if (index > number)
                                            {
                                                if (number > i)
                                                {
                                                    i = number;
                                                }
                                            }
                                        }

                                        if (i != -1)
                                        {
                                            int s = 0;
                                            string categoryOfLook = categories[i];
                                            foreach (string column in columnNames)
                                            {
                                                if (categoryOfLook == column)
                                                {
                                                    result1 = s;
                                                }
                                                s++;
                                            }
                                            productIds.Insert(result1 < productIds.Count ? result1 : productIds.Count, item);
                                            colors.Insert(result1 < colors.Count ? result1 : colors.Count, color);
                                            DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                            newColumn.HeaderText = $"{item} {nickname}";
                                            newColumn.Name = category.ToString();
                                            string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                            newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                            columnNames.Insert(result1 + 1, newColumn.Name);
                                            indexOfNames.Add(index);
                                            guna2DataGridView1.Columns.Insert(result1 + 1, newColumn);
                                            newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                            newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                            int existingColumnIndexToCheck = 0;
                                            int newColumnIndex = result1;

                                            foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                            {
                                                bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                                if (isRowDisabled)
                                                {
                                                    row.Cells[newColumnIndex].ReadOnly = true;
                                                }
                                            }
                                            comboBox11.Text = "";
                                            comboBox14.Text = "";
                                            comboBox11.Focus();
                                            IsnotDone = false;
                                        }
                                        else
                                        {
                                            int minNumber = indexOfNames.Min();
                                            string categoryOfLook = categories[minNumber];
                                            foreach (string column in columnNames)
                                            {
                                                if (categoryOfLook == column)
                                                {
                                                    result1 = guna2DataGridView1.Columns[categoryOfLook].Index;
                                                    break;
                                                }
                                            }
                                            productIds.Insert(result1 - 1, item);
                                            colors.Insert(result1 - 1, color);
                                            DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                            newColumn.HeaderText = $"{item} {nickname}";
                                            newColumn.Name = category.ToString();
                                            string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                            newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                            columnNames.Insert(result1, newColumn.Name);
                                            indexOfNames.Add(index);
                                            guna2DataGridView1.Columns.Insert(result1, newColumn);
                                            newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                            newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                            int existingColumnIndexToCheck = 0;
                                            int newColumnIndex = result1;

                                            foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                            {
                                                bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                                if (isRowDisabled)
                                                {
                                                    row.Cells[newColumnIndex].ReadOnly = true;
                                                }
                                            }
                                            comboBox11.Text = "";
                                            comboBox14.Text = "";
                                            comboBox11.Focus();
                                            IsnotDone = false;
                                        }
                                    }
                                }
                                else if (IsnotDone)
                                {
                                    bool havingTestColumn = false;
                                    foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                                    {
                                        if (column.Name == "Test" || column.Name == "Sample")
                                        {
                                            havingTestColumn = true;
                                        }
                                    }

                                    if (havingTestColumn)
                                    {
                                        productIds.Insert(1, item);
                                        colors.Insert(1, color);
                                        DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                        newColumn.HeaderText = $"{item} {nickname}";
                                        newColumn.Name = category;
                                        string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                        newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                        columnNames.Insert(2, newColumn.Name);
                                        indexOfNames.Add(index);
                                        guna2DataGridView1.Columns.Insert(2, newColumn);
                                        newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                        newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                        int existingColumnIndexToCheck = 0;
                                        int newColumnIndex = 2;

                                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                        {
                                            bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                            if (isRowDisabled)
                                            {
                                                row.Cells[newColumnIndex].ReadOnly = true;
                                            }
                                        }
                                        comboBox11.Text = "";
                                        comboBox14.Text = "";
                                        comboBox11.Focus();
                                    }
                                    else
                                    {
                                        productIds.Insert(0, item);
                                        colors.Insert(0, color);
                                        DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                        newColumn.HeaderText = $"{item} {nickname}";
                                        newColumn.Name = category;
                                        string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                        newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                        columnNames.Insert(1, newColumn.Name);
                                        indexOfNames.Add(index);
                                        guna2DataGridView1.Columns.Insert(1, newColumn);
                                        newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                        newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                        int existingColumnIndexToCheck = 0;
                                        int newColumnIndex = 1;

                                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                        {
                                            bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                            if (isRowDisabled)
                                            {
                                                row.Cells[newColumnIndex].ReadOnly = true;
                                            }
                                        }
                                        comboBox11.Text = "";
                                        comboBox14.Text = "";
                                        comboBox11.Focus();
                                    }
                                }
                            }
                            connection.Close();
                        }
                        else
                        {
                            MessageBox.Show("දැනටමත් අයිතම අංකය පවතිනවා", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    else if ((comboBox11.Text == "Test") || (comboBox11.Text == "Sample") && string.IsNullOrEmpty(comboBox14.Text) && !guna2RadioButton1.Checked)
                    {
                        bool columnNotExists2 = true;
                        string item = comboBox11.Text;

                        foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                        {
                            if (column.Name == "Test")
                            {
                                columnNotExists2 = false;
                                break;
                            }
                        }
                        if (columnNotExists2)
                        {
                            productIds.Insert(0, item);
                            colors.Insert(0, "");
                            DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                            newColumn.HeaderText = $"{item} {""}";
                            newColumn.Name = "Test";
                            string itemNumber = newColumn.HeaderText.Split(' ')[0];
                            newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                            columnNames.Insert(1, newColumn.Name);
                            guna2DataGridView1.Columns.Insert(1, newColumn);

                            int existingColumnIndexToCheck = 0;
                            int newColumnIndex = 1;

                            foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                            {
                                bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                if (isRowDisabled)
                                {
                                    row.Cells[newColumnIndex].ReadOnly = true;
                                }
                            }
                            comboBox11.Text = "";
                            comboBox14.Text = "";
                            comboBox11.Focus();
                        }
                        else
                        {
                            MessageBox.Show("දැනටමත් අයිතමය පවතිනවා", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("ක්ෂේත්‍ර නිවැරදිව පුරවන්න", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void comboBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Space)
            {
                guna2Button3.Focus();
                e.Handled = true;
            }
        }

        private void guna2Button3_MouseClick(object sender, MouseEventArgs e)
        {
            List<string> categories = new List<string>();

            categories.AddRange(itemsToAdd);

            string query = "SELECT category FROM item WHERE items = @item";
            SqlConnection connection = new SqlConnection(Properties.Settings.Default.ConnectionString);
            bool Isok = true;
            bool isDone = true;
            bool columnnotExists = true;
            foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
            {
                if (column.HeaderText == "Kiln Car")
                {
                    columnnotExists = false;
                    break;
                }
            }
            if (columnnotExists)
            {
                if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && guna2RadioButton1.Checked)
                {
                    string item = comboBox11.Text;
                    string color = comboBox14.Text;
                    if (string.IsNullOrEmpty(color))
                    {
                        color = "White";
                    }
                    string enteredColor = color;
                    string nickname = GetNicknameForColor(enteredColor);
                    productIds.Add(item);
                    colors.Add(color);
                    comboBox11.Text = "";
                    comboBox14.Text = "";
                    guna2RadioButton1.Checked = false;

                    textColumn1.HeaderText = "Kiln Car";
                    textColumn1.Name = "Kiln Car";
                    textColumn1.Width = 35;
                    columnNames.Add(textColumn1.Name);
                    guna2DataGridView1.Columns.Add(textColumn1);
                    int i = 0;
                    foreach (string Item in productIds)
                    {
                        DataGridViewTextBoxColumn Text = new DataGridViewTextBoxColumn();
                        Text.HeaderText = $"{Item} {nickname} (R)";
                        Text.Name = "Repair";
                        string itemNumber = Text.HeaderText.Split(' ')[0];
                        Text.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                        guna2DataGridView1.Columns.Add(Text);
                        columnNames.Add(Text.Name);
                        i++;
                    }
                    check.Name = "check";
                    check.ValueType = typeof(bool);
                    check.HeaderText = "Finished";
                    guna2DataGridView1.Columns.Add(check);
                    textColumn.Name = "Total";
                    textColumn.Width = 60;
                    textColumn.HeaderText = "Total";
                    columnNames.Add(check.Name);
                    columnNames.Add(textColumn.Name);
                    guna2DataGridView1.Columns.Add(textColumn);
                    guna2DataGridView1.DataSource = table;
                    SetRowCount(speed);
                    guna2DataGridView1.AllowUserToAddRows = false;
                    guna2DataGridView1.Columns[0].Frozen = true;
                    guna2DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                    System.Drawing.Color desiredColor = System.Drawing.Color.PeachPuff;
                    guna2DataGridView1.Columns[0].DefaultCellStyle.BackColor = desiredColor;
                    guna2DataGridView1.Columns[0].HeaderCell.Style.BackColor = System.Drawing.Color.PaleVioletRed;
                    guna2DataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.FromArgb(244, 187, 255);
                    guna2DataGridView1.Columns[1].HeaderCell.Style.BackColor = Color.FromArgb(241, 167, 254);
                    guna2DataGridView1.AlternatingRowsDefaultCellStyle = null;
                    comboBox11.Focus();
                    Isok = false;

                }
                else if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && comboBox11.Text != "Test" && comboBox11.Text != "Sample" && Isok)
                {
                    bool Isok2 = false;
                    string item = comboBox11.Text;
                    string color = comboBox14.Text;
                    if (string.IsNullOrEmpty(color))
                    {
                        color = "White";
                    }
                    string enteredColor = color;
                    string nickname = GetNicknameForColor(enteredColor);

                    connection.Open();
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@item", item);

                        // Execute the SQL query
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string category = reader["category"].ToString();
                                int index = categories.IndexOf(category);
                                productIds.Add(item);
                                colors.Add(color);
                                textColumn1.HeaderText = "Kiln Car";
                                textColumn1.Name = "Kiln Car";
                                textColumn1.Width = 35;
                                columnNames.Add(textColumn1.Name);
                                guna2DataGridView1.Columns.Add(textColumn1);
                                int i = 0;
                                foreach (string Item in productIds)
                                {
                                    DataGridViewTextBoxColumn Text = new DataGridViewTextBoxColumn();
                                    Text.HeaderText = $"{Item} {nickname}";
                                    Text.Name = category;
                                    string itemNumber = Text.HeaderText.Split(' ')[0];
                                    Text.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                    columnNames.Add(Text.Name);
                                    indexOfNames.Add(index);
                                    guna2DataGridView1.Columns.Add(Text);
                                    Text.DefaultCellStyle.BackColor = colorCategories[category];
                                    Text.HeaderCell.Style.BackColor = colorCategories[category];
                                    i++;
                                }
                                check.Name = "check";
                                check.ValueType = typeof(bool);
                                check.HeaderText = "Finished";
                                guna2DataGridView1.Columns.Add(check);
                                textColumn.Name = "Total";
                                textColumn.HeaderText = "Total";
                                textColumn.Width = 60;
                                columnNames.Add(check.Name);
                                columnNames.Add(textColumn.Name);
                                guna2DataGridView1.Columns.Add(textColumn);
                                guna2DataGridView1.DataSource = table;
                                SetRowCount(speed);

                                guna2DataGridView1.AllowUserToAddRows = false;
                                guna2DataGridView1.Columns[0].Frozen = true;
                                guna2DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                                System.Drawing.Color desiredColor = System.Drawing.Color.PeachPuff;
                                guna2DataGridView1.Columns[0].DefaultCellStyle.BackColor = desiredColor;
                                guna2DataGridView1.Columns[0].HeaderCell.Style.BackColor = System.Drawing.Color.PaleVioletRed;
                                guna2DataGridView1.AlternatingRowsDefaultCellStyle = null;
                                comboBox11.Text = "";
                                comboBox14.Text = "";
                                comboBox11.Focus();
                            }
                            else
                            {
                                Isok2 = false;
                            }
                        }
                    }
                    connection.Close();
                }
                else if ((comboBox11.Text == "Test" || comboBox11.Text == "Sample") && string.IsNullOrEmpty(comboBox14.Text) && !guna2RadioButton1.Checked)
                {
                    string item = comboBox11.Text;
                    productIds.Add(item);
                    colors.Add("");
                    comboBox11.Text = "";
                    comboBox14.Text = "";

                    textColumn1.HeaderText = "Kiln Car";
                    textColumn1.Name = "Kiln Car";
                    textColumn1.Width = 35;
                    columnNames.Add(textColumn1.Name);
                    guna2DataGridView1.Columns.Add(textColumn1);
                    int i = 0;
                    foreach (string Item in productIds)
                    {
                        DataGridViewTextBoxColumn Text = new DataGridViewTextBoxColumn();
                        Text.HeaderText = $"{Item} {colors[i]}";
                        Text.Name = "Test";
                        string itemNumber = Text.HeaderText.Split(' ')[0];
                        Text.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                        guna2DataGridView1.Columns.Add(Text);
                        columnNames.Add(Text.Name);
                        i++;
                    }
                    check.Name = "check";
                    check.ValueType = typeof(bool);
                    check.HeaderText = "Finished";
                    guna2DataGridView1.Columns.Add(check);
                    textColumn.Name = "Total";
                    textColumn.HeaderText = "Total";
                    textColumn.Width = 60;
                    columnNames.Add(check.Name);
                    columnNames.Add(textColumn.Name);
                    guna2DataGridView1.Columns.Add(textColumn);
                    guna2DataGridView1.DataSource = table;
                    SetRowCount(speed);
                    guna2DataGridView1.AllowUserToAddRows = false;
                    guna2DataGridView1.Columns[0].Frozen = true;
                    guna2DataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                    System.Drawing.Color desiredColor = System.Drawing.Color.PeachPuff;
                    guna2DataGridView1.Columns[0].DefaultCellStyle.BackColor = desiredColor;
                    guna2DataGridView1.Columns[0].HeaderCell.Style.BackColor = System.Drawing.Color.PaleVioletRed;
                    guna2DataGridView1.AlternatingRowsDefaultCellStyle = null;
                    comboBox11.Focus();
                }
                else
                {
                    MessageBox.Show("ක්ෂේත්‍ර නිවැරදිව පුරවන්න", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && guna2RadioButton1.Checked)
                {
                    bool columnNotExists = true;
                    string item = comboBox11.Text;
                    string color = comboBox14.Text;
                    if (string.IsNullOrEmpty(color))
                    {
                        color = "White";
                    }
                    string enteredColor = color;
                    bool isColorAvailable = GetAvailableColor(enteredColor);
                    if (isColorAvailable)
                    {
                        string nickname = GetNicknameForColor(enteredColor);
                        foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                        {
                            if (column.HeaderText == $"{item} {nickname} (R)")
                            {
                                if (color == colors[column.Index - 1])
                                {
                                    columnNotExists = false;
                                    break;
                                }
                            }
                        }
                        if (columnNotExists)
                        {
                            productIds.Add(item);
                            colors.Add(color);

                            DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                            newColumn.HeaderText = $"{item} {nickname} (R)";
                            newColumn.Name = "Repair";
                            string itemNumber = newColumn.HeaderText.Split(' ')[0];
                            newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                            columnNames.Insert(guna2DataGridView1.Columns.Count - 2, newColumn.Name);

                            int insertionIndex = guna2DataGridView1.Columns.Count - 2;
                            guna2DataGridView1.Columns.Insert(insertionIndex, newColumn);
                            guna2DataGridView1.Columns[guna2DataGridView1.Columns.Count - 3].DefaultCellStyle.BackColor = Color.FromArgb(244, 187, 255);
                            guna2DataGridView1.Columns[guna2DataGridView1.Columns.Count - 3].HeaderCell.Style.BackColor = Color.FromArgb(241, 167, 254);

                            int existingColumnIndexToCheck = 0;
                            int newColumnIndex = guna2DataGridView1.Columns.Count - 3;

                            foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                            {
                                bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                if (isRowDisabled)
                                {
                                    row.Cells[newColumnIndex].ReadOnly = true;
                                }
                            }
                            comboBox11.Text = "";
                            comboBox14.Text = "";
                            comboBox11.Focus();
                            guna2RadioButton1.Checked = false;
                            isDone = false;
                        }
                        else
                        {
                            MessageBox.Show("දැනටමත් අයිතම අංකය පවතිනවා", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("වර්ණය වලංගු නොවේ", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (!string.IsNullOrEmpty(comboBox11.Text) && (!string.IsNullOrEmpty(comboBox14.Text) || string.IsNullOrEmpty(comboBox14.Text)) && comboBox11.Text != "Test" && comboBox11.Text != "Sample" && isDone)
                {
                    bool columnNotExists1 = true;
                    string item = comboBox11.Text;
                    string color = comboBox14.Text;
                    if (string.IsNullOrEmpty(color))
                    {
                        color = "White";
                    }
                    string enteredColor = color;
                    string nickname = GetNicknameForColor(enteredColor);

                    foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                    {
                        if (column.HeaderText == $"{item} {nickname}")
                        {
                            if (color == colors[column.Index - 1])
                            {
                                columnNotExists1 = false;
                                break;
                            }
                        }
                    }
                    if (columnNotExists1)
                    {
                        bool sameColumnContain = false;
                        bool NotcontainsRepairColumn = true;

                        foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                        {
                            if (column.Name == "repair")
                            {
                                NotcontainsRepairColumn = false;
                                break;
                            }
                        }
                        connection.Open();
                        SqlCommand command = new SqlCommand(query, connection);
                        command.Parameters.AddWithValue("@item", item);
                        // Execute the SQL query
                        SqlDataReader reader = command.ExecuteReader();

                        if (reader.Read())
                        {
                            string category = reader["category"].ToString();
                            int index = categories.IndexOf(category);
                            bool IsnotDone = true;
                            foreach (string name in columnNames)
                            {
                                if (category == name)
                                {
                                    sameColumnContain = true;
                                }
                            }

                            if (sameColumnContain)
                            {
                                int i = 0;
                                foreach (string name in columnNames)
                                {
                                    if (category == name)
                                    {
                                        sameColumnContain = true;
                                        break;
                                    }
                                    i++;
                                }
                                productIds.Insert(i - 1, item);
                                colors.Insert(i - 1, color);
                                DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                newColumn.HeaderText = $"{item} {nickname}";
                                newColumn.Name = category;
                                string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                columnNames.Insert(i, newColumn.Name);
                                indexOfNames.Add(index);
                                guna2DataGridView1.Columns.Insert(i, newColumn);
                                newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                int existingColumnIndexToCheck = 0;
                                int newColumnIndex = i;

                                foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                {
                                    bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                    if (isRowDisabled)
                                    {
                                        row.Cells[newColumnIndex].ReadOnly = true;
                                    }
                                }
                                comboBox11.Text = "";
                                comboBox14.Text = "";
                                comboBox11.Focus();
                                IsnotDone = false;
                            }
                            if (IsnotDone && indexOfNames.Count != 0 && !sameColumnContain)
                            {
                                int result = 0;
                                int maxNumber = indexOfNames.Max();

                                if (index > maxNumber)
                                {
                                    int r = 0;
                                    string categoryOfLook = categories[maxNumber];
                                    foreach (string column in columnNames)
                                    {
                                        if (categoryOfLook == column)
                                        {
                                            result = r;
                                        }
                                        r++;
                                    }
                                    productIds.Insert(result < productIds.Count ? result : productIds.Count, item);
                                    colors.Insert(result < colors.Count ? result : colors.Count, color);

                                    DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                    newColumn.HeaderText = $"{item} {nickname}";
                                    newColumn.Name = category.ToString();
                                    string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                    newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                    columnNames.Insert(result + 1, newColumn.Name);
                                    indexOfNames.Add(index);
                                    guna2DataGridView1.Columns.Insert(result + 1, newColumn);
                                    newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                    newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                    int existingColumnIndexToCheck = 0;
                                    int newColumnIndex = result + 1;

                                    foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                    {
                                        bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                        if (isRowDisabled)
                                        {
                                            row.Cells[newColumnIndex].ReadOnly = true;
                                        }
                                    }
                                    comboBox11.Text = "";
                                    comboBox14.Text = "";
                                    comboBox11.Focus();
                                    IsnotDone = false;

                                }
                                else
                                {
                                    int i = -1;
                                    int result1 = 0;
                                    foreach (int number in indexOfNames)
                                    {
                                        if (index > number)
                                        {
                                            if (number > i)
                                            {
                                                i = number;
                                            }
                                        }
                                    }

                                    if (i != -1)
                                    {
                                        int s = 0;
                                        string categoryOfLook = categories[i];
                                        foreach (string column in columnNames)
                                        {
                                            if (categoryOfLook == column)
                                            {
                                                result1 = s;
                                            }
                                            s++;
                                        }
                                        productIds.Insert(result1 < productIds.Count ? result1 : productIds.Count, item);
                                        colors.Insert(result1 < colors.Count ? result1 : colors.Count, color);
                                        DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                        newColumn.HeaderText = $"{item} {nickname}";
                                        newColumn.Name = category.ToString();
                                        string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                        newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                        columnNames.Insert(result1 + 1, newColumn.Name);
                                        indexOfNames.Add(index);
                                        guna2DataGridView1.Columns.Insert(result1 + 1, newColumn);
                                        newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                        newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                        int existingColumnIndexToCheck = 0;
                                        int newColumnIndex = result1;

                                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                        {
                                            bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                            if (isRowDisabled)
                                            {
                                                row.Cells[newColumnIndex].ReadOnly = true;
                                            }
                                        }
                                        comboBox11.Text = "";
                                        comboBox14.Text = "";
                                        comboBox11.Focus();
                                        IsnotDone = false;
                                    }
                                    else
                                    {
                                        int minNumber = indexOfNames.Min();
                                        string categoryOfLook = categories[minNumber];
                                        foreach (string column in columnNames)
                                        {
                                            if (categoryOfLook == column)
                                            {
                                                result1 = guna2DataGridView1.Columns[categoryOfLook].Index;
                                                break;
                                            }
                                        }
                                        productIds.Insert(result1 - 1, item);
                                        colors.Insert(result1 - 1, color);
                                        DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                        newColumn.HeaderText = $"{item} {nickname}";
                                        newColumn.Name = category.ToString();
                                        string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                        newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                        columnNames.Insert(result1, newColumn.Name);
                                        indexOfNames.Add(index);
                                        guna2DataGridView1.Columns.Insert(result1, newColumn);
                                        newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                        newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                        int existingColumnIndexToCheck = 0;
                                        int newColumnIndex = result1;

                                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                        {
                                            bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                            if (isRowDisabled)
                                            {
                                                row.Cells[newColumnIndex].ReadOnly = true;
                                            }
                                        }
                                        comboBox11.Text = "";
                                        comboBox14.Text = "";
                                        comboBox11.Focus();
                                        IsnotDone = false;
                                    }
                                }
                            }
                            else if (IsnotDone)
                            {
                                bool havingTestColumn = false;
                                foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                                {
                                    if (column.Name == "Test" || column.Name == "Sample")
                                    {
                                        havingTestColumn = true;
                                    }
                                }

                                if (havingTestColumn)
                                {
                                    productIds.Insert(1, item);
                                    colors.Insert(1, color);
                                    DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                    newColumn.HeaderText = $"{item} {nickname}";
                                    newColumn.Name = category;
                                    string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                    newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                    columnNames.Insert(2, newColumn.Name);
                                    indexOfNames.Add(index);
                                    guna2DataGridView1.Columns.Insert(2, newColumn);
                                    newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                    newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                    int existingColumnIndexToCheck = 0;
                                    int newColumnIndex = 2;

                                    foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                    {
                                        bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                        if (isRowDisabled)
                                        {
                                            row.Cells[newColumnIndex].ReadOnly = true;
                                        }
                                    }
                                    comboBox11.Text = "";
                                    comboBox14.Text = "";
                                    comboBox11.Focus();
                                }
                                else
                                {
                                    productIds.Insert(0, item);
                                    colors.Insert(0, color);
                                    DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                                    newColumn.HeaderText = $"{item} {nickname}";
                                    newColumn.Name = category;
                                    string itemNumber = newColumn.HeaderText.Split(' ')[0];
                                    newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                                    columnNames.Insert(1, newColumn.Name);
                                    indexOfNames.Add(index);
                                    guna2DataGridView1.Columns.Insert(1, newColumn);
                                    newColumn.DefaultCellStyle.BackColor = colorCategories[category];
                                    newColumn.HeaderCell.Style.BackColor = colorCategories[category];

                                    int existingColumnIndexToCheck = 0;
                                    int newColumnIndex = 1;

                                    foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                                    {
                                        bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                                        if (isRowDisabled)
                                        {
                                            row.Cells[newColumnIndex].ReadOnly = true;
                                        }
                                    }
                                    comboBox11.Text = "";
                                    comboBox14.Text = "";
                                    comboBox11.Focus();
                                }
                            }
                        }
                        connection.Close();
                    }
                    else
                    {
                        MessageBox.Show("දැනටමත් අයිතම අංකය පවතිනවා", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                else if ((comboBox11.Text == "Test") || (comboBox11.Text == "Sample") && string.IsNullOrEmpty(comboBox14.Text) && !guna2RadioButton1.Checked)
                {
                    bool columnNotExists2 = true;
                    string item = comboBox11.Text;

                    foreach (DataGridViewColumn column in guna2DataGridView1.Columns)
                    {
                        if (column.Name == "Test")
                        {
                            columnNotExists2 = false;
                            break;
                        }
                    }
                    if (columnNotExists2)
                    {
                        productIds.Insert(0, item);
                        colors.Insert(0, "");
                        DataGridViewTextBoxColumn newColumn = new DataGridViewTextBoxColumn();
                        newColumn.HeaderText = $"{item} {""}";
                        newColumn.Name = "Test";
                        string itemNumber = newColumn.HeaderText.Split(' ')[0];
                        newColumn.Width = CalculateWidthBasedOnItemNumber(itemNumber);
                        columnNames.Insert(1, newColumn.Name);
                        guna2DataGridView1.Columns.Insert(1, newColumn);

                        int existingColumnIndexToCheck = 0;
                        int newColumnIndex = 1;

                        foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                        {
                            bool isRowDisabled = (bool)row.Cells[existingColumnIndexToCheck].ReadOnly;
                            if (isRowDisabled)
                            {
                                row.Cells[newColumnIndex].ReadOnly = true;
                            }
                        }
                        comboBox11.Text = "";
                        comboBox14.Text = "";
                        comboBox11.Focus();
                    }
                    else
                    {
                        MessageBox.Show("දැනටමත් අයිතමය පවතිනවා", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("ක්ෂේත්‍ර නිවැරදිව පුරවන්න", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

            Form2 inputDialog = new Form2();
            inputDialog.speed = speed;

            if (inputDialog.ShowDialog() == DialogResult.OK)
            {
                guna2TextBox1.Visible = true;
                guna2TextBox2.Visible = true;
                DataTable dataTable = new DataTable();
                string shift = inputDialog.shift;
                string sheet = inputDialog.sheet;
                speed = inputDialog.speed;
                foreach (var item in inputDialog.items)
                {
                    names.Add(item.ToString());
                }

                guna2TextBox2.Text = shift;
                guna2TextBox1.Text = sheet;
                dataTable.Columns.Add("names", typeof(string));
                foreach (string Item in names)
                {
                    dataTable.Rows.Add(Item);
                }
                guna2DataGridView2.DataSource = dataTable;
                guna2DataGridView2.AllowUserToAddRows = false;

                label5.Visible = false;
                if (guna2DataGridView1.Rows.Count > 0)
                {
                    SetRowCountonDataGrid(speed);
                }

            }
        }


        private void SetRowCountonDataGrid(int speed)
        {
            int Count = speed / 6;
            // Ensure the DataTable has the required number of rows
            int rowCount = guna2DataGridView1.Rows.Count - 1;
            if (rowCount < Count)
            {
                int i = 0;
                int neededRows = Count - rowCount;
                while (i < neededRows)
                {
                    DataRow newRow = table.NewRow();
                    table.Rows.Add(newRow);
                    i++;
                }
                guna2DataGridView1.Refresh();
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }


        private void DataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            rowIndexFromMouseDown = guna2DataGridView2.HitTest(e.X, e.Y).RowIndex;

            if (rowIndexFromMouseDown != -1)
            {
                Size dragSize = SystemInformation.DragSize;
                movingRow = guna2DataGridView2.Rows[rowIndexFromMouseDown];

                guna2DataGridView2.DoDragDrop(movingRow, DragDropEffects.Move);
            }
        }

        private void DataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.AllowedEffect;
        }

        private void DataGridView1_DragOver(object sender, DragEventArgs e)
        {
            Point clientPoint = guna2DataGridView2.PointToClient(new Point(e.X, e.Y));
            int rowIndex = guna2DataGridView2.HitTest(clientPoint.X, clientPoint.Y).RowIndex;

            if (rowIndexFromMouseDown != rowIndex && rowIndex >= 0 && rowIndex < guna2DataGridView2.Rows.Count)
            {
                guna2DataGridView2.Rows.RemoveAt(rowIndexFromMouseDown);
                guna2DataGridView2.Rows.Insert(rowIndex, movingRow);
            }
        }

        private void DataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            rowIndexFromMouseDown = -1;
        }


        // Updating the Shift
        private void guna2TextBox2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);
            string shift = guna2TextBox2.Text;
            DialogResult result = ShowInput1(ref shift, 300, 200);
            if (result == DialogResult.OK)
            {
                if (!string.IsNullOrEmpty(shift))
                {
                    foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                    {
                        if (row.Cells[0].ReadOnly)
                        {
                            string car_number = row.Cells[0].Value.ToString();
                            string query = $"update GKloading set Shift = @shift where Kiln_Car_number = @KilnCarNumber and DateandTime = (SELECT MAX(DateandTime) FROM GKloading WHERE kiln_car_number = @KilnCarNumber)";
                            SqlCommand cmd0 = new SqlCommand(query, con);
                            cmd0.Parameters.AddWithValue("@shift", shift);
                            cmd0.Parameters.AddWithValue("@KilnCarNumber", car_number);
                            try
                            {
                                con.Open();
                                int rowsAffected = cmd0.ExecuteNonQuery();
                            }
                            catch (Exception ex)
                            {
                            }
                            con.Close();
                        }
                    }
                    guna2TextBox2.Text = shift;
                }
            }
        }
        // Dialog Box to Update Shift
        private static DialogResult ShowInput1(ref string shift, int width = 300, int height = 200)
        {
            Size size = new Size(width, height);
            Form inputBox = new Form();
            inputBox.Location = new Point(0, 0);
            inputBox.MaximizeBox = false;
            inputBox.FormBorderStyle = FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "Details";

            CheckBox checkBox1 = new CheckBox();
            checkBox1.Text = "8 Hour";
            checkBox1.Location = new Point(5, 10);
            inputBox.Controls.Add(checkBox1);

            CheckBox checkBox2 = new CheckBox();
            checkBox2.Text = "12 Hour";
            checkBox2.Location = new Point(115, 10);
            inputBox.Controls.Add(checkBox2);

            Label label = new Label();
            label.Text = "Shift:";
            label.Location = new Point(5, checkBox1.Location.Y + 35);
            label.Width = size.Width - 50;
            inputBox.Controls.Add(label);

            ComboBox comboBox = new ComboBox();
            comboBox.Size = new Size(size.Width - 100, 23);
            comboBox.Location = new Point(5, label.Location.Y + 25);
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            inputBox.Controls.Add(comboBox);

            checkBox1.CheckedChanged += (sender, e) => UpdateComboBox(comboBox, checkBox1, checkBox2);
            checkBox2.CheckedChanged += (sender, e) => UpdateComboBox(comboBox, checkBox1, checkBox2);

            Button okButton = new Button();
            okButton.DialogResult = DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new Point(size.Width - 80 - 80, size.Height - 30);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new Point(size.Width - 80, size.Height - 30);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;
            DialogResult result = inputBox.ShowDialog();
            shift = comboBox.Text;
            return result;
        }
        // Selecting the items for comboBox 
        private static void UpdateComboBox(ComboBox comboBox, CheckBox checkBox1, CheckBox checkBox2)
        {
            comboBox.Items.Clear();

            if (checkBox1.Checked && !checkBox2.Checked)
            {
                comboBox.Items.Add("06:00 - 14:00");
                comboBox.Items.Add("14:00 - 22:00");
                comboBox.Items.Add("22:00 - 06:00");
            }
            else if (checkBox2.Checked && !checkBox1.Checked)
            {
                comboBox.Items.Add("06:00 - 18:00");
                comboBox.Items.Add("18:00 - 06:00");
            }
        }


        // Updating the sheet number
        private void guna2TextBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);
            string sheet = guna2TextBox1.Text;
            DialogResult result = ShowInput2(ref sheet, 300, 200);
            if (result == DialogResult.OK)
            {
                if (!string.IsNullOrEmpty(sheet))
                {
                    foreach (DataGridViewRow row in guna2DataGridView1.Rows)
                    {
                        if (row.Cells[0].ReadOnly)
                        {
                            string car_number = row.Cells[0].Value.ToString();
                            string query = $"update GKloading set sheet = @sheet where Kiln_Car_number = @KilnCarNumber and DateandTime = (SELECT MAX(DateandTime) FROM GKloading WHERE kiln_car_number = @KilnCarNumber)";
                            SqlCommand cmd0 = new SqlCommand(query, con);
                            cmd0.Parameters.AddWithValue("@sheet", sheet);
                            cmd0.Parameters.AddWithValue("@KilnCarNumber", car_number);
                            try
                            {
                                con.Open();
                                int rowsAffected = cmd0.ExecuteNonQuery();
                            }
                            catch (Exception ex)
                            {
                            }
                            con.Close();
                        }
                    }
                    guna2TextBox1.Text = sheet;
                }
            }
        }
        // Dialog Box to Update Sheet
        private static DialogResult ShowInput2(ref string sheet, int width = 300, int height = 200)
        {
            Size size = new Size(width, height);
            Form inputBox = new Form();
            inputBox.Location = new Point(0, 0);
            inputBox.MaximizeBox = false;
            inputBox.FormBorderStyle = FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "Details";

            Label label = new Label();
            label.Text = "Sheet:";
            label.Location = new Point(5, 10);
            label.Width = size.Width - 50;
            inputBox.Controls.Add(label);

            ComboBox comboBox = new ComboBox();
            comboBox.Size = new Size(size.Width - 100, 23);
            comboBox.Location = new Point(5, label.Location.Y + 25);
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox.Items.Add("1");
            comboBox.Items.Add("2");
            comboBox.Items.Add("3");
            inputBox.Controls.Add(comboBox);

            Button okButton = new Button();
            okButton.DialogResult = DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new Point(size.Width - 80 - 80, size.Height - 30);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new Point(size.Width - 80, size.Height - 30);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;
            DialogResult result = inputBox.ShowDialog();
            sheet = comboBox.Text;
            return result;
        }

        private void guna2DataGridView2_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            using (Form3 inputDialog = new Form3())
            {
                inputDialog.SetNamesList(names, speed);
                if (inputDialog.ShowDialog() == DialogResult.OK)
                {
                    names.Clear();
                    DataTable dataTable = new DataTable();
                    speed = inputDialog.speed1;
                    foreach (var item in inputDialog.items)
                    {
                        names.Add(item.ToString());

                    }

                    dataTable.Columns.Add("names", typeof(string));
                    foreach (string Item in names)
                    {
                        dataTable.Rows.Add(Item);
                    }
                    guna2DataGridView2.DataSource = dataTable;
                    guna2DataGridView2.AllowUserToAddRows = false;
                    if (guna2DataGridView1.Rows.Count > 0)
                    {
                        SetRowCountonDataGrid(speed);
                    }
                }
            }
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {

        }
    }
}