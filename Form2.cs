using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace WinFormsApp1
{
    public partial class Form2 : Form
    {
        string query1 = "select top 1 speed from GKloading order by DateandTime desc";
        string query4 = "select Name from employee";
        public string shift { get; private set; }
        public string sheet { get; private set; }
        public int speed { get; set; }

        public List<string> items = new List<string>();
        private object draggedItem;

        
        public Form2()
        {
            InitializeComponent();

            using (SqlConnection connection = new SqlConnection(Properties.Settings.Default.ConnectionString))
            {
                using (SqlCommand command = new SqlCommand(query1, connection))
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
                textBox1.Text = speed.ToString();
            }

            SqlConnection con = new SqlConnection(Properties.Settings.Default.ConnectionString);

            SqlCommand cmd4 = new SqlCommand(query4, con);
            con.Open();
            SqlDataReader reader2 = cmd4.ExecuteReader();
            while (reader2.Read())
            {
                string data2 = reader2.GetString(0);
                comboBox3.Items.Add(data2);
            }
            con.Close();

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            shift = comboBox2.Text;
            sheet = comboBox1.Text;
            int numberOfItems = listBox1.Items.Count;

            if (int.TryParse(textBox1.Text, out int speedValue))
            {
                if (!string.IsNullOrEmpty(shift) && !string.IsNullOrEmpty(sheet) && speedValue >= 100 && speedValue <= 300 && numberOfItems != 0)
                {
                    speed = speedValue;
                    foreach (string item in listBox1.Items)
                    {
                        items.Add(item);
                    }

                    DialogResult = DialogResult.OK;
                    Close();
                }
                else
                {
                    MessageBox.Show("Please fill the Fields correctly", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Enter valid speed", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string name = comboBox3.Text;
            if (!string.IsNullOrEmpty(name))
            {
                if (listBox1.Items.Contains(comboBox3.Text))
                {
                    MessageBox.Show("දත්ත දැනටමත් ඇතුළත් කර ඇත", "වැරදි", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    listBox1.Items.Add(name);
                }
            }
            comboBox3.Text = "";
        }

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
                comboBox3.Items.Add(name);
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

            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
            textBox.Size = new Size(size.Width - 10, 23);
            textBox.Location = new Point(5, label.Location.Y + 20);
            inputBox.Controls.Add(textBox);

            System.Windows.Forms.Button okButton = new System.Windows.Forms.Button();
            okButton.DialogResult = DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new Point(size.Width - 80 - 80, size.Height - 30);
            inputBox.Controls.Add(okButton);

            System.Windows.Forms.Button cancelButton = new System.Windows.Forms.Button();
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

        private void checkBox1_Click(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            comboBox1.Items.Clear();
            if (checkBox1.Checked && !checkBox2.Checked)
            {
                comboBox2.Items.Add("06:00 - 14:00");
                comboBox2.Items.Add("14:00 - 22:00");
                comboBox2.Items.Add("22:00 - 06:00");
                comboBox1.Items.Add("1");
                comboBox1.Items.Add("2");
            }
            else if (checkBox2.Checked && !checkBox1.Checked)
            {
                comboBox2.Items.Add("06:00 - 18:00");
                comboBox2.Items.Add("18:00 - 06:00");
                comboBox1.Items.Add("1");
                comboBox1.Items.Add("2");
                comboBox1.Items.Add("3");
            }
        }

        private void ListBox1_MouseDown(object sender, MouseEventArgs e)
        {
            int index = listBox1.IndexFromPoint(e.X, e.Y);
            if (index != -1)
            {
                listBox1.DoDragDrop(listBox1.Items[index], DragDropEffects.Move);
            }
        }

        private void ListBox1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void ListBox1_DragDrop(object sender, DragEventArgs e)
        {
            Point point = listBox1.PointToClient(new Point(e.X, e.Y));
            int index = listBox1.IndexFromPoint(point);

            if (index != -1)
            {
                object data = e.Data.GetData(typeof(string));
                listBox1.Items.Remove(data);
                listBox1.Items.Insert(index, data);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }
    }
}
