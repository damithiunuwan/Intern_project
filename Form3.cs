using System;
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
    public partial class Form3 : Form
    {
        string query4 = "select Name from employee";
        public int speed1;
        private List<string> namesList = new List<string>();
        public List<string> items = new List<string>();
        public Form3()
        {
            InitializeComponent();

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

        // Method to set the names list in the dialog box
        public void SetNamesList(List<string> names, int speed)
        {
            speed1 = speed;
            namesList = names;
            PopulateListView();
        }

        private void PopulateListView()
        {
            foreach (string name in namesList)
            {
                listBox1.Items.Add(name);
            }
            textBox1.Text = speed1.ToString();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int numberOfItems = listBox1.Items.Count;
            if (int.TryParse(textBox1.Text, out int speedValue))
            {
                if (speedValue >= 100 && speedValue <= 300 && numberOfItems != 0)
                {
                    speed1 = speedValue;
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
