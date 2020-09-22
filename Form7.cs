using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PriceCheck
{
    public partial class Form7 : Form
    {
        public List<string> lst1; // Список колонок для выбора
        public List<string> lst2; // Список колонок для выбора
        public List<string> filteredLst; // Список колонок для выбора
        string str; // Строка условия
        public Form4 pf; // Родительская форма
        public Form7()
        {
            InitializeComponent();
        }

        private void Form7_Load(object sender, EventArgs e)
        {
            comboBox1.DataSource = lst1;
            comboBox2.DataSource = lst2;
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            int idx = lst1.IndexOf(comboBox1.SelectedValue.ToString());

            filteredLst = (from string str in lst2 where lst2.IndexOf(str) >= idx select str).ToList();
            comboBox2.DataSource = filteredLst;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            str = " WHERE [Sheet6].[N] >= " + comboBox1.SelectedValue.ToString() + " AND [Sheet6].[N] <= " + comboBox2.SelectedValue.ToString();
            pf.dobavka = str;
        }

    }
}
