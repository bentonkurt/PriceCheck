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
    public partial class Form8 : Form
    {
        public Form1 frm1; // родительская форма
        public Form8()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) // Кнопка Вывести всю таблицу
        {
            try
            {
                frm1.showAll();
            }
            catch
            {
                MessageBox.Show("Таблица не выведена!", "Сообщение");
            }
        }

        private void button2_Click(object sender, EventArgs e) // Кнопка Вернуть
        {
            textBox1.Text = frm1.returnPageSize().ToString();
        }

        private void button3_Click(object sender, EventArgs e) // Кнопка Установить
        {
            frm1.setupPageSize(Convert.ToInt32(textBox1.Text));
        }

        private void Form8_Load(object sender, EventArgs e)
        {
            this.BackColor = Color.FromArgb(240, 189, 170);
        }
        private void Form8_FormClosed(object sender, FormClosedEventArgs e) // При закрытии формы
        {
            frm1.frm8Loaded = false;
        }
    }
}
