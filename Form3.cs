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
    public partial class Form3 : Form
    {
        public Form1 frm1; // родительская форма
        public string filterString; // Строка фильтра
        public string sortString; // Строка сортировки
        public List<string> colForHide = new List<string>(); // Список колонок для скрытия
        List<string> filteredColForHide; // Отфильтрованный список колонок для скрытия
        string[] hCols = new string[2]; // Диапазон скрытых колонок

        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            this.BackColor = Color.FromArgb(240, 189, 170);
            comboBox1.DataSource = colForHide;
            comboBox2.DataSource = colForHide;
        }

        private void Form3_FormClosed(object sender, FormClosedEventArgs e) // При закрытии формы
        {
            frm1.frm3Loaded = 0;
        }

        private void button1_Click(object sender, EventArgs e) // Кнопка Применить / Фильтр
        {
            frm1.formFiltered = true;
            frm1.filterString = textBox1.Text;
            frm1.customFilter();
        }

        private void button2_Click(object sender, EventArgs e) // Кнопка Очистить
        {
            textBox1.Text = "";
            
        }
        private void button3_Click(object sender, EventArgs e)  // Кнопка Отменить изменения
        {
            resetFilter();
        }

        public void reset()
        {
            resetFilter();
            resetSort();
        }
        private void resetFilter() // Сбрасываем изменения в строке фильтра
        {
            textBox1.Text = frm1.filterString;
            textBox1.Refresh();
        }

        private void button4_Click(object sender, EventArgs e) // Кнопка удалить фильтр
        {
            frm1.removeFilter(false);
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e) //При смене выбранного значения в ComboBox1
        {
            int idx = colForHide.IndexOf(comboBox1.SelectedValue.ToString());

            filteredColForHide = (from string str in colForHide where colForHide.IndexOf(str) >= idx select str).ToList();
            comboBox2.DataSource = filteredColForHide;
        }

        private void button5_Click(object sender, EventArgs e) // Кнопка Фильтр по выд.
        {
            frm1.filterBySelection();
            resetFilter();

        }

        private void button6_Click(object sender, EventArgs e) // Кнопка Скрыть
        {
            hCols[0] = comboBox1.SelectedValue.ToString();
            hCols[1] = comboBox2.SelectedValue.ToString();
            frm1.hiddenColumns = hCols;
            frm1.hideColumns();
        }

        private void button7_Click(object sender, EventArgs e) // Кнопка Показать
        {
            frm1.showColumns();
        }

        private void button8_Click(object sender, EventArgs e) // Кнопка Применить / Сортировка
        {
            sortString = textBox2.Text;
            frm1.sort(sortString);
        }

        private void button9_Click(object sender, EventArgs e) // Кнопка Очистить  / Сортировка
        {
            textBox2.Text = "";
        }

        private void button10_Click(object sender, EventArgs e) // Кнопка Отменить изменения / Сортировка
        {
            resetSort();
        }

        private void resetSort() // Сбрасываем изменения в строке сортировки
        {
            textBox2.Text = frm1.sortString;
            textBox2.Refresh();
        }
        private void button11_Click(object sender, EventArgs e) // Кнопка Убрать сортировку
        {
            /*
            if (frm1.formFiltered == false)
            {
                frm1.isViewed = false;
            }
            frm1.sort("");
            */
            frm1.removeSort(false);
        }
    }
}
