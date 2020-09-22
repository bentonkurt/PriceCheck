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
    public partial class Form5 : Form
    {
        public Form1 frm1; // родительская форма
        public int[] sel = new int[4]; // Выделенный диапазон
        int i = 0; // Текущая ячейка
        bool initSearch = false; // Инициирован ли поиск
        bool secondSearch = false; // Начат ли поиск сначала
        string[] comboSource1 = new string[3];
        string[] comboSource2 = new string[3];
        string[] comboSource3 = new string[2];


        public Form5()
        {
            InitializeComponent();
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            this.BackColor = Color.FromArgb(240, 189, 170);
            
            comboSource1[0] = "Текущий столбец";
            comboSource1[1] = "Вся таблица";
            comboSource1[2] = "Выделенный диапазон";
            
            comboSource2[0] = "Равно";
            comboSource2[1] = "Содержит";
            comboSource2[2] = "Начинается с";

            comboSource3[0] = "Вниз";
            comboSource3[1] = "Вверх";

            comboBox1.DataSource = comboSource1;
            comboBox4.DataSource = comboSource1;

            comboBox2.DataSource = comboSource2;
            comboBox5.DataSource = comboSource2;

            comboBox3.DataSource = comboSource3;
            comboBox6.DataSource = comboSource3;
        }

        private void Form5_FormClosed(object sender, FormClosedEventArgs e)
        {
            frm1.frm5Loaded = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            initSearch = false;
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            initSearch = false;
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            initSearch = false;
        }

        private void comboBox3_SelectedValueChanged(object sender, EventArgs e)
        {
            initSearch = false;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            initSearch = false;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            initSearch = false;
        }

        private void comboBox4_SelectedValueChanged(object sender, EventArgs e)
        {
            initSearch = false;
        }

        private void comboBox5_SelectedValueChanged(object sender, EventArgs e)
        {
            initSearch = false;
        }

        private void comboBox6_SelectedValueChanged(object sender, EventArgs e)
        {
            initSearch = false;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            initSearch = false;
        }

        private void button1_Click(object sender, EventArgs e) // Кнопка Найти на первой вкладке
        {
            found(textBox1.Text, comboBox2.SelectedItem.ToString(), comboBox1.SelectedItem.ToString(), comboBox3.SelectedItem.ToString() == "Вверх", checkBox1.Checked == true);
        }

        private void button2_Click(object sender, EventArgs e) // Кнопка Найти на второй вкладке
        {
            found(textBox2.Text, comboBox5.SelectedItem.ToString(), comboBox4.SelectedItem.ToString(), comboBox6.SelectedItem.ToString() == "Вверх", checkBox2.Checked == true);
        }

        private void found(string what, string how, string place, bool way, bool withRegister) // Поиск
        {
            if (initSearch == false)
            {
                sel = frm1.selectedRange;
                frm1.find(what, how, place, way, withRegister);
                if (frm1.foundRange.Count < 1)                          // Если массив пустой
                {
                    if (secondSearch == true)
                    {
                        MessageBox.Show("Ничего не найдено!", "Сообщение");
                        initSearch = false;
                        secondSearch = false;
                        return;
                    }
                    else
                    {
                        secondSearch = true;
                        DialogResult result = MessageBox.Show("Ничего не найдено! Поискать в остальном диапазоне?", "Диалог", MessageBoxButtons.YesNo);
                        switch (result)
                        {
                            case DialogResult.Yes:
                                {
                                    secondSearch = true;
                                    switch (place)
                                    {
                                        case "Текущий столбец":
                                            {
                                                frm1.selectedRange[0] = 0;
                                                frm1.selectedRange[1] = sel[1];
                                            };
                                            break;
                                        case "Вся таблица":
                                            {
                                                frm1.selectedRange[0] = 0;
                                                frm1.selectedRange[1] = 0;
                                            };
                                            break;
                                        case "Выделенный диапазон":
                                            {
                                                frm1.selectedRange[0] = frm1.selectedRange[2];
                                                frm1.selectedRange[1] = frm1.selectedRange[3];
                                            };
                                            break;
                                    }
                                    found(what, how, place, way, withRegister);
                                    i = 0;
                                    return;
                                };
                                break;
                            case DialogResult.No:
                                {
                                    initSearch = false;
                                    secondSearch = false;
                                    return;
                                };
                                break;
                        }
                    }
                }
                else                                                    // Продоложаем работу
                {
                    initSearch = true;
                    sel = frm1.selectedRange;
                    i = 0;
                }
            }

            if (i >= frm1.foundRange.Count)
            {
                MessageBox.Show("Найдена последняя позиция!", "Сообщение");
                initSearch = false;
                secondSearch = false;
                return;
            }
            frm1.showFound(i);
            i++;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e) // При смене вкладки
        {
            switch(tabControl1.SelectedIndex)
            {
                case 0: 
                    {
                        textBox1.Text = textBox2.Text;
                        comboBox1.SelectedItem = comboBox4.SelectedItem;
                        comboBox2.SelectedItem = comboBox5.SelectedItem;
                        comboBox3.SelectedItem = comboBox6.SelectedItem;
                        checkBox1.Checked = checkBox2.Checked;
                    }; 
                    break;
                case 1:
                    {
                        textBox2.Text = textBox1.Text;
                        comboBox4.SelectedItem = comboBox1.SelectedItem;
                        comboBox5.SelectedItem = comboBox2.SelectedItem;
                        comboBox6.SelectedItem = comboBox3.SelectedItem;
                        checkBox2.Checked = checkBox1.Checked;
                    };
                    break;
            }
        }

        private void button3_Click(object sender, EventArgs e) // Кнопка Заменить
        {
            found(textBox2.Text, comboBox5.SelectedItem.ToString(), comboBox4.SelectedItem.ToString(), comboBox6.SelectedItem.ToString() == "Вверх", checkBox2.Checked == true);
            if (frm1.foundRange.Count > 0)
            {
                change(frm1.cCell, textBox2.Text, textBox3.Text);
            }      

            if (textBox2.Text != textBox3.Text) // Меняем статус в родительской форме
            {
                frm1.dataRedacted = true;
            }
            
        }

        private void  change (DataGridViewCell cell, string initStr, string chngStr) // Замена значения в заданной ячейке
        {
            if (cell.Value.ToString().Length < 1)
            {
                cell.Value = chngStr;
            }
            else
            {
                if (initStr.Length > 0)
                {
                    cell.Value = cell.Value.ToString().Replace(initStr, chngStr);
                }
            }
            
        }

        private void button4_Click(object sender, EventArgs e) // Кнопка Заменить все
            {
                found(textBox2.Text, comboBox5.SelectedItem.ToString(), comboBox4.SelectedItem.ToString(), comboBox6.SelectedItem.ToString() == "Вверх", checkBox2.Checked == true);
                changeAll(textBox2.Text, textBox3.Text);
            }
            private void changeAll(string initStr, string chngStr) // Замена значения в заданной ячейке
            {
                foreach (DataGridViewCell cell in frm1.foundRange)
                {
                    if (cell.Value == null || cell.Value.ToString().Length < 1)
                    {
                        cell.Value = chngStr;
                    }
                    else
                    {
                        if (initStr.Length > 0)
                        {
                            cell.Value = cell.Value.ToString().Replace(initStr, chngStr);
                        }
                    }
                }
            }
        }
}
