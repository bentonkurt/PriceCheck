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
    public partial class Form2 : Form
    {
        public Form1 frm1; // родительская форма
        public string tableName; // Имя таблицы
        public string colName; // Имя колонки
        public bool colIsText; // Является ли текстовым тип данных в колонке
        public bool colIsDate; // Является ли тип данных в колонке датой
        public List<string> filterItems; // Источник значений для checkedListBox1
        List<string> comboSource; // Источник данных для ComboBox
        string comboBoxItem; // Выбранное значение в ComboBox
        bool dialogShowed = false; // Показан ли диалог удаления фильтра

        public DataTable filterSource; // Источник значений для checkedListBox1
        string strFilter; // строка для фильтра для DataGridWiev
        
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.BackColor = Color.FromArgb(240, 189, 170);
            // Определяем размер контроллов для формы
            foreach (Control cnt in this.Controls)
            {
                cnt.Width = this.Width - 19;
            }

            getcomboSource(); // Определяем источник данных для comboBox

            // Устанавливаем источник значений            
            checkedListBox1.DataSource = filterSource;
            checkedListBox1.DisplayMember = colName;

            comboBoxItem = "=";

        }

        private void getcomboSource() // Определяем источник данных для comboBox
        {
            comboSource = new List<string>();
            comboSource.Add("равно");
            comboSource.Add("содержит");
            comboSource.Add("начинается с");
            comboSource.Add("не равно");
            comboSource.Add("не содержит");
            comboSource.Add("не начинается с");

            if (colIsText == false)
            {
                comboSource.Add("больше");
                comboSource.Add("больше или равно");
                comboSource.Add("меньше");
                comboSource.Add("меньше или равно");
            }
            comboBox1.DataSource = comboSource;
        }

        private void Form2_Deactivate(object sender, EventArgs e) // При деактивации скрываем форму
        {
            if (dialogShowed == false)
            {
                this.Hide();
                frm1.frm2Loaded = 1;
            }
        }

        private void Form2_Resize(object sender, EventArgs e)
        {
            foreach (Control cnt in this.Controls)
            {
                cnt.Width = this.Width - 19;
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e) // При изменении выбора в comboBox
        {
            valueChange();
        }


        private void valueChange() // При изменении строки запроса или выбора в ComboBox
        {
            switch (comboBox1.SelectedValue.ToString())
            {
                case "равно":
                    {
                        comboBoxItem = "=";
                    };
                    break;
                case "содержит":
                    {
                        comboBoxItem = "Cont";
                    };
                    break;
                case "начинается с":
                    {
                        comboBoxItem = "Beg";
                    };
                    break;
                case "не равно":
                    {
                        comboBoxItem = "<>";
                    }
                    break;
                case "не содержит":
                    {
                        comboBoxItem = "NotCont";
                    }
                    break;
                case "не начинается с":
                    {
                        comboBoxItem = "NotBeg";
                    }
                    break;

                case "больше":
                    {
                        comboBoxItem = ">";
                    };
                    break;

                case "больше или равно":
                    {
                        comboBoxItem = ">=";
                    };
                    break;

                case "меньше":
                    {
                        comboBoxItem = "<";
                    };
                    break;

                case "меньше или равно":
                    {
                        comboBoxItem = "<=";
                    };
                    break;
            }
        }

        private void button1_Click(object sender, EventArgs e) // Кнопка По возрастанию
        {
            dialogShowed = true;
            DialogResult res = MessageBox.Show("Добавить к сортировке?", "Диалог", MessageBoxButtons.YesNoCancel);
            dialogShowed = false;
            switch (res)
            {
                case DialogResult.No:
                    {
                        frm1.isSorted = true;
                        frm1.sort("[" + colName + "] ASC");
                    };
                    break;
                case DialogResult.Yes:
                    {
                        frm1.isSorted = true;
                        frm1.sortAdv("[" + colName + "] ASC");
                    };
                    break;
                case DialogResult.Cancel:
                    {
                        frm1.removeSort(true);
                    };
                    break;
            }
        }

        private void button2_Click(object sender, EventArgs e) // Кнопка По убыванию
        {
            dialogShowed = true;
            DialogResult res = MessageBox.Show("Добавить к сортировке?", "Диалог", MessageBoxButtons.YesNoCancel);
            dialogShowed = false;
            switch (res)
            {
                case DialogResult.No:
                    {
                        frm1.isSorted = true;
                        frm1.sort("[" + colName + "] DESC");
                    };
                    break;
                case DialogResult.Yes:
                    {
                        frm1.isSorted = true;
                        frm1.sortAdv("[" + colName + "] DESC");
                    };
                    break;
                case DialogResult.Cancel:
                    {
                        frm1.removeSort(true);
                    };
                    break;
            }
        }

        private void button3_Click(object sender, EventArgs e) // Кнопка Применить фильтр
        {
            filterItems = new List<string>(); // Список выбранных значений
            strFilter = ""; // Строка фильтра
            filterItems = (from DataRowView str in checkedListBox1.CheckedItems select str[colName].ToString()).ToList();

            if (filterItems.Count == 0) // Если нет ни одной галочки
            {
                if (textBox1.Text.Length > 0) // Если в текстбоксе введен какой-то текст
                {
                    string filParam;
                    if (colIsText == true)
                    {
                        filParam = "[" + colName + "]";
                    }
                    else
                    {
                        filParam = "Convert([" + colName + "], System.String)";
                    }


                    switch (comboBoxItem)
                    {
                        case "Cont":
                            strFilter = filParam + " Like '%" + textBox1.Text + "%'";
                            break;
                        case "Beg":
                            strFilter = filParam + " Like '" + textBox1.Text + "%'";
                            break;
                        case "NotCont":
                            strFilter = filParam + " Not Like '%" + textBox1.Text + "%'";
                            break;
                        case "NotBeg":
                            strFilter = filParam + " Not Like '" + textBox1.Text + "%'";
                            break;
                        default:
                            strFilter = "[" + colName + "] " + comboBoxItem + "'" + textBox1.Text + "'";
                            break;
                    }
                }
                else
                {
                    switch (comboBoxItem)
                    {
                        case "=":
                            strFilter = "[" + colName + "]" + " IS NULL";
                            break;
                        case "<>":
                            strFilter = "[" + colName + "]" + " IS NOT NULL";
                            break;
                        default:
                            {
                                dialogShowed = true;
                                MessageBox.Show("Данные для фильтра не выбраны!", "Сообщение");
                                dialogShowed = false;
                                return;
                            };
                            break;
                    }
                }
            }
            else // Если есть проставленные галочки
            {

                foreach (string str in filterItems) // Перебираем все выбранные значения
                {
                    if (str == "") // Если значение пустое, то ставим IS NULL OR
                    {
                        strFilter = strFilter + "[" + colName + "] IS NULL OR ";
                    }
                    else
                    {
                        strFilter = strFilter + "[" + colName + "] ='" + str + "' OR ";
                    }
                }
                strFilter = strFilter.Substring(0, strFilter.Length - 3);
            }
            frm1.applyFilter(strFilter);
            frm1.addFilter(colName);
            checkedListBox1.DataSource = filterSource;
            checkedListBox1.DisplayMember = colName;
            for (int i = 0; i < checkedListBox1.CheckedIndices.Count; i++)
            {
                checkedListBox1.SetItemChecked(checkedListBox1.CheckedIndices[i], false);
            }
        }

        private void button4_Click(object sender, EventArgs e) // Кнопка Удалить фильтр
        {
            dialogShowed = true;
            DialogResult res = MessageBox.Show("Очистить фильтр?", "Диалог", MessageBoxButtons.YesNoCancel);
            dialogShowed = false;

            switch (res)
            {
                case DialogResult.Yes:
                    {
                        frm1.removeFilter(true);
                        frm1.addFilter(colName);
                        checkedListBox1.DataSource = filterSource;
                        checkedListBox1.DisplayMember = colName;


                    };
                    break;
                case DialogResult.No:
                    {
                        frm1.removeFilter(false);
                        frm1.addFilter(colName);
                        checkedListBox1.DataSource = filterSource;
                        checkedListBox1.DisplayMember = colName;
                    }
                    break;
                case DialogResult.Cancel:
                    {
                        return;
                    };
                    break;
            }
        }

        private void button5_Click(object sender, EventArgs e) // Кнопка Скопир. наим.
        {
            Clipboard.SetText(colName);
        }

        private void button6_Click(object sender, EventArgs e) // Кнопка Выделить столбец
        {
            frm1.selectColumn();
        }

        public void moveON(int x, int y) // Для перемещения вместе с родительской формой
        {
            this.Location.Offset(x, y);
        }


    }
}
