using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.Common;
using System.Threading;
using System.Reflection;
using System.Globalization;

namespace PriceCheck
{
    public partial class Form1 : Form
    {
        string strConn; // строка подключения
        string strSQL; // строка SQL
        SqlConnection cn; // соедиенение
        SqlCommand cmd; // SQL - команда
        SqlTransaction txn; // SQL - транзакция
        SqlDataReader rdr; // DataReader

        // для вывода таблиц в TreeView
        string[,] tabs = new string[2, 17]; // список таблиц
        List<TreeNode> lTn; // Список нод

        // Для вывода данных
        BindingSource bsPage = new BindingSource(); // BindingSource для страниц
        BindingSource bsVal = new BindingSource(); // BindingSource для значений
        string tabName; // Имя текущей таблицы
        advDataSet ads; // Датасет для вывода значений
        DataTable tbl; // Текущая таблица 
        DataView dw; // Текущий DataView - источник данных для dataGridView1
        List<string> allColumns = new List<string>(); // Список названий всех колонок в dataGridView
        int CurPg; // Номер текущей страницы
        int cntPg; // Кол-во страниц в датасете
        int RowCnt; // Количество строк в текущей таблице
        bool shwSchema = false; // Выведена ли схема

        // Для обновления
        SqlCommand[] sLogics; // массив логики передачи обновлений
        List<string[]> sParam; // массив параметров логики передачи обновлений

        // Для представления
        public bool isSorted = false; // Применена ли сортировка к исходной таблице
        public string filterString = ""; // Строка фильтра
        public string sortString = ""; // Строка сортировки

        Form2 frm2; // Форма для фильтра в столбце
        public int frm2Loaded = 0; // Загружена ли форма 2
        public DataTable filSource; // Источник строк для фильтра колонки
        string colName; // Имя колонки

        // Для фильтра
        bool colIsText = false; // Является ли текстовым тип данных в колонке
        bool colIsDate = false; // Является ли тип данных в колонке датой
        bool formFilt = false; // Определяет, фильтрована ли данная форма (public bool formFiltered)

        Form3 frm3; // Форма для фильтрации и скрытия столбцов
        public int frm3Loaded = 0; // Загружена ли форма 3
        List<string> filteredColumns; // Список названий отфильтрованных колонок в dataGridView
        List<string> sortedAscColumns; // Список названий отсортированных по возрастанию колонок в dataGridView
        List<string> sortedDescColumns; // Список названий отсортированных по убыванию колонок в dataGridView

        public string[] hiddenColumns = new string[2];  // Диапазон скрытых колонок
        bool colsHid = false; // Скрыты ли колонки (public bool colsIsHidden = false)

        // Для поиска
        Form5 frm5; // Форма для поиска
        public bool frm5Loaded = false; // Загружена ли форма 5
        public int[] selectedRange = new int[4]; // Выделенный диапазон
        public List<DataGridViewCell> foundRange = new List<DataGridViewCell>(); // Список найденных ячеек
        public DataGridViewCell cCell; // Текущая ячейка

        // Остальные формы
        Form4 frm4; // Форма для работы с базой
        public bool frm4Loaded = false; // Загружена ли форма 4
        Form6 frm6; // Форма для экспорта
        public bool frm6Loaded = false; // Загружена ли форма 6

        // Отредактированы ли данные в DataGridView
        bool dataRed;

        // Шрифт по дефолту
        //Font defFont = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular);
        Font defFont = new Font("Times New Roman", 10F, FontStyle.Regular);
        Color cvet = Color.FromArgb(255, 0, 0); // Цвет для выделения
        object filler = "e"; // Заполнитель

        // Для выделения измененных ячеек
        List<DataGridViewCell> changedCells = new List<DataGridViewCell>(); // Измененные ячейки
        Dictionary<DataRow, List<string>> lst = new Dictionary<DataRow, List<string>>();
        bool showChanged = false; // Показаны ли выделенные ячейки

        // Для изменения параметров вывода
        Form8 frm8;
        public bool frm8Loaded = false; // Загружена ли форма 8
        int pSize = 50000; // Размер страницы
        bool isWhole = false; // Выведена ли вся таблица

        public Form1()
        {
            InitializeComponent();
            this.KeyPreview = true; // Нужно для работы сочетания клавиш

            strConn = @"Data Source=W69\W69SQLEXPRESS;Initial Catalog=CheckPrice;Integrated Security=True";
            cn = new SqlConnection(strConn);
            cn.StateChange += new StateChangeEventHandler(cn_StateChange); // добавляем обработку состояния

            initTabList(); // Список таблиц

            if (cn.State.ToString() == "Open")
            {
                button1.Text = "Закрыть подключение";
                toolStripStatusLabel1.Text = "Открыто";
            }
            else
            {
                button1.Text = "Открыть подключение";
                toolStripStatusLabel1.Text = "Закрыто";
            }
            // Делаем недоступными кнопки "Выполнить и очистить"
            button13.Enabled = false;
            button14.Enabled = false;

            dataRedacted = false; // Отредактированы ли данные
        }

        // Новый класс advDataSet
        class advDataSet  // Расширенный датасет
        {
            string tabName; // Имя таблицы
            int pageSize; // Чисто строк на странице
            DataTable tabl; // Целевая таблица
            List<string> Columns; // Список названий колонок в таблице

            DataView dv; // Первоначальный DataView
            List<int[]> pageCnt; // Номера строк для заполнения страниц
            int cntPg; // Количество страниц
            List<DataView> pages; // Массив страниц
            public advDataSet()
            {
            }

            public void setPageSize(int PageSize) // Устанавливаем размер страницы
            {
                pageSize = PageSize;
            }

            public void Fill(string TableName, SqlConnection cn) // Заполняем таблицу первичными данными с помощью DataAdapter'а
            {
                tabName = TableName; // Задаем имя таблицы
                tabl = new DataTable(TableName); // Создаем новую таблицу с заданными именем
                SqlDataAdapter ad = new SqlDataAdapter("SELECT * FROM " + TableName, cn);
                ad.Fill(tabl);
                Columns = (from DataColumn col in tabl.Columns select col.ColumnName).ToList(); // Заполняем массив имен столбцов
            }

            public List<string> getAllColumns() // Возвращаем список всех колонок
            {
                return Columns;
            }

            public void ReFill() // Команды заполнения с параметрами
            {
                refilling("", "", pageSize);
            }

            public void ReFill(int PageSize) // Команды заполнения с параметрами
            {
                refilling("", "", PageSize);
            }

            public void ReFill(string RowFilter, string Sort) // Команды заполнения с параметрами
            {
                refilling(RowFilter, Sort, pageSize);
            }

            public void ReFill(string RowFilter, string Sort, int PageSize) // Команды заполнения с параметрами
            {
                refilling(RowFilter, Sort, PageSize);
            }

            private void refilling(string RowFilter, string Sort, int PageSize) // Заполняем данными из обработанной таблицы
            {
                pageSize = PageSize; // Сохраняем разме строки

                Dictionary<DataRow, int> dic = new Dictionary<DataRow, int>(); // Создаем словарь
                for (int p = 0; p < tabl.Rows.Count; p++)
                {
                    dic.Add(tabl.Rows[p], -1);
                }

                dv = new DataView(tabl, RowFilter, Sort, DataViewRowState.CurrentRows); // Создаем промежуточный DataView

                for (int p = 0; p < dv.Count; p++)
                {
                    dic[dv[p].Row] = p;
                }

                int i = dv.Count; // Количество строк в промежуточном DataView
                int j = i % pageSize; // Остаток от деления
                int k = (i - j) / pageSize; // Количество страниц

                pageCnt = new List<int[]>(); // Инициализируем новый массив номеров строк для разбивки и заполняем

                for (int l = 0; l < k; l++)
                {
                    int[] item = new int[2];
                    item[0] = l * pageSize;
                    item[1] = item[0] + PageSize - 1;
                    pageCnt.Add(item);
                }
                if (j > 0)
                {
                    pageCnt.Add(new int[] { k * pageSize, i - 1 });
                }

                cntPg = pageCnt.Count; // задаем количество страниц
                pages = new List<DataView>(); // инициализируем новый список страниц

                for (int m = 0; m < cntPg; m++) // Проводим процедуру для всех страниц
                {
                    int nach = pageCnt[m][0];
                    int kon = pageCnt[m][1];

                    EnumerableRowCollection<DataRow> query = from row in tabl.AsEnumerable() where inDick(dic, row, nach, kon) == true select row;
                    DataView promDv = query.AsDataView();
                    promDv.Sort = Sort;
                    pages.Add(promDv);
                }
            }

            private bool inDick(Dictionary<DataRow, int> dic, DataRow row, int nachalo, int koniec)
            {
                try
                {
                    if (dic[row] >= nachalo && dic[row] <= koniec)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                catch
                {
                    return true;
                }

            }

            public void ReFillWhole() // Команды заполнения с параметрами
            {
                refillingWhole("", "", pageSize);
            }

            public void ReFillWhole(int PageSize) // Команды заполнения с параметрами
            {
                refillingWhole("", "", PageSize);
            }

            public void ReFillWhole(string RowFilter, string Sort) // Команды заполнения с параметрами
            {
                refillingWhole(RowFilter, Sort, pageSize);
            }

            public void ReFillWhole(string RowFilter, string Sort, int PageSize) // Команды заполнения с параметрами
            {
                refillingWhole(RowFilter, Sort, PageSize);
            }

            private void refillingWhole(string RowFilter, string Sort, int PageSize) // Заполняем данными из обработанной таблицы
            {
                pageSize = PageSize; // Сохраняем размер строки
                dv = new DataView(tabl, RowFilter, Sort, DataViewRowState.CurrentRows); // Создаем промежуточный DataView
                pages = new List<DataView>(); // инициализируем новый список страниц
                pages.Add(dv);
                cntPg = 1;
            }

            public int returnCntPg() // Возвращаем количество страниц
            {
                cntPg = pages.Count;
                return cntPg;
            }

            public DataTable returnSrc() // Возвращаем первоначальную таблицу
            {
                return tabl;
            }

            public DataView returnSrc(int Page) // Возвращаем нужную таблицу
            {
                if (pages.Count == 0)
                {
                    pages.Add(dv);
                }
                return pages[Page - 1];
           
            }

            public List<DataView> returnBindSrc() // Возвращаем список страниц
            {
                if (pages.Count == 0)
                {
                    pages.Add(dv);
                }
                return pages;
            }
        }
        // Конец нового класса 


        private void initTabList() // Инициирует список таблиц
        {
            tabs[0, 0] = "PriceBuh"; tabs[1, 0] = "1";
            tabs[0, 1] = "VedCenPr"; tabs[1, 1] = "1";
            tabs[0, 2] = "Prov"; tabs[1, 2] = "1";
            tabs[0, 3] = "SvodView"; tabs[1, 3] = "2";
            tabs[0, 4] = "Sheet6View"; tabs[1, 4] = "2";
            tabs[0, 5] = "Svod"; tabs[1, 5] = "2";
            tabs[0, 6] = "Sheet6"; tabs[1, 6] = "2";
            tabs[0, 7] = "PromPrice"; tabs[1, 7] = "2";
            tabs[0, 8] = "OMTS"; tabs[1, 8] = "2";
            tabs[0, 9] = "PromCode"; tabs[1, 9] = "2";
            tabs[0, 10] = "PromName"; tabs[1, 10] = "2";
            tabs[0, 11] = "Sheet2"; tabs[1, 11] = "3";
            tabs[0, 12] = "Sheet3"; tabs[1, 12] = "3";
            tabs[0, 13] = "Sheet4"; tabs[1, 13] = "3";
            tabs[0, 14] = "Sheet5"; tabs[1, 14] = "3";
            tabs[0, 15] = "Prover"; tabs[1, 15] = "3";
            tabs[0, 16] = "OKEI"; tabs[1, 16] = "3";
        }

        private void execSQLCmd(string str) // Выполнить SQL-комманду
        {
            strSQL = str;
            cmd = new SqlCommand(strSQL, cn, txn);
            cmd.ExecuteNonQuery();
        }

        private void execMassSQLCmd(List<SqlCommand> lst) // Выполнить массив SQL-комманд
        {
            txn = cn.BeginTransaction();
            for (int i = 0; i < lst.Count; i++)
            {
                lst[i].Transaction = txn;
                lst[i].ExecuteNonQuery();
            }
            txn.Commit();
        }


        private int СurPage // Номер текущей страницы
        {
            set
            {
                CurPg = value;
                if (CurPg == 0)
                {
                    toolStripStatusLabel2.Text = "Загружена страница.: 0 из 0";
                }
                else
                {
                    toolStripStatusLabel2.Text = "Загружена страница.: " + CurPg.ToString() + " из " + cntPg.ToString() + " (" + RowCnt.ToString() + " строк)";
                }
            }
            get

            {
                return CurPg;
            }
        }

        public bool formFiltered // Отфильтрована ли форма
        {
            set
            {
                formFilt = value;

                if (filterString == "")
                {
                    formFilt = false;
                }


                if (formFilt == false)
                {
                    toolStripStatusLabel3.Text = "";
                }
                else
                {
                    toolStripStatusLabel3.Text = "ФЛТР";
                }
                statusStrip1.Refresh();
                bolding();
            }
            get
            {
                return formFilt;
            }
        }
        public bool colsIsHidden  // Скрыты ли колонки
        {
            set
            {
                colsHid = value;
                if (colsHid == false)
                {
                    toolStripStatusLabel4.Text = "";
                }
                else
                {
                    toolStripStatusLabel4.Text = "СКР";
                }
                statusStrip1.Refresh();
            }
            get
            {
                return colsHid;
            }
        }

        public bool dataRedacted // Отредактированы ли данные 
        {
            set
            {
                dataRed = value;
                if (dataRed == false)
                {
                    toolStripStatusLabel5.Text = "";
                    button8.Enabled = false;
                    button9.Enabled = false;
                    button10.Enabled = false;
                }
                else
                {
                    toolStripStatusLabel5.Text = "РЕД";
                    button8.Enabled = true;
                    button9.Enabled = true;
                    button10.Enabled = true;
                }
                statusStrip1.Refresh();
            }
            get
            {
                return dataRed;
            }
        }

        private bool showedSchema // Выведена ли схема
        {
            set
            {
                shwSchema = value;
                if (shwSchema == false)
                {
                    bindingNavigatorAddNewItem.Enabled = true;
                    bindingNavigatorDeleteItem.Enabled = true;
                    dataGridView1.ReadOnly = false;
                }
                else
                {
                    bindingNavigatorAddNewItem.Enabled = false;
                    bindingNavigatorDeleteItem.Enabled = false;
                    dataGridView1.ReadOnly = true;
                    dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.Automatic;
                }
            }
            get
            {
                return shwSchema;
            }
        }

        private void Form1_Resize(object sender, EventArgs e) // При изменении размера формы
        {
            dataGridView1.Width = this.Width - 319; // Растягиваем dataGridView1 по длине
            dataGridView1.Height = this.Height - 304; // Растягиваем dataGridView1 по ширине
            label1.Left = dataGridView1.Left + (dataGridView1.Width / 2) - (label1.Width / 2); // Чтобы надпись оставалась в серединке dataGridView1
            panel1.Top = this.Height - 250; //   Перемещаем panel1 в самый низ 224
            panel1.Left = (this.Width / 2) - (panel1.Width / 2); //   Перемещаем panel1 в самый низ
            treeView1.Height = this.Height - 304; // растягиваем treeView1 вместе с dataGridView1
        }

        private void Form1_Load(object sender, EventArgs e) // при загрузке формы выводим список таблиц и столбцов в dataGridView1
        {
            bindingNavigatorAddNewItem.Enabled = false;
            bindingNavigatorDeleteItem.Enabled = false;
            bindingNavigator1.BindingSource = bsPage;
            bindingNavigator2.BindingSource = bsVal;
            this.BackColor = Color.FromArgb(240, 189, 170);
            panel2.BackColor = Color.FromArgb(240, 189, 170);
            panel3.BackColor = Color.FromArgb(240, 189, 170);
            bindingNavigator1.BackColor = Color.FromArgb(240, 189, 170);
            bindingNavigator2.BackColor = Color.FromArgb(240, 189, 170);

            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            dataGridView1.EnableHeadersVisualStyles = false; // Реализует возможность менять стили заголовков
            dataGridView1.VirtualMode = true;
            dataGridView1.DataSource = bsVal;

            typeof(Control).InvokeMember("DoubleBuffered", BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.NonPublic,
            null, dataGridView1, new object[] { true }); // Делаем двойную буферизацию
        }

        private void Form1_Move(object sender, EventArgs e) // Закрываем форму 2 при перемещении
        {
            if (frm2Loaded == 2)
            {
                frm2.Close();
                frm2Loaded = 0;
            }
        }

        private void button1_Click(object sender, EventArgs e) // Открыть/закрыть соединение
        {
            try
            {
                switch (cn.State)
                {
                    case ConnectionState.Closed:
                        cn.Open();
                        break;
                    case ConnectionState.Broken:
                        {
                            cn.Close();
                            cn.Open();
                        };
                        break;
                    case ConnectionState.Open:
                        cn.Close();
                        break;
                    default:
                        { }
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
            }
        }

        private void cnSQLOpen() // Открываем соединение SQL
        {
            switch (cn.State)
            {
                case ConnectionState.Closed:
                    cn.Open();
                    break;
                case ConnectionState.Broken:
                    {
                        cn.Close();
                        cn.Open();
                    };
                    break;
                default:
                    { }
                    break;
            }
        }


        private void button2_Click(object sender, EventArgs e) // Кнопка Вывести список таблиц
        {
            // Открываем подключение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            overriding();
            showTables();
            showSchema();
            cn.Close(); // Закрываем подключение
        }

        public void showSchema() //Вывести в dataGridView1 список таблиц и столбцов
        {
            strSQL = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES ORDER BY TABLE_NAME";
            cmd = new SqlCommand(strSQL, cn);
            rdr = cmd.ExecuteReader();
            tbl = new DataTable("Schema");
            label1.Text = "Схема";
            label1.Refresh();
            tbl.Load(rdr);
            rdr.Close();
            bsPage.DataSource = null;
            bsVal.DataSource = tbl;
            dataGridView1.Refresh();
            showedSchema = true;
        }


        public void showTables() // Вывести в treeView список таблиц и столбцов
        {
            treeView1.Nodes.Clear(); // очищаем старые значения
            lTn = new List<TreeNode>(); // очищаем старые значения
            TreeNode node = new TreeNode("Таблицы"); // нода в дереве
            treeView1.Nodes.Add(node);
            TreeNode nd1 = new TreeNode("Первоначальные таблицы");
            node.Nodes.Add(nd1);
            TreeNode nd2 = new TreeNode("Проверочные таблицы");
            node.Nodes.Add(nd2);
            TreeNode nd3 = new TreeNode("Промежуточные таблицы");
            node.Nodes.Add(nd3);

            List<string> lst = new List<string>(); // список нод
            strSQL = "SELECT * FROM INFORMATION_SCHEMA.TABLES";
            cmd = new SqlCommand(strSQL, cn);
            rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                lst.Add(rdr["Table_Name"].ToString());
            }
            rdr.Close();

            for (int i = 0; i < 17; i++)
            {
                string nd = tabs[0, i];
                int num = Convert.ToInt16(tabs[1, i]);

                foreach (string str in lst)
                {
                    if (nd == str)
                    {
                        TreeNode tn = new TreeNode(nd);
                        lTn.Add(tn);
                        switch (num)
                        {
                            case 1:
                                nd1.Nodes.Add(tn);
                                break;
                            case 2:
                                nd2.Nodes.Add(tn);
                                break;
                            case 3:
                                nd3.Nodes.Add(tn);
                                break;
                        }
                        break;
                    }
                }
            }
            treeView1.ExpandAll(); // развернуть все
        }

        private void treeView1_BeforeSelect(object sender, TreeViewCancelEventArgs e) // Отменяем, если отмечена не таблица
        {
            List<TreeNode> tnds = new List<TreeNode>();
            tnds.Add(treeView1.Nodes[0]);
            tnds.Add(treeView1.Nodes[0].Nodes[0]);
            tnds.Add(treeView1.Nodes[0].Nodes[1]);
            tnds.Add(treeView1.Nodes[0].Nodes[2]);
            for (int i = 0; i < 4; i++)
            {
                if (e.Node == tnds[i])
                {
                    e.Cancel = true;
                }
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e) // После выбора ноды
        {
            tabName = e.Node.Text;
        }

        private void button3_Click(object sender, EventArgs e) // Кнопка Удалить таблицу
        {
            DialogResult res = MessageBox.Show("Удалить выбранную таблицу?", "Диалог", MessageBoxButtons.YesNo);
            switch(res)
            {
                case DialogResult.Yes:
                    {
                        switch (tabName)
                        {
                            case "Svod":
                                {
                                    strSQL = "DROP VIEW [SvodView]; DROP TABLE [Svod]";
                                };
                                break;
                            case "SvodView":
                                {
                                    strSQL = "DROP VIEW [SvodView]; DROP TABLE [Svod]";
                                };
                                break;
                            case "Sheet6":
                                {
                                    strSQL = "DROP VIEW [Sheet6View]; DROP TABLE [Sheet6]";
                                };
                                break;
                            case "Sheet6View":
                                {
                                    strSQL = "DROP VIEW [Sheet6View]; DROP TABLE [Sheet6]";
                                };
                                break;
                            default:
                                {
                                    strSQL = "DROP TABLE " + tabName;
                                };
                                break;
                        }

                        // Открываем подключение
                        try
                        {
                            cnSQLOpen();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Сообщение");
                            return;
                        }

                        txn = cn.BeginTransaction();
                        try
                        {
                            execSQLCmd(strSQL);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Сообщение");
                            txn.Commit();
                            cn.Close();
                            return;
                        }
                        txn.Commit();

                        showTables();
                        showSchema();
                        cn.Close(); // Закрывем подключение
                    };
                    break;

                case DialogResult.No: return;
                    break;
            }            
        }

        private void cn_StateChange(object sender, StateChangeEventArgs e) // добавляем обработку состояния
        {
            switch(e.CurrentState)
            {
                case ConnectionState.Closed:
                    {
                        button1.Text = "Открыть подключение";
                        toolStripStatusLabel1.Text = "Закрыто";
                    };
                    break;
                case ConnectionState.Open:
                    {
                        button1.Text = "Закрыть подключение";
                        toolStripStatusLabel1.Text = "Открыто";
                    };
                    break;
                case ConnectionState.Broken:
                    {
                        button1.Text = "Переподключиться";
                        toolStripStatusLabel1.Text = "Прервано";
                    };
                    break;
                case ConnectionState.Connecting:
                    {
                        toolStripStatusLabel1.Text = "Подключаемся";
                    };
                    break;
                case ConnectionState.Executing:
                    {
                        toolStripStatusLabel1.Text = "Выполняется команда";
                    };
                    break;
                case ConnectionState.Fetching:
                    {
                        toolStripStatusLabel1.Text = "Получение данных";
                    };
                    break;
            }
        }

        private void button4_Click(object sender, EventArgs e) //  Кнопка Вывести содержимое
        {
            try
            {
                showContent(); // Выводим содержимое
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("Таблица не выбрана!", "Сообщение");
            }
        }

        public void showContent() //Показать содержимое
        {
            cnSQLOpen(); // Открываем подключение
            overriding(); // Обнуляем все важные параметры
            showedSchema = false;
            dataGridView1.Columns.Clear(); // Очищаем коллекцию столбцов, чтобы не было проблем с оформлением


            // Выводим название таблицы
            label1.Text = tabName;
            label1.Refresh();

            // Создаем и заполняем расширенный датасет
            ads = new advDataSet();
            ads.setPageSize(pSize);
            ads.Fill(tabName, cn);
            refilling();

            tbl = ads.returnSrc(); // Задаем текущую таблицу
            vyvodTabl(1);

            allColumns = ads.getAllColumns(); // Задаем список колонок
            addLogics(); // Добавляем логику обновления
            cn.Close();
        }
        private void addLogics() // Создание логики обновления, добавления, удаления 
        {
            sLogics = new SqlCommand[3]; // Инициализируем новый список логики обновления
            sParam = new List<string[]>(); // Инициализируем новый список параметров логики обновления

            SqlCommand uCom; // команда обновления
            SqlCommand iCom; // команда вставки
            SqlCommand dCom; // команда удаления

            string dComProm = ""; // промежуточная строка удаления
            string iComProm = ""; // промежуточная строка вставки
            string iComVal = ""; // промежуточная строка для строки VALUES
            string uComProm = ""; // промежуточная строка для значений
            string uComVal = ""; // промежуточная строка для строки VALUES условий

            foreach (string node in allColumns)
            {
                string[] massParam = new string[5]; // массив параметров
                sParam.Add(massParam);

                massParam[0] = node; //Первоначальный параметр
                massParam[1] = "@" + node + "_Del"; //Параметры для удаления
                massParam[2] = "@" + node + "_Ins"; //Параметры для вставки
                massParam[3] = "@" + node + "_UpdOld"; //Параметры для обновления
                massParam[4] = "@" + node + "_UpdNew"; //Параметры для обновления

                dComProm = dComProm + "[" + massParam[0] + "]" + " = " + massParam[1] + " AND ";

                iComProm = iComProm + "[" + massParam[0] + "]" + ", ";
                iComVal = iComVal + massParam[2] + ", ";

                uComProm = uComProm + "[" + massParam[0] + "]" + " = " + massParam[4] + ", ";
                uComVal = uComVal + massParam[0] + " = " + massParam[3] + " AND ";
            }

            dComProm = dComProm.Substring(0, dComProm.Length - 5);
            dComProm = "DELETE FROM [" + tabName + "] WHERE " + dComProm;

            iComProm = iComProm.Substring(0, iComProm.Length - 2);
            iComVal = iComVal.Substring(0, iComVal.Length - 2);
            iComProm = "INSERT INTO [" + tabName + "] (" + iComProm + ") VALUES (" + iComVal + ")";

            uComProm = uComProm.Substring(0, uComProm.Length - 2);
            uComVal = uComVal.Substring(0, uComVal.Length - 5);
            uComProm = "UPDATE [" + tabName + "] SET " + uComProm + " WHERE " + uComVal;

            iCom = new SqlCommand(iComProm, cn);
            dCom = new SqlCommand(dComProm, cn);
            uCom = new SqlCommand(uComProm, cn);

            sLogics = new SqlCommand[3]; // создаем новый массив
            sLogics[0] = iCom;
            sLogics[1] = dCom;
            sLogics[2] = uCom;

        }

        private string withoutNull(dynamic i) // Превращаем строки с нулем в пустые строчки
        {
            string str = "";
            if (i == null)
            {
                str = "";
            }
            else
            {
                str = i.ToString();
            }
            return str;
        }

        private void button5_Click(object sender, EventArgs e) // Кнопка Парам. отобр.
        {
            if (frm8Loaded == false)
            {
                frm8 = new Form8();
                frm8.frm1 = this;
                frm8.Show();
                frm8Loaded = true;
            }
        }

        public void showAll() // Вывести всю таблицу в DataGridView
        {
            overriding();
            isWhole = true;
            allColumns = ads.getAllColumns();
            refilling();
            vyvodTabl(1);
            bolding();
        }

        public int returnPageSize() // Вернуть размер страницы
        {
            return pSize;
        }

        public void setupPageSize(int PageSize) // Установить размер страницы
        {
            pSize = PageSize;
            try
            {
                ads.setPageSize(pSize);
            }
            catch
            { }
        }

        private void bindingNavigatorMoveFirstItem_Click(object sender, EventArgs e)
        {
            if (showedSchema == false)
            {
                vyvodTabl(1);
                bolding();
            }
        }

        private void bindingNavigatorMovePreviousItem_Click(object sender, EventArgs e)
        {
            if (showedSchema == false)
            {
                СurPage--;
                vyvodTabl(СurPage);
                bolding();
            }
        }

        private void bindingNavigatorMoveNextItem_Click(object sender, EventArgs e)
        {
            if (showedSchema == false)
            {
                СurPage++;
                vyvodTabl(СurPage);
                bolding();
            }

        }

        private void bindingNavigatorMoveLastItem_Click(object sender, EventArgs e)
        {
            if (showedSchema == false)
            {
                СurPage = cntPg;
                vyvodTabl(СurPage);
                bolding();
            }
        }

        private void bindingNavigatorPositionItem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (showedSchema == false)
                {
                    try
                    {
                        СurPage = Convert.ToInt32(bindingNavigatorPositionItem.Text);
                        if (СurPage > cntPg)
                        {
                            СurPage = cntPg;
                        }

                        if (СurPage < 1)
                        {
                            СurPage = 1;
                        }
                    }
                    catch
                    {

                    }

                    vyvodTabl(СurPage);
                    bolding();
                }
            }
        }

        private void refilling() // Пересчет датасурса
        {
            if (isWhole == true)
            {
                if (formFiltered == false)
                {
                    if (isSorted == true)
                    {
                        ads.ReFillWhole("", sortString);
                    }
                    else
                    {
                        ads.ReFillWhole();
                    }
                }
                else
                {
                    if (isSorted == true)
                    {
                        ads.ReFillWhole(filterString, sortString);
                    }
                    else
                    {
                        ads.ReFillWhole(filterString, "");
                    }

                }
            }
            else
            {

                if (formFiltered == false)
                {
                    if (isSorted == true)
                    {
                        ads.ReFill("", sortString);
                    }
                    else
                    {
                        ads.ReFill();
                    }
                }
                else
                {
                    if (isSorted == true)
                    {
                        ads.ReFill(filterString, sortString);
                    }
                    else
                    {
                        ads.ReFill(filterString, "");
                    }
                }
            }
        }

        private void vyvodTabl(int nTabl) // Вывод нужной таблицы из Датасета
        {
            bsPage.DataSource = ads.returnBindSrc();
            bsPage.Position = nTabl - 1;
            dw = ads.returnSrc(nTabl); // Возвращаем нужную таблицу
            RowCnt = dw.Count;
            cntPg = ads.returnCntPg();
            СurPage = nTabl;
            showChanged = false;
            bsVal.DataSource = dw;
            dataGridView1.Refresh();
        }

        public void overriding() // Сброс при смене источника данных
        {
            try // закрываем форму фильтра
            {
                frm2.Close();
            }
            catch
            {
            }
            frm2Loaded = 0; // Обнуляем статус открытия формы для фильтра
            formFiltered = false; // Устанавливаем, что форма не отфильтрована
            filterString = ""; // Обнуляем строку фильтра
            sortString = ""; // Обнуляем строку сортировки

            try // закрываем форму для фильтрации и скрытия столбцов
            {
                frm3.Close();
            }
            catch
            {
            }
            frm3Loaded = 0; //  Обнуляем статус открытия формы для фильтрации и скрытия столбцов
            allColumns = new List<string>(); // Обнуляем список названий всех колонок в dataGridView
            filteredColumns = new List<string>(); // Обнуляем список названий отфильтрованных колонок в dataGridView
            hiddenColumns = new string[2];  // Обнуляем диапазон скрытых колонок
            colsIsHidden = false; // Скрыты ли колонки
            isSorted = false; // Применена ли сортировка

            try // Закрываем
            {
                frm5.Close(); // Форма для поиска
            }
            catch
            {
            }

            frm5Loaded = false; // Обнуляем статус открытия формы поиска
            selectedRange = new int[4]; // Обнуляем инфу о выделенном диапазоне
            foundRange = new List<DataGridViewCell>(); // Обнуляем список найденных ячеек
            changedCells = new List<DataGridViewCell>(); // Обнуляем список измененных ячеек
            lst = new Dictionary<DataRow, List<string>>(); // Обнуляем словарь измененных ячеек
            showChanged = false;
            isWhole = false;
        }

        private void button6_Click(object sender, EventArgs e) // Кнопка Фильт. и сорт.
        {
            if (dataGridView1.Columns.Count > 0 && showedSchema == false)
            {
                switch (frm3Loaded)
                {
                    case 0:
                        {
                            frm3 = new Form3();
                            frm3.frm1 = this;
                            frm3.filterString = filterString;
                            frm3.sortString = sortString;
                            frm3.colForHide = allColumns;
                            frm3.reset();
                            frm3.Show();
                            frm3Loaded = 2;
                        };
                        break;
                    case 1:
                        {
                            frm3.Show();
                            frm3Loaded = 2;
                        };
                        break;
                    case 2:
                        {
                            frm3.Hide();
                            frm3Loaded = 1;
                        };
                        break;
                }
            }
        }

        private void button7_Click(object sender, EventArgs e) // Кнопка Поиск и замена
        {
            callFrm5();
        }

        protected override void OnKeyDown(KeyEventArgs e) // Отрабатываем сочетания клавиш
        {
            if (e.KeyCode == Keys.F && e.Control)  // CTRL + F - поиск
            {
                if (dataGridView1.Columns.Count > 0)
                {
                    callFrm5();
                    e.Handled = true;
                }
            }

            if (e.KeyCode == Keys.S && e.Control) // CTRL + S - сохранение
            {
                Obnov();

                e.Handled = true;
            }

            if (e.KeyCode == Keys.Z && e.Control) // CTRL + Z - отмена изменений
            {
                reject();
            }

            if (e.KeyCode == Keys.A && e.Control) // CTRL + A - Выделить все
            {
                try
                {
                    dataGridView1.SelectionMode = DataGridViewSelectionMode.FullColumnSelect;
                    dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        dataGridView1.Columns[i].Selected = true;
                    }
                }
                catch
                {
                }

            }

            if (e.KeyCode == Keys.C && e.Control) // CTRL + С - скопировать выделенные ячейки
            {
                try
                {
                    Clipboard.SetDataObject(dataGridView1.GetClipboardContent());
                }
                catch
                { }
            }
        }

        private void callFrm5() // Вызываем форму поиска
        {
            if (frm5Loaded == false)
            {
                if (dataGridView1.Columns.Count > 0)
                {
                    frm5 = new Form5();
                    frm5.frm1 = this;
                    frm5.Show();
                    frm5Loaded = true;
                }
            }
            else
            {
                frm5.Focus();
            }
        }

        private void button8_Click(object sender, EventArgs e)  // Кнопка Сохранить
        {            
            Obnov();
        }

        private void Obnov() // Сохранение изменений
        {
            // Открываем подключение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            List<SqlCommand> lCmd = new List<SqlCommand>(); // Список комманд для выполнения

            List<DataRow> lstAdd = (from DataRow row in tbl.Rows where row.RowState == DataRowState.Added select row).ToList(); // Список добавленных строк
            List<DataRow> lstDel = (from DataRow row in tbl.Rows where row.RowState == DataRowState.Deleted select row).ToList(); // Список удаленных строк
            List<DataRow> lstUpd = (from DataRow row in tbl.Rows where row.RowState == DataRowState.Modified select row).ToList(); // Список измененных строк

            for (int i = 0; i < lstAdd.Count; i++) // Перебираем добавленные строки
            {
                SqlCommand eCmd = new SqlCommand(); // Элемент массива команд таблицы
                string cmdTxt = sLogics[0].CommandText;
                eCmd.Connection = cn;
                DataRow row = lstAdd[i];
                for (int j = 0; j < sParam.Count; j++)
                {
                    if (withoutNull(row[sParam[j][0], DataRowVersion.Current]) == "")  // if (row[sParam[j][0], DataRowVersion.Current] == null)
                    {
                        eCmd.Parameters.AddWithValue(sParam[j][2], DBNull.Value);
                    }
                    else
                    {
                        eCmd.Parameters.AddWithValue(sParam[j][2], row[sParam[j][0], DataRowVersion.Current]);
                    }
                }
                eCmd.CommandText = cmdTxt;
                lCmd.Add(eCmd);
            }

            for (int i = 0; i < lstDel.Count; i++) // Перебираем удаленные строки
            {
                SqlCommand eCmd = new SqlCommand(); // Элемент массива команд таблицы
                string cmdTxt = sLogics[1].CommandText;
                eCmd.Connection = cn;
                for (int j = 0; j < sParam.Count; j++)
                {
                    DataRow row = lstDel[i];

                    if (withoutNull(row[sParam[j][0], DataRowVersion.Original]) == "") // if (row[sParam[j][0], DataRowVersion.Original].ToString() == "")
                    {
                        cmdTxt = cmdTxt.Replace("= " + sParam[j][1], "IS NULL");
                    }
                    else
                    {
                        eCmd.Parameters.AddWithValue(sParam[j][1], row[sParam[j][0], DataRowVersion.Original]); // Если не поставить DataRowVersion, выдаст ошибку
                    }
                }
                eCmd.CommandText = cmdTxt;
                lCmd.Add(eCmd);
            }

            for (int i = 0; i < lstUpd.Count; i++) // Перебираем измененные строки
            {
                SqlCommand eCmd = new SqlCommand(); // Элемент массива команд таблицы
                string cmdTxt = sLogics[2].CommandText;
                eCmd.Connection = cn;
                DataRow row = lstUpd[i];

                for (int j = 0; j < sParam.Count; j++)
                {
                    if (row[sParam[j][0], DataRowVersion.Original] == DBNull.Value) // Меняем строку, иначе не будет работать
                    {
                        string prom = "= " + sParam[j][3] + " AND";
                        if (cmdTxt.IndexOf(prom) == -1)
                        {
                            cmdTxt = cmdTxt.Substring(0, cmdTxt.Length - prom.Length + 4);
                            cmdTxt = cmdTxt + "IS NULL";
                        }
                        else // если найдена строка с запятой, то заменяем
                        {
                            cmdTxt = cmdTxt.Replace(prom, "IS NULL AND");
                        }
                    }
                    else
                    {
                        eCmd.Parameters.AddWithValue(sParam[j][3], row[sParam[j][0], DataRowVersion.Original]); // Если не поставить DataRowVersion, выдаст ошибку
                    }

                    if (withoutNull(row[sParam[j][0], DataRowVersion.Current]) == "") // if (row[sParam[j][0], DataRowVersion.Current] == DBNull.Value || row[sParam[j][0], DataRowVersion.Current].ToString() == "")
                    {
                        eCmd.Parameters.AddWithValue(sParam[j][4], DBNull.Value); // Если не поставить DataRowVersion, выдаст ошибку
                    }
                    else
                    {
                        eCmd.Parameters.AddWithValue(sParam[j][4], row[sParam[j][0], DataRowVersion.Current]); // Если не поставить DataRowVersion, выдаст ошибку
                    }
                }
                eCmd.CommandText = cmdTxt;
                lCmd.Add(eCmd);
            }

            try
            {
                execMassSQLCmd(lCmd);// выполняем все команды из полученного массива
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                txn.Rollback();
                cn.Close();
                return;
            }

            dataRedacted = false; // Убираем статус Отредактировано
            ads.Fill(tabName, cn);
            tbl = ads.returnSrc();
            refilling();


            // Выбираем текущую страницу или первую, если такой страницы нет
            try
            {
                vyvodTabl(СurPage);
            }
            catch
            {
                vyvodTabl(1);
            }
            changedCells = new List<DataGridViewCell>(); // Обнуляем список измененных ячеек
            lst = new Dictionary<DataRow, List<string>>();
            showChanged = false;
            bolding();

            cn.Close();
            MessageBox.Show("Изменения сохранены!");
        }


        private void button9_Click(object sender, EventArgs e) // Кнопка Отменить
        {
            reject();
        }

        private void reject() // Отменить внесенные в dataGridView изменения
        {
            tbl.RejectChanges(); // Отменяем изменения в таблице
            dataGridView1.Refresh();
            dataRedacted = false; // Убираем статус Отредактировано
            changedCells = new List<DataGridViewCell>(); // Обнуляем список измененных ячеек
            lst = new Dictionary<DataRow, List<string>>();
            showChanged = false;

            try
            {
                frm2.Close();
            }
            catch
            { }

            try
            {
                frm5.Close();
            }
            catch
            { }
        }

        private void dataGridView1_UserDeletedRow(object sender, DataGridViewRowEventArgs e) // При удалении строки
        {
            dataRedacted = true;
        }

        private void button10_Click(object sender, EventArgs e) // Кнопка Выд. изм.
        {
            if (showChanged == false)
            {
                changedCells = new List<DataGridViewCell>();
                List<DataGridViewRow> lstRows = (from DataGridViewRow row in dataGridView1.Rows where isChanged(row) == true select row).ToList();
                for (int i = 0; i < lstRows.Count; i++)
                {
                    DataRow rw = ((DataRowView)(lstRows[i].DataBoundItem)).Row;
                    List<string> cols = lst[rw];
                    for (int j = 0; j < cols.Count; j++)
                    {
                        changedCells.Add(lstRows[i].Cells[cols[j]]);
                    }
                }
                for (int i = 0; i < changedCells.Count; i++)
                {
                    changedCells[i].Style.BackColor = cvet;
                }
                showChanged = true;
            }
            else
            {
                for (int i = 0; i < changedCells.Count; i++)
                {
                    changedCells[i].Style.BackColor = dataGridView1.Columns[changedCells[i].ColumnIndex].DefaultCellStyle.BackColor;
                }
                showChanged = false;
            }
        }

        private bool isChanged(DataGridViewRow row)
        {

            if (row.DataBoundItem == null)
            {

                return false;
            }
            else
            {
                DataRow rw = ((DataRowView)(row.DataBoundItem)).Row;
                if (rw.RowState == DataRowState.Modified || rw.RowState == DataRowState.Added)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
        }
        private void button11_Click(object sender, EventArgs e) // Кнопка работа с базой
        {
            if (frm4Loaded == false)
            {
                frm4 = new Form4();
                frm4.pf = this;
                frm4Loaded = true;
                frm4.Show();
            }
            else
            {
                frm4.Focus();
            }
        }

        private void button12_Click(object sender, EventArgs e)  // Кнопка Экспорт
        {
            if (frm6Loaded == false)
            {
                frm6 = new Form6();
                frm6.pf = this;
                frm6Loaded = true;
                frm6.Show();
            }
            else
            {
                frm6.Focus();
            }
        }

        private void button13_Click(object sender, EventArgs e)    // Очистить строку запроса
        {
            textBox1.Text = "";
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0)
            {
                button13.Enabled = true;
                button14.Enabled = true;
            }
            else
            {
                button13.Enabled = false;
                button14.Enabled = false;
            }
        }

        public void sort(string Sort) // Сортировка столбцов dataGridView1
        {
            isSorted = true;
            sortString = Sort;
            refilling();
            vyvodTabl(1);
            bolding();
        }

        public void sortAdv(string Sort) // Сортировка столбцов dataGridView1
        {
            if (sortString == "")
            {
                sortString = Sort;
            }
            else
            {
                sortString = sortString + ", " + Sort;
            }
            refilling();
            vyvodTabl(1);
            bolding();
        }
        public void removeSort(bool Clear) // Удаляем сортировку к dataGridView1
        {
            isSorted = false;
            refilling();
            vyvodTabl(1);
            bolding();

            if (Clear == true)
            {
                sortString = "";
            }
        }

        public void applyFilter(string FilStr) // Фильтруем dataGridView1
        {
            if (filterString == "")
            {
                filterString = "(" + FilStr + ")";
            }
            else
            {
                filterString = filterString + " AND (" + FilStr + ")";
            }
            formFiltered = true;
            refilling();
            vyvodTabl(1);
            bolding();
        }

        public void customFilter() // Фильтр, созданный ручной редакцией текста
        {
            refilling();
            vyvodTabl(1);
            bolding();
            if (filterString == "")
            {
                formFiltered = false;
            }

            try
            {
                frm2.Close();
            }
            catch { }
            frm2Loaded = 0;
        }

        public void filterBySelection() // Фильтр, созданный ручной редакцией текста
        {
            string prom;

            if (withoutNull(dataGridView1.Rows[selectedRange[0]].Cells[selectedRange[1]].Value) == "")
            {
                prom = "[" + dataGridView1.Columns[selectedRange[1]].Name + "] IS NULL";
            }
            else
            {
                prom = "[" + dataGridView1.Columns[selectedRange[1]].Name + "] = '" + dataGridView1.Rows[selectedRange[0]].Cells[selectedRange[1]].Value + "'";
            }

            if (filterString == "")
            {
                filterString = "(" + prom + ")";
            }
            else
            {
                filterString = filterString + " AND (" + prom + ")";
            }
            formFiltered = true;

            refilling();
            vyvodTabl(1);
            bolding();
            try
            {
                frm2.Close();
            }
            catch { }
            frm2Loaded = 0;
        }


        public void bolding() // Делаем жирным шрифт в отфильтрованных столбцах
        {
            foreach (DataGridViewColumn cl in dataGridView1.Columns)
            {
                cl.HeaderCell.Style.Font = new Font(defFont, FontStyle.Regular);
                cl.HeaderCell.Style.ForeColor = Color.Black;
            }

            if (isSorted == true)
            {
                sortedAscColumns = (from string str in allColumns where sortString.IndexOf("[" + str + "] ASC", StringComparison.CurrentCultureIgnoreCase) != -1 select str).ToList();
                sortedDescColumns = (from string str in allColumns where sortString.IndexOf("[" + str + "] DESC", StringComparison.CurrentCultureIgnoreCase) != -1 select str).ToList();

                foreach (string str in sortedAscColumns) // Ставим фиолетовый шрифт там, где отфильтрованные по возрастанию колонки
                {
                    dataGridView1.Columns[str].HeaderCell.Style.ForeColor = Color.FromArgb(255, 0, 255);
                }

                foreach (string str in sortedDescColumns) // Ставим зеленый шрифт там, где отфильтрованные по убыванию колонки
                {
                    dataGridView1.Columns[str].HeaderCell.Style.ForeColor = Color.FromArgb(4, 189, 55);
                }
            }

            if (formFiltered == true)
            {
                filteredColumns = (from string str in allColumns where filterString.IndexOf("[" + str + "]", StringComparison.CurrentCultureIgnoreCase) != -1 select str).ToList();
                foreach (string str in filteredColumns) // Ставим жирный шрифт там, где отфильтрованные колонки
                {
                    dataGridView1.Columns[str].HeaderCell.Style.Font = new Font(defFont, FontStyle.Bold);
                }
            }
        }

        public void removeFilter(bool Clear) // Удаляем фильтр к dataGridView1
        {
            formFiltered = false;
            filteredColumns = new List<string>();
            refilling();
            vyvodTabl(1);
            bolding();

            if (Clear == true)
            {
                filterString = "";
            }

            try
            {
                frm2.Close();
            }
            catch
            {

            }
        }

        public void hideColumns() // Скрываем колонки
        {
            colsIsHidden = true;
            int beg = allColumns.IndexOf(hiddenColumns[0]); // Индекс начальной позиции
            int fin = allColumns.IndexOf(hiddenColumns[1]); // Индекс конечной позиции
            for (int i = beg; i < fin + 1; i++)
            {
                dataGridView1.Columns[i].Visible = false;
            }
            dataGridView1.Refresh();
        }

        public void showColumns() // Убираем скрытие
        {
            colsIsHidden = false;
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].Visible = true;
            }
            dataGridView1.Refresh();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e) // При изменении значения в ячейке
        {
            
            if (e.ColumnIndex > -1)
            {
                DataRow rw;
                try
                {
                    rw = ((DataRowView)(dataGridView1.Rows[e.RowIndex].DataBoundItem)).Row;
                }
                catch
                {
                    return;
                }
                
                List<string> cols;

                try 
                {
                    cols = lst[rw];
                    cols.Add(dataGridView1.Columns[e.ColumnIndex].Name);
                }
                catch
                {
                    cols = new List<string>();
                    cols.Add(dataGridView1.Columns[e.ColumnIndex].Name);
                    lst.Add(rw, cols);
                }
                changedCells.Add(dataGridView1[e.ColumnIndex, e.RowIndex]);                
                dataRedacted = true;
            }

            DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
            try
            {
                switch (tabName)
                {
                    case "Sheet6": // два условия
                    case "Sheet6View":
                        {
                            if (dataGridView1.Columns[e.ColumnIndex].Name == "Price_Sub")
                            {
                                if (withoutNull(row.Cells["Price_Sub"].Value) != "")
                                {
                                    row.Cells["Price_Fin"].Value = row.Cells["Price_Sub"].Value;
                                }
                                else
                                {
                                    if (withoutNull(row.Cells["Price_Svod"].Value) != "")
                                    {
                                        row.Cells["Price_Fin"].Value = row.Cells["Price_Svod"].Value;
                                    }
                                    else
                                    {
                                        row.Cells["Price_Fin"].Value = row.Cells["Price"].Value;
                                    }
                                }
                            }

                            if (withoutNull(row.Cells["Price"].Value) == withoutNull(row.Cells["Price_Fin"].Value))  // Добавляем условное форматирование
                            {
                                row.Cells["Price_Fin"].Style.ForeColor = Color.FromArgb(0, 0, 255);
                            }
                            else
                            {
                                row.Cells["Price_Fin"].Style.ForeColor = Color.FromArgb(125, 0, 125);
                            }

                        };
                        break;

                    case "PromName":
                        {
                            if (dataGridView1.Columns[e.ColumnIndex].Name == "Name")
                            {
                                if (withoutNull(row.Cells["Name_Buh"].Value) != "")
                                {
                                    if (withoutNull(row.Cells["Name"].Value) == withoutNull(row.Cells["Name_Buh"].Value))
                                    {
                                        row.Cells["Comp2"].Value = "-";
                                        row.Cells["Name_Source"].Value = "ведомость";
                                    }
                                    else
                                    {
                                        row.Cells["Comp2"].Value = "+";
                                    }
                                }

                                if (withoutNull(row.Cells["Name_IVC"].Value) != "")
                                {
                                    if (withoutNull(row.Cells["Name"].Value) == withoutNull(row.Cells["Name_IVC"].Value))
                                    {
                                        row.Cells["Comp1"].Value = "-";
                                        row.Cells["Name_Source"].Value = "база";
                                    }
                                    else
                                    {
                                        row.Cells["Comp1"].Value = "+";
                                    }
                                }

                                if (withoutNull(row.Cells["Name_Sub"].Value) != "")
                                {
                                    if (withoutNull(row.Cells["Name"].Value) == withoutNull(row.Cells["Name_Sub"].Value))
                                    {
                                        row.Cells["Comp3"].Value = "-";
                                    }
                                    else
                                    {
                                        row.Cells["Comp3"].Value = "+";
                                    }
                                }

                                if (withoutNull(row.Cells["Name_Ver"].Value) != "")
                                {
                                    if (withoutNull(row.Cells["Name"].Value) == withoutNull(row.Cells["Name_Ver"].Value))
                                    {
                                        row.Cells["Comp4"].Value = "-";
                                    }
                                    else
                                    {
                                        row.Cells["Comp4"].Value = "+";
                                    }
                                }

                                if (withoutNull(row.Cells["Name_Copy"].Value) != "")
                                {
                                    if (withoutNull(row.Cells["Name"].Value) == withoutNull(row.Cells["Name_Copy"].Value))
                                    {
                                        row.Cells["Comp5"].Value = "-";
                                        row.Cells["Name_Copy_Source"].Value = row.Cells["Name_Source"].Value;
                                    }
                                    else
                                    {
                                        row.Cells["Comp5"].Value = "+";
                                    }
                                }

                                if (withoutNull(row.Cells["Name_IVC"].Value) != withoutNull(row.Cells["Name"].Value) &&
                                    withoutNull(row.Cells["Name_Buh"].Value) != withoutNull(row.Cells["Name"].Value) &&
                                    withoutNull(row.Cells["Name"].Value) != "")
                                {
                                    row.Cells["Name_Source"].Value = "подстановка";
                                }

                                if (withoutNull(row.Cells["Comp1"].Value) == "" && withoutNull(row.Cells["Comp2"].Value) == "" &&
                                    withoutNull(row.Cells["Comp3"].Value) == "" && withoutNull(row.Cells["Comp4"].Value) == "")
                                {
                                    row.Cells["Name_Copy"].Value = row.Cells["Name"].Value;
                                    row.Cells["Name_Copy_Source"].Value = row.Cells["Name_Source"].Value;
                                    row.Cells["Comp5"].Value = "-";
                                }
                            }
                        };
                        break;
                }
            }
            catch
            {

            }

            try // Применяем новый фильтр к измененным ячейкам
            {
                frm2.Close();
                frm2Loaded = 0;
            }
            catch
            {
            }
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e) // При ошибочном вводе значение
        {
            dataGridView1.CancelEdit();       
        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e) // Заменяем точку на запятую в качестве разделителя
        {
            if (dataGridView1.Columns[e.ColumnIndex].ValueType == typeof(decimal))
            {
                string str = e.FormattedValue.ToString();
                str = str.Replace('.', ',');
                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = str;
                dataGridView1.EndEdit();
            }
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e) // Перед заливкой ячеек
        {
            DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
            string indexStr = (e.RowIndex + 1).ToString();

            object header = row.HeaderCell.Value; // Проставляем номера строк
            if (header == null || !header.Equals(indexStr))
            {
                row.HeaderCell.Value = indexStr;
            }

            try
            {
                switch (tabName)
                {
                    case "Sheet6":   // Добавляем условное форматирование
                        {
                            if (withoutNull(row.Cells["Price"].Value) == withoutNull(row.Cells["Price_Fin"].Value))
                            {
                                row.Cells["Price_Fin"].Style.ForeColor = Color.FromArgb(0, 0, 255);
                            }
                            else
                            {
                                row.Cells["Price_Fin"].Style.ForeColor = Color.FromArgb(125, 0, 125);
                            }
                        };
                        break;

                    case "Sheet6View":   // Добавляем условное форматирование
                        {
                            if (withoutNull(row.Cells["Price"].Value) == withoutNull(row.Cells["Price_Fin"].Value))
                            {
                                row.Cells["Price_Fin"].Style.ForeColor = Color.FromArgb(0, 0, 255);
                            }
                            else
                            {
                                row.Cells["Price_Fin"].Style.ForeColor = Color.FromArgb(125, 0, 125);
                            }
                        };
                        break;

                    case "PromCode":   // Добавляем пересчет сравнения при подстановке кодов
                        {
                            if (withoutNull(row.Cells["Analog"].Value) != "")
                            {
                                if (withoutNull(row.Cells["Code_Analog1"].Value) == withoutNull(row.Cells["Analog"].Value))
                                {
                                    row.Cells["Comp1"].Value = "-";
                                }
                                else
                                {
                                    row.Cells["Comp1"].Value = "+";
                                }

                                if (withoutNull(row.Cells["Code_Analog2"].Value) != "")
                                {
                                    if (withoutNull(row.Cells["Code_Analog2"].Value) == withoutNull(row.Cells["Analog"].Value))
                                    {
                                        row.Cells["Comp2"].Value = "-";
                                    }
                                    else
                                    {
                                        row.Cells["Comp2"].Value = "+";
                                    }
                                }

                            }

                        };
                        break;
                    case "PromName":   // Добавляем пересчет сравнения при подстановке наименований
                        {
                        };
                        break;
                }
            }
            catch { }
        }

        private void dataGridView1_ColumnAdded(object sender, DataGridViewColumnEventArgs e) // При добавлении колонок
        {
            // Делаем только программную сортировку            
            e.Column.SortMode = DataGridViewColumnSortMode.Programmatic;
            e.Column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            e.Column.DefaultCellStyle.Alignment = getAlignment(e.Column.Name);

            // Применяем стили
            e.Column.DefaultCellStyle.Font = defFont;
            e.Column.HeaderCell.Style.Font = defFont;

            if ((e.Column.Index + 1) % 2 != 0)
            {
                e.Column.DefaultCellStyle.BackColor = Color.FromArgb(230, 255, 253);
                e.Column.HeaderCell.Style.BackColor = Color.FromArgb(13, 255, 237);
            }
            else
            {
                e.Column.DefaultCellStyle.BackColor = Color.FromArgb(255, 254, 213);
                e.Column.HeaderCell.Style.BackColor = Color.FromArgb(240, 240, 0);
            }
        }
        private DataGridViewContentAlignment getAlignment(string ColumnName)
        {
            DataGridViewContentAlignment align = DataGridViewContentAlignment.MiddleCenter;
            string tip = tbl.Columns[ColumnName].DataType.ToString();

            switch (tip)
            {
                case "System.DateTime":
                    align = DataGridViewContentAlignment.MiddleCenter;
                    break;

                case "System.String":
                    {
                        if (ColumnName.IndexOf("Code") > -1 || ColumnName.IndexOf("EI") > -1 || ColumnName == "Dep" || ColumnName == "Obn"
                            || ColumnName.IndexOf("Comp") > -1 || ColumnName == "Ved" || ColumnName.IndexOf("Age") > -1 || ColumnName == "Filter"
                            || ColumnName == "Analog" || ColumnName == "Код" || ColumnName == "ЕИ")
                        {
                            align = DataGridViewContentAlignment.MiddleCenter;
                        }
                        else
                        {
                            align = DataGridViewContentAlignment.MiddleLeft;
                        }
                    };
                    break;

                default:
                    align = DataGridViewContentAlignment.MiddleRight;
                    break;
            }

            return align;
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e) // При нажатии на заголовок столбца
        {
            if (shwSchema == false)
            {
                switch (frm2Loaded)
                {
                    case 0:
                        {
                            colName = dataGridView1.Columns[e.ColumnIndex].Name;
                            switchDataType(dataGridView1.Columns[e.ColumnIndex].ValueType.ToString());
                            createForm2(e.ColumnIndex, e.RowIndex + 1);
                        };
                        break;
                    case 1:
                        {
                            if (colName == dataGridView1.Columns[e.ColumnIndex].Name)
                            {
                                frm2.Show();
                                frm2Loaded = 2;
                            }
                            else
                            {
                                frm2.Close();
                                colName = dataGridView1.Columns[e.ColumnIndex].Name;
                                switchDataType(dataGridView1.Columns[e.ColumnIndex].ValueType.ToString());
                                createForm2(e.ColumnIndex, e.RowIndex + 1);
                            }
                        };
                        break;
                    case 2:
                        {
                            if (colName == dataGridView1.Columns[e.ColumnIndex].Name)
                            {
                                frm2.Hide();
                                frm2.Width = dataGridView1.Columns[e.ColumnIndex].Width;
                                frm2.StartPosition = FormStartPosition.Manual;
                                Point pnt = dataGridView1.PointToScreen(dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex + 1, false).Location);
                                frm2.Location = pnt;
                                frm2Loaded = 1;
                            }
                            else
                            {
                                frm2.Close();
                                colName = dataGridView1.Columns[e.ColumnIndex].Name;
                                switchDataType(dataGridView1.Columns[e.ColumnIndex].ValueType.ToString());
                                createForm2(e.ColumnIndex, e.RowIndex + 1);
                            }
                        }
                        break;
                }
            }
        }

        private void switchDataType(string dt) // Выбираем тип данных
        {
            switch (dt)
            {
                case "System.String":
                    {
                        colIsText = true;
                        colIsDate = false;
                    };
                    break;
                case "System.DateTime":
                    {
                        colIsText = false;
                        colIsDate = true;
                    }
                    break;
                default:
                    {
                        colIsText = false;
                        colIsDate = false;
                    }
                    break;
            }
        }

        public void addFilter(string ColumnName) // Добавляем источник данных для фильтра в DataGridViewWithFilter
        {
            if (formFiltered == true)
            {
                filSource = (new DataView(tbl, filterString, ColumnName + " ASC", DataViewRowState.CurrentRows)).ToTable(true, ColumnName);
            }
            else
            {
                filSource = (new DataView(tbl, "", ColumnName + " ASC", DataViewRowState.CurrentRows)).ToTable(true, ColumnName);
            }


            if (colIsText == false)
            {
                filSource.Columns[0].ColumnName = "Old";
                filSource.Columns.Add(ColumnName, typeof(string)); // Добавляем новую строку с прежним наименованием

                if (colIsDate == true)
                {
                    for (int i = 0; i < filSource.Rows.Count; i++)
                    {
                        if (filSource.Rows[i][0].GetType() == typeof(DBNull))
                        {
                            filSource.Rows[i][1] = "";
                        }
                        else
                        {
                            filSource.Rows[i][1] = Convert.ToDateTime(filSource.Rows[i][0]).ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < filSource.Rows.Count; i++)
                    {
                        if (filSource.Rows[i][0].GetType() == typeof(DBNull))
                        {
                            filSource.Rows[i][1] = "";
                        }
                        else
                        {
                            filSource.Rows[i][1] = filSource.Rows[i][0].ToString();
                        }
                    }

                }

                filSource.Columns.Remove("Old"); // удаляем старую колонку
                EnumerableRowCollection<DataRow> query; // Создаем запрос
                if (colIsDate == true)
                {
                    query = from row in filSource.AsEnumerable() orderby dateSort(row[ColumnName]) select row;
                }
                else
                {
                    query = from row in filSource.AsEnumerable() orderby numSort(row[ColumnName]) select row;
                }

                dw = query.AsDataView();
                filSource = dw.ToTable();
            }
            frm2.filterSource = filSource;
        }

        private DateTime dateSort(object i) // Возвращаем формат отображения даты
        {
            if (i == "")
            {
                return new DateTime(1900, 1, 1);
            }
            else
            {
                return Convert.ToDateTime(i);
            }
        }

        private decimal numSort(object i) // Возвращаем формат отображения числа
        {
            if (i == "")
            {
                return 0;
            }
            else
            {
                return Convert.ToDecimal(i);
            }
        }

        public void createForm2(int ColumnIndex, int RowIndex) // Для создания новой формы 2
        {
            frm2 = new Form2(); // Создаем новую форму

            frm2.StartPosition = FormStartPosition.Manual; // Делаем пользовательскую стартовую позицию
            Point pnt = dataGridView1.PointToScreen(dataGridView1.GetCellDisplayRectangle(ColumnIndex, RowIndex, false).Location); // Задаем начальную позицию формы2

            frm2.frm1 = this; // Задаем данную форму в качестве родительской
            frm2.tableName = tabName; // Задаем имя текущей таблицы
            frm2.colName = dataGridView1.Columns[ColumnIndex].Name; // Задаем имя текущей колонки


            frm2.colIsText = this.colIsText; // Задаем, является ли поле текстовым
            frm2.colIsDate = this.colIsDate; // Задаем, является ли поле датой
            addFilter(frm2.colName); // Задаем источник для фильтра

            frm2.Show();
            frm2Loaded = 2;

            frm2.Location = pnt;
            frm2.Width = dataGridView1.Columns[ColumnIndex].Width; // Делаем длину формы равной ширине столбца
        }

        private void dataGridView1_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e) // При изменении размера столбца
        {
            if (frm2Loaded != 0)
            {
                if (frm2Loaded == 2)
                {
                    frm2.Hide();
                    frm2.Width = e.Column.Width;
                    frm2.Show();
                }
                else
                {
                    frm2.Width = e.Column.Width;
                }
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e) // При изменении выбора
        {
            List<DataGridViewCell> listCells = (from DataGridViewCell cell in dataGridView1.SelectedCells select cell).ToList();
            if (listCells.Count > 0)
            {
                selectedRange[0] = listCells[listCells.Count - 1].RowIndex;
                selectedRange[1] = listCells[listCells.Count - 1].ColumnIndex;
                selectedRange[2] = listCells[0].RowIndex;
                selectedRange[3] = listCells[0].ColumnIndex;
            }
        }

        public void find(string what, string how, string place, bool way, bool withRegister) // Поиск в DataGridView
        {
            int begRow = 0; // Строка начала поиска
            int begColumn = 0; // Столбец начала поиска
            int finRow = 0; // Строка конца поиска
            int finColumn = 0; // Столбец конца поиска
            List<DataGridViewCell> listCells = new List<DataGridViewCell>();
            switch (place)
            {
                case "Текущий столбец":
                    {
                        if (way == false)
                        {
                            begRow = selectedRange[0]; begColumn = selectedRange[1]; finRow = dataGridView1.RowCount - 1; finColumn = selectedRange[1];
                        }
                        else
                        {
                            begRow = 0; begColumn = selectedRange[1]; finRow = selectedRange[0]; finColumn = selectedRange[1];
                        }
                    };
                    break;
                case "Вся таблица":
                    {
                        if (way == false)
                        {
                            begRow = selectedRange[0]; begColumn = selectedRange[1]; finRow = dataGridView1.RowCount - 1; finColumn = dataGridView1.ColumnCount - 1;
                        }
                        else
                        {
                            begRow = 0; begColumn = 0; finRow = selectedRange[0]; finColumn = selectedRange[1];
                        }
                    };
                    break;
                case "Выделенный диапазон":
                    {
                        begRow = selectedRange[0]; begColumn = selectedRange[1]; finRow = selectedRange[2]; finColumn = selectedRange[3];
                    };
                    break;
            }

            if (place != "Выделенный диапазон")
            {

                switch (finColumn - begColumn)
                {
                    case 0:
                        {
                            List<DataGridViewCell> prom1 = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[begColumn].RowIndex >= begRow && row.Cells[begColumn].RowIndex <= finRow select row.Cells[begColumn]).ToList();
                            listCells.AddRange(prom1);
                        };
                        break;
                    case 1:
                        {
                            List<DataGridViewCell> prom1 = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[begColumn].RowIndex >= begRow select row.Cells[begColumn]).ToList();
                            List<DataGridViewCell> prom2 = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[finColumn].RowIndex <= finRow select row.Cells[finColumn]).ToList();
                            listCells.AddRange(prom1);
                            listCells.AddRange(prom2);
                        };
                        break;
                    default:
                        {
                            List<DataGridViewCell> prom1 = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[begColumn].RowIndex >= begRow select row.Cells[begColumn]).ToList();
                            List<DataGridViewCell> prom2 = new List<DataGridViewCell>();
                            for (int i = begColumn + 1; i < finColumn; i++)
                            {
                                List<DataGridViewCell> prom3 = (from DataGridViewRow row in dataGridView1.Rows select row.Cells[i]).ToList();
                                prom2.AddRange(prom3);
                            }

                            List<DataGridViewCell> prom4 = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[finColumn].RowIndex <= finRow select row.Cells[finColumn]).ToList();
                            listCells.AddRange(prom1);
                            listCells.AddRange(prom2);
                            listCells.AddRange(prom4);
                        };
                        break;
                }
            }
            else
            {
                if (way == false)
                {
                    for (int i = begColumn; i < finColumn + 1; i++)
                    {
                        List<DataGridViewCell> prom1 = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[i].RowIndex >= begRow && row.Cells[i].RowIndex <= finRow select row.Cells[i]).ToList();
                        listCells.AddRange(prom1);
                    }
                }
                else
                {
                    listCells.Add(dataGridView1[begColumn, begRow]);
                }
            }

            foundRange = new List<DataGridViewCell>(); // Список найденных ячеек

            switch (how)
            {
                case "Равно":
                    {
                        if (withRegister == false)
                        {
                            foundRange = (from DataGridViewCell cell in listCells where withoutNull(cell.Value).IndexOf(what, StringComparison.CurrentCultureIgnoreCase) == 0 && withoutNull(cell.Value).Length == what.Length select cell).ToList();
                        }
                        else
                        {
                            foundRange = (from DataGridViewCell cell in listCells where withoutNull(cell.Value) == what select cell).ToList();
                        }
                    };
                    break;
                case "Содержит":
                    {
                        if (withRegister == false)
                        {
                            foundRange = (from DataGridViewCell cell in listCells where withoutNull(cell.Value).IndexOf(what, StringComparison.CurrentCultureIgnoreCase) >= 0 select cell).ToList();
                        }
                        else
                        {
                            foundRange = (from DataGridViewCell cell in listCells where withoutNull(cell.Value).IndexOf(what, StringComparison.CurrentCulture) >= 0 select cell).ToList();
                        }
                    };
                    break;
                case "Начинается с":
                    {
                        if (withRegister == false)
                        {
                            foundRange = (from DataGridViewCell cell in listCells where withoutNull(cell.Value).IndexOf(what, StringComparison.CurrentCultureIgnoreCase) == 0 select cell).ToList();
                        }
                        else
                        {
                            foundRange = (from DataGridViewCell cell in listCells where withoutNull(cell.Value).IndexOf(what, StringComparison.CurrentCulture) == 0 select cell).ToList();
                        }
                    };
                    break;
            }

            if (way == true)
            {
                foundRange.Reverse();
            }

            /*
            foreach (DataGridViewCell cell in foundRange)
            {
                cell.Style.BackColor = cvet; 
            }
            */
        }

        public void showFound(int i) // Показать найденные ячейки
        {
            dataGridView1.CurrentCell = foundRange[i];
            cCell = dataGridView1.CurrentCell;
            dataGridView1.ClearSelection();
            foundRange[i].Selected = true;
        }

        public void selectColumn() // Выделить всю колонку
        {
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullColumnSelect;
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView1.Columns[colName].Selected = true;
        }

        private void dataGridView1_CurrentCellChanged(object sender, EventArgs e) // При смене текущей позиции возвращаем как было
        {
            cCell = dataGridView1.CurrentCell;
            if (dataGridView1.SelectionMode == DataGridViewSelectionMode.FullColumnSelect)
            {
                dataGridView1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
                dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            }

        }
        private void button14_Click(object sender, EventArgs e) // Выполнить SQL-запрос
        {
            try
            {
                cmd = new SqlCommand(textBox1.Text, cn);
                cmd.ExecuteNonQuery();
                showSchema(); // Показываем в dataGridView1 список таблиц и столбцов
                showTables(); // Показываем список таблиц и столцов
                MessageBox.Show("Запрос выполнен!", "Сообщение");
            }
            catch
            {
                MessageBox.Show("Недопустимый запрос!", "Сообщение");
            }
        }

        private void button15_Click(object sender, EventArgs e) // Кнопка Выделить цветом
        {
            foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
            {
                cell.Style.BackColor = cvet;
            }
        }

        private void button16_Click(object sender, EventArgs e) // Кнопка Убрать выделение
        {
            foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
            {
                cell.Style = default;
            }
        }

        private void button17_Click(object sender, EventArgs e) // Кнопка Выбрать цвет
        {
            colorDialog1.ShowDialog();
            cvet = colorDialog1.Color;
        }

        private void button18_Click(object sender, EventArgs e) // Кнопка Очистить значения
        {
            if (dataGridView1.ReadOnly == false)
            {
                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    cell.Value = DBNull.Value;
                }
            }
        }

        private void button19_Click(object sender, EventArgs e) // Кнопка Скопир. значения
        {
            try
            {
                Clipboard.SetDataObject(dataGridView1.GetClipboardContent());
            }
            catch
            {
                MessageBox.Show("Данные для копирования не выбраны!", "Диалог");
            }
        }

        private void button20_Click(object sender, EventArgs e) // Кнопка Заполн. значен.
        {
            if (dataGridView1.ReadOnly == false)
            {
                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    cell.Value = filler;
                }
            }
        }

        private void button21_Click(object sender, EventArgs e) // Кнопка Выбрать заполн.
        {
            if (textBox2.Visible == false)
            {
                textBox2.Visible = true;
                textBox2.Text = filler.ToString();
                button21.Text = "Введите значение:";
            }
            else
            {
                filler = textBox2.Text;
                textBox2.Visible = false;
                textBox2.Text = "Выбрать заполн.";                
            }
        }
    }
}
