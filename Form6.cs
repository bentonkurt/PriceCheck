using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient; // Для SQL-сервера
using System.Data.OleDb; // Для OleDb-сервера
using System.Diagnostics; // Для открытия файла
using System.IO; // Для получения расширения

namespace PriceCheck
{
    public partial class Form6 : Form
    {
        public Form1 pf; // родительская форма
        bool iProg; // идентификатор программы (True = Access)
        

        SqlConnection cnSQL; // соедиенение с SQL-сервером        
        string strSQLConn; // строка соедиенениея с SQL-сервером
        string strSQL; // строка SQL
        SqlDataReader rdrSQL; // SQL  DataReader
        SqlCommand cmdSQL; // SQL - команда

        string strOleConn; // строка соедиенениея с OleDb-сервером
        string strOle; // строка Ole
        OleDbConnection cnOle; // соедиенение с OleDb-сервером        
        OleDbCommand cmdOle; // Ole - команда
        OleDbTransaction txn; // Транзакция OleDb

        string oleFile; // Файл для поставщика OleDb

        List<string> spTabEx; // Список таблиц для Экспорта
        List<string[]> spColEx; // Список столбцов для Экспорта

        TreeNode mySelectedNode; // Выбраная нода
        string initText; // Первоначальный текст ноды
        ContextMenu cm = new ContextMenu(); // Контектное меню для вывода таблиц Access
        bool contShowed = false; // Показано ли содержимое файла Access

        public Form6()
        {
            InitializeComponent();

            strSQLConn = @"Data Source=W69\W69SQLEXPRESS;Initial Catalog=CheckPrice;Integrated Security=True";
            cnSQL = new SqlConnection(strSQLConn);
        }
        private void Form3_Load(object sender, EventArgs e) // при загрузке формы
        {
            this.BackColor = Color.FromArgb(240, 189, 170);
            iProg = true; // True = Access
            radioButton1.CheckedChanged += new EventHandler(swCom);
            radioButton2.CheckedChanged += new EventHandler(swCom);
            treeView1.AfterCheck += new TreeViewEventHandler(handleOnTree);

            MenuItem rename = new MenuItem("Переименовать", rename_Click);
            cm.MenuItems.Add(rename);
            MenuItem delete = new MenuItem("Удалить", delete_Click);
            cm.MenuItems.Add(delete);
        }

        private void treeView1_MouseDown(object sender, MouseEventArgs e) // Нужно, чтобы правильно выбиралась нода для контекстного меню
        {
            mySelectedNode = treeView1.GetNodeAt(e.X, e.Y);
            treeView1.SelectedNode = mySelectedNode;
        }

        private void rename_Click(object sender, EventArgs e) // Кнопка контекстного меню "Переименовать"
        {
            mySelectedNode.BeginEdit();
        }

        private void delete_Click(object sender, EventArgs e) // Кнопка контекстного меню "Удалить"
        {
            mySelectedNode.Remove();

            // Открываем подключение
            try
            {
                cnOleOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            txn = cnOle.BeginTransaction();
            strOle = "DROP TABLE " + mySelectedNode.Text;
            cmdOle = new OleDbCommand(strOle, cnOle, txn);
            try
            {
                cmdOle.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                txn.Rollback();
                cnOle.Close();
                return;
            }
            txn.Commit();
            showTables();
            cnOle.Close();
        }

        private void treeView1_BeforeLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            initText = e.Node.Text;
        }

        private void treeView1_AfterLabelEdit(object sender, NodeLabelEditEventArgs e) // Выполняется после переименования
        {
            string finText = e.Label;

            if (finText == "")
            {
                e.CancelEdit = true;
                return;
            }

            
            if (finText != initText && finText != null)
            {
                // Открываем подключение
                try
                {
                    cnOleOpen();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Сообщение");
                    return;
                }

                txn = cnOle.BeginTransaction();
                strOle = "SELECT * INTO [" + finText + "] FROM [" + initText + "]";
                cmdOle = new OleDbCommand(strOle, cnOle, txn);
                OleDbCommand cmdOle1 = new OleDbCommand("DROP TABLE [" + initText + "]", cnOle, txn);
                try
                {
                    cmdOle.ExecuteNonQuery();
                    cmdOle1.ExecuteNonQuery();
                }
                catch
                {
                    txn.Rollback();
                    cnOle.Close();
                    return;
                }
                                
                txn.Commit();
                showTables();
                cnOle.Close();
            }
        }

        private void Form6_FormClosing(object sender, FormClosingEventArgs e)
        {
            pf.frm6Loaded = false;
        }
        private void swCom(object sender, EventArgs e) // Событие для switchCom()
        {
            if (radioButton1.Checked == false)
            {
                iProg = true;

            }
            else
            {
                iProg = false;
            }
        }
        private void execCom(string[] par) // Последовательное выполнение команд SQL
        {
            for (int i = 0; i < par.Length; i++)
            {
                cmdSQL = new SqlCommand(par[i], cnSQL);
                cmdSQL.ExecuteNonQuery();
            }
        }

        private string findType(string p1, string p2) // Находит сответствующий тип данных Access
        {
            string cel = "";
            for (int i = 0; i < spColEx.Count; i++)
            {
                if (p1 == spColEx[i][0] && p2 == spColEx[i][1])
                {
                    cel = "[" + spColEx[i][1] + "]" + " " + spColEx[i][5];
                }
            }
            return cel;
        }

        private void cnSQLOpen() // Открываем соединение SQL
        {
            switch (cnSQL.State)
            {
                case ConnectionState.Closed:
                    cnSQL.Open();
                    break;
                case ConnectionState.Broken:
                    {
                        cnSQL.Close();
                        cnSQL.Open();
                    };
                    break;
                default:
                    { }
                    break;
            }
        }

        private void cnOleOpen() // Открываем соединение OleDb
        {
            switch (cnOle.State)
            {
                case ConnectionState.Closed:
                    cnOle.Open();
                    break;
                case ConnectionState.Broken:
                    {
                        cnOle.Close();
                        cnOle.Open();
                    };
                    break;
                default:
                    { }
                    break;
            }
        }

        private void button1_Click(object sender, EventArgs e) // Кнопка обновить
        {
            obnForExp();
        }

        private void handleOnTree(Object sender, TreeViewEventArgs e) // добавляем ивент на выбор дочерних объектов в дереве
        {
            checkAllNodes(e.Node, e.Node.Checked);
        }

        private void checkAllNodes(TreeNode node, Boolean isChecked)  // выбираем все дочерние объекты в дереве при выборе родительского
        {
            if (isChecked == true)
            {
                foreach (TreeNode item in node.Nodes)
                {
                    item.Checked = isChecked;

                    if (item.Nodes.Count > 0)
                    {
                        this.checkAllNodes(item, isChecked);
                    }
                }
            }
            else
            {
                try
                {
                    node.Parent.Checked = false;
                }
                catch { }
            }
        }
        private void obnForExp() // Вывод таблиц и столбцов для экспорта!
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

            spTabEx = new List<string>();
            spColEx = new List<string[]>(); // Список столбцов для Экспорта

            strSQL = "SELECT * FROM INFORMATION_SCHEMA.TABLES ORDER BY TABLE_NAME"; // Заполняем имена таблиц
            cmdSQL = new SqlCommand(strSQL, cnSQL);
            rdrSQL = cmdSQL.ExecuteReader();

            while (rdrSQL.Read())
            {
                spTabEx.Add((string)rdrSQL["Table_Name"]);
            }
            rdrSQL.Close();

            strSQL = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS ORDER BY TABLE_NAME"; // Заполняем имена и типы данных колонок
            cmdSQL = new SqlCommand(strSQL, cnSQL);
            rdrSQL = cmdSQL.ExecuteReader();

            while (rdrSQL.Read())
            {
                string[] spColExElem = new string[6];
                spColExElem[0] = (string)rdrSQL["TABLE_NAME"];
                spColExElem[1] = (string)rdrSQL["COLUMN_NAME"];
                spColExElem[2] = (string)rdrSQL["DATA_TYPE"];
                switch (spColExElem[2])
                {
                    case "decimal":
                        {
                            spColExElem[3] = System.Convert.ToString(rdrSQL["NUMERIC_PRECISION"]);
                            spColExElem[4] = System.Convert.ToString(rdrSQL["NUMERIC_SCALE"]);
                        };
                        break;
                    case "numeric":
                        {
                            spColExElem[3] = System.Convert.ToString(rdrSQL["NUMERIC_PRECISION"]);
                            spColExElem[4] = System.Convert.ToString(rdrSQL["NUMERIC_SCALE"]);
                        };
                        break;
                    case "char":
                        {
                            spColExElem[3] = System.Convert.ToString(rdrSQL["CHARACTER_MAXIMUM_LENGTH"]);
                            spColExElem[4] = "";
                        };
                        break;
                    case "nchar":
                        {
                            spColExElem[3] = System.Convert.ToString(rdrSQL["CHARACTER_MAXIMUM_LENGTH"]);
                            spColExElem[4] = "";
                        };
                        break;
                    case "nvarchar":
                        {
                            spColExElem[3] = System.Convert.ToString(rdrSQL["CHARACTER_MAXIMUM_LENGTH"]);
                            spColExElem[4] = "";
                        };
                        break;
                    case "varchar":
                        {
                            spColExElem[3] = System.Convert.ToString(rdrSQL["CHARACTER_MAXIMUM_LENGTH"]);
                            spColExElem[4] = "";
                        };
                        break;
                    default:
                        {
                            spColExElem[3] = "";
                            spColExElem[4] = "";
                        };
                        break;
                }
                spColExElem[5] = getDataType(spColExElem[2], spColExElem[3], spColExElem[4]);
                spColEx.Add(spColExElem);
            }
            rdrSQL.Close();

            treeView1.LabelEdit = false; // Запрещаем редактирование нод
            treeView1.CheckBoxes = true; // Добавляем галочки
            treeView1.Nodes.Clear(); // очищаем старые значения
            TreeNode node = new TreeNode("Columns"); // нода в дереве
            treeView1.Nodes.Add(node);

            foreach (string tb in spTabEx)
            {
                TreeNode tn = new TreeNode(tb);
                node.Nodes.Add(tn);
                for (int i = 0; i < spColEx.Count; i++)
                {
                    if (tb == spColEx[i][0])
                    {
                        TreeNode tnc = new TreeNode(spColEx[i][1]);
                        tn.Nodes.Add(tnc);
                    }
                }
            }

            treeView1.Nodes[0].Expand(); // развернуть стволовую ноду
            cnSQL.Close();
        }

        public string getDataType(string p1, string p2, string p3) // Установление соответствия типов Access 
        {
            string cel = "";
            switch (p1)
            {
                case "bigint":
                    cel = "LONGTEXT";
                    break;
                case "binary":
                    cel = "OLEOBJECT";
                    break;
                case "bit":
                    cel = "BIT";
                    break;
                case "char":
                    if (System.Convert.ToInt64(p2) > 255)
                    {
                        cel = "LONGTEXT";
                    }
                    else
                    {
                        cel = "VARCHAR" + "(" + p2 + ")";
                    };
                    break;
                case "date":
                    cel = "DATETIME";
                    break;
                case "datetime":
                    cel = "DATETIME";
                    break;
                case "datetime2":
                    cel = "DATETIME";
                    break;
                case "datetimeoffset":
                    cel = "DATETIME";
                    break;
                case "decimal":
                    cel = "DECIMAL" + "(" + p2 + "," + p3 + ")";
                    break;
                case "FileStream":
                    cel = "OLEOBJECT";
                    break;
                case "float":
                    cel = "DOUBLE";
                    break;
                case "image":
                    cel = "OLEOBJECT";
                    break;
                case "int":
                    cel = "LONG";
                    break;
                case "money":
                    cel = "MONEY";
                    break;
                case "nchar":
                    if (System.Convert.ToInt64(p2) > 255)
                    {
                        cel = "LONGTEXT";
                    }
                    else
                    {
                        cel = "VARCHAR" + "(" + p2 + ")";
                    };
                    break;
                case "ntext":
                    cel = "LONGTEXT";
                    break;
                case "numeric":
                    cel = "DECIMAL" + "(" + p2 + "," + p3 + ")";
                    break;
                case "nvarchar":
                    if (System.Convert.ToInt64(p2) > 255)
                    {
                        cel = "LONGTEXT";
                    }
                    else
                    {
                        cel = "VARCHAR" + "(" + p2 + ")";
                    };
                    break;
                case "real":
                    cel = "SINGLE";
                    break;
                case "smalldatetime":
                    cel = "DATETIME";
                    break;
                case "smallint":
                    cel = "SMALLINT";
                    break;
                case "smallmoney":
                    cel = "MONEY";
                    break;
                case "sql_variant":
                    cel = "LONGTEXT";
                    break;
                case "text":
                    cel = "LONGTEXT";
                    break;
                case "time":
                    cel = "DATETIME";
                    break;
                case "timestamp":
                    cel = "LONGTEXT";
                    break;
                case "tinyint":
                    cel = "BYTE";
                    break;
                case "uniqueidentifier":
                    cel = "UNIQUEIDENTIFIER";
                    break;
                case "varbinary":
                    cel = "OLEOBJECT";
                    break;
                case "varchar":
                    if (System.Convert.ToInt64(p2) > 255)
                    {
                        cel = "LONGTEXT";
                    }
                    else
                    {
                        cel = "VARCHAR" + "(" + p2 + ")";
                    };
                    break;
                case "xml":
                    cel = "LONGTEXT";
                    break;
                default:
                    cel = "OLEOBJECT";
                    break;
            }
            return cel;
        }

        private void button2_Click(object sender, EventArgs e) // Кнопка обзор
        {
            if (contShowed == true)
            {
                oleFile = null;
                label1.Text = "";
                treeView1.Nodes.Clear();
                contShowed = false;
            }

            openFileDialog1.InitialDirectory = @"C:\Мои документы\Дима\Проверка цен";
            openFileDialog1.FileName = "";
            if(iProg == false)
            {
                this.openFileDialog1.Filter = "Книга Excel|*.xls; *.xlsx";
            }
            else
            {
                this.openFileDialog1.Filter = "База данных Access|*.mdb; *.accdb";
            }

            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e) // При выборе файла
        {
            oleFile = openFileDialog1.FileName; // Получаем путь к файлу
            label1.Text = oleFile;
            if (iProg == false)  // Формируем подключение на основе выбора 
            {
                strOleConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + oleFile + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
            }
            else
            {
                strOleConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + oleFile;
            }
            cnOle = new OleDbConnection(strOleConn);
        }

        private void button3_Click(object sender, EventArgs e) // Кнопка "Выполнить"
        {
            if (oleFile != null && oleFile != "")
            {
                // Открываем соединения
                try
                {
                    cnSQLOpen();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Сообщение");
                    return;
                }

                try
                {
                    cnOleOpen();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Сообщение");
                    return;
                }

                List<string[]> lstOle = new List<string[]>();
                string[] lstOleElem;
                List<List<string>> param = new List<List<string>>();
                List<string> paramElem;

                try
                {
                    txn = cnOle.BeginTransaction(); // Начинаем транзакцию

                    string spisTablCr = ""; // Список таблиц для создания
                    string spisTablIns = ""; // Список таблиц для вставки
                    string spisTablSel = ""; // Список таблиц для выбора

                    foreach (TreeNode tr1 in treeView1.Nodes[0].Nodes)
                    {
                        string spisColCr = ""; // Список параметров для создания
                        string spisColIns1 = ""; // Список параметров для вcтавки
                        string spisColIns2 = ""; // Список параметров для вcтавки

                        paramElem = new List<string>();

                        foreach (TreeNode tr2 in tr1.Nodes)
                        {
                            if (tr2.Checked == true)
                            {
                                spisColCr = spisColCr + findType(tr1.Text, tr2.Text) + ", ";
                                spisColIns1 = spisColIns1 + "[" + tr2.Text + "], ";
                                spisColIns2 = spisColIns2 + "@" + tr2.Text + ", ";
                                paramElem.Add("@" + tr2.Text);
                            }
                        }

                        if (spisColCr.Length != 0)
                        {
                            spisColCr = spisColCr.Substring(0, spisColCr.Length - 2);
                            spisTablCr = "CREATE TABLE [" + tr1.Text + "] (" + spisColCr + ")";

                            spisColIns1 = spisColIns1.Substring(0, spisColIns1.Length - 2);
                            spisColIns2 = spisColIns2.Substring(0, spisColIns2.Length - 2);
                            spisTablSel = "SELECT " + spisColIns1 + " FROM " + tr1.Text;
                            spisTablIns = "INSERT INTO [" + tr1.Text + "] (" + spisColIns1 + ") VALUES (" + spisColIns2 + ")";

                            lstOleElem = new string[3];
                            lstOleElem[0] = spisTablCr;
                            lstOleElem[1] = spisTablSel;
                            lstOleElem[2] = spisTablIns;

                            lstOle.Add(lstOleElem);

                            param.Add(paramElem);
                        }
                    }
                    
                    if (spisTablCr.Length != 0)
                    {
                        for (int i = 0; i < lstOle.Count; i++)
                        {
                            cmdOle = new OleDbCommand(lstOle[i][0], cnOle, txn);
                            cmdOle.ExecuteNonQuery();

                            cmdSQL = new SqlCommand(lstOle[i][1], cnSQL);
                            rdrSQL = cmdSQL.ExecuteReader();

                            while (rdrSQL.Read())
                            {
                                cmdOle = new OleDbCommand(lstOle[i][2], cnOle, txn);
                                for (int j = 0; j < rdrSQL.FieldCount; j++)
                                {
                                    if (rdrSQL[j] == DBNull.Value)
                                    {
                                        cmdOle.Parameters.AddWithValue(param[i][j], DBNull.Value);
                                    }
                                    else
                                    {
                                        if (rdrSQL[j].GetType().ToString() == "System.DateTime")
                                        {
                                            cmdOle.Parameters.AddWithValue(param[i][j], Convert.ToDateTime(rdrSQL[j]).ToString("dd.MM.yyyy"));
                                        }
                                        else
                                        {
                                            cmdOle.Parameters.AddWithValue(param[i][j], rdrSQL[j].ToString());
                                        }

                                    }
                                }
                                cmdOle.ExecuteNonQuery();
                            }
                            rdrSQL.Close();
                        }
                        txn.Commit();
                        cnSQL.Close(); // Закрываем соедиения
                        cnOle.Close();
                        MessageBox.Show("Эскпорт выполнен!", "Сообщение");
                    }
                    else 
                    {
                        txn.Rollback();
                        cnSQL.Close(); // Закрываем соедиения
                        cnOle.Close();
                        MessageBox.Show("Не выбраны таблицы для экспорта!", "Сообщение");
                        return;
                    }
                    
                }
                catch (Exception ex)
                {
                    txn.Rollback();
                    cnSQL.Close(); // Закрываем соедиения
                    cnOle.Close();
                    MessageBox.Show(ex.Message, "Сообщение");                    
                    return;
                }
            }
        }
        private void button4_Click(object sender, EventArgs e) // Кнопка Показать содержимое
        {
            if (oleFile != null && oleFile != "")
            {
                if (iProg == false)
                {
                    DialogResult res = MessageBox.Show("Открыть файл или папку?", "Диалог", MessageBoxButtons.YesNo);
                    switch (res)
                    {
                        case DialogResult.Yes:
                            {
                                Process.Start(oleFile); // Открываем сам файл
                            };
                            break;
                        case DialogResult.No:
                            {
                                Process.Start("explorer.exe", "/select," + oleFile); // Открываем содержащую папку
                            };
                            break;
                    }
                }
                else
                {
                    DialogResult res = MessageBox.Show("Вывести содержимое или открыть папку?", "Диалог", MessageBoxButtons.YesNo);
                    switch (res)
                    {
                        case DialogResult.Yes:
                            {
                                try
                                {
                                    cnOleOpen();
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, "Сообщение");
                                    return;
                                }
                                showTables();
                                cnOle.Close();
                            };
                            break;
                        case DialogResult.No:
                            {
                                Process.Start("explorer.exe", "/select," + oleFile); // Открываем содержащую папку
                            };
                            break;
                    }
                }
            }
        }

        private void showTables() // Выводим список таблиц, содержащихся в целевом файле
        {
            treeView1.LabelEdit = true; // Разрешаем редактирование нод
            treeView1.CheckBoxes = false; // Убираем галочки
            treeView1.Nodes.Clear(); // очищаем старые значения
            TreeNode node = new TreeNode("Tables"); // нода в дереве
            treeView1.Nodes.Add(node);

            for (int i = 0; i < cnOle.GetSchema("Tables").Rows.Count; i++)
            {
                if (cnOle.GetSchema("Tables").Rows[i]["TABLE_TYPE"].ToString() == "TABLE")
                {
                    node.Nodes.Add(cnOle.GetSchema("Tables").Rows[i]["TABLE_NAME"].ToString()).ContextMenu = cm;
                }
            }
            treeView1.Nodes[0].Expand(); // развернуть стволовую ноду
            contShowed = true;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            CheckedChanged();
            
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            CheckedChanged();           
        }

        private void CheckedChanged()
        {
            oleFile = null;
            label1.Text = "";
            if (contShowed == true)
            {
                treeView1.Nodes.Clear();
                contShowed = false;
            }
        }

        private void button5_Click(object sender, EventArgs e) // Кнопка Создать копию БД
        {
            saveFileDialog1.DefaultExt = "bak";
            saveFileDialog1.FileName = "CheckPrice";
            saveFileDialog1.InitialDirectory = @"C:\Мои документы\Дима\Проверка цен";
            saveFileDialog1.ShowDialog();
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            // Открываем соединения
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            strSQL = "BACKUP DATABASE CheckPrice TO DISK = '" + saveFileDialog1.FileName + "'";
            cmdSQL = new SqlCommand(strSQL, cnSQL);
            try
            {
                cmdSQL.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Чтобы воспользоваться данной функцией, \nнастройте разрешение для целевой папки!", "Сообщение");
                cnSQL.Close();
                return;
            }
            cnSQL.Close();
            MessageBox.Show("Резервная копия БД создана!", "Сообщение");
        }

        private void button6_Click(object sender, EventArgs e) // Кнопка Сжать БД
        {
            // Открываем соединения
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            strSQL = "DBCC SHRINKDATABASE ('CheckPrice')";
            cmdSQL = new SqlCommand(strSQL, cnSQL);
            try
            {
                cmdSQL.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                cnSQL.Close();
                return;
            }
            cnSQL.Close();
            MessageBox.Show("Сжатие БД выполнено!", "Сообщение");
        }

        private void button7_Click(object sender, EventArgs e) // Кнопка Восст. БД из файла
        {
            openFileDialog2.InitialDirectory = @"C:\Мои документы\Дима\Проверка цен";
            openFileDialog2.FileName = "";
            openFileDialog2.Filter = "Файл резевной копии БД|*.bak";
            openFileDialog2.ShowDialog();
        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            // Открываем соединения
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            strSQL = "USE MASTER RESTORE DATABASE [CheckPrice] FROM DISK = '" + openFileDialog2.FileName + "' WITH REPLACE"; // Если не использовать USE MASTER, выдаст ошибку; WITH REPLACE - с удалением предыдущей базы

            cmdSQL = new SqlCommand(strSQL, cnSQL);
            try
            {
                cmdSQL.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                cnSQL.Close();
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }
            cnSQL.Close();
            MessageBox.Show("Восстановление БД выполнено!", "Сообщение");
        }
    }
}
