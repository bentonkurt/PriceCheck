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
using Excel = Microsoft.Office.Interop.Excel;
using System.IO; // Для получения расширения

namespace PriceCheck
{
    public partial class Form4 : Form
    {
        public Form1 pf; // родительская форма
        int iOper; // идентификатор программы
        string curTabl; // выбраная таблица

        string strSQLConn; // строка соедиенениея с SQL-сервером
        SqlConnection cnSQL; // соедиенение с SQL-сервером
        string strSQL; // строка SQL
        SqlCommand cmdSQL; // SQL - команда

        List<string> lstN1; // Список номеров проверок
        List<string> lstN2; // Список номеров проверок

        SqlDataReader rdrSQL; // SQL  DataReader
        public string dobavka; // Добавка к выбору позиций для проверки
        SqlTransaction txn; // Транзакция

        string strOleConn; // строка соедиенениея с OleDb-сервером
        string oleFile; // Файл для поставщика OleDb
        OleDbConnection cnOle; // соедиенение с OleDb-сервером
        OleDbCommand cmdOle; // Ole - команда
        OleDbDataReader rdrOle; // Ole  DataReader
                       

        public Form4()
        {
            InitializeComponent();

            // Добавляем имена таблиц
            ComboBox cbc = comboBox1;
            cbc.Items.Add("PriceBuh");
            cbc.Items.Add("VedCenPr");
            cbc.Items.Add("Prov");
            cbc.Items.Add("Sheet4");
            cbc.Items.Add("Sheet5");
            cbc.Items.Add("Sheet6");
            cbc.Items.Add("OKEI");

            // Присваиваем значение идентификатору программы
            iOper = 0;

            // Присваиваем обработчик событий RadioButton'ам
            radioButton1.CheckedChanged += new EventHandler(swCom);
            radioButton2.CheckedChanged += new EventHandler(swCom);

            // Создаем SQL-подключение
            strSQLConn = @"Data Source=W69\W69SQLEXPRESS;Initial Catalog=CheckPrice;Integrated Security=True";
            cnSQL = new SqlConnection(strSQLConn);
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            this.BackColor = Color.FromArgb(240, 189, 170);
        }

        private void Form4_FormClosing(object sender, FormClosingEventArgs e)
        {
            pf.frm4Loaded = false;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) // При выборе имени в ComboBox
        {
            curTabl = comboBox1.SelectedItem.ToString();
        }

        private void switchCom() // Выбираем вариант действий
        {
            if (radioButton1.Checked == true)
            {
                iOper = 0;
            }
            else
            {
                iOper = 1;
            }
        }

        private void swCom(object sender, EventArgs e) // Событие для switchCom()
        {
            switchCom();
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

        private void button1_Click(object sender, EventArgs e) // Кнопка "Обновить"
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
            catch
            {
                MessageBox.Show("Не выбран файл для подключения!", "Сообщение");
                return;
            }


            txn = cnSQL.BeginTransaction(); // Начинаем транзакцию

            try // Пытаемся удалить таблицу
            {
                execSQLCmd("DROP TABLE " + curTabl);
            }
            catch
            {
            }

            // Создаем ее заново
            execSQLCmd(comChoice());

            if (iOper == 0)
            {
                cmdOle = new OleDbCommand("SELECT * FROM [" + curTabl + "$]", cnOle);
            }
            else
            {
                cmdOle = new OleDbCommand("SELECT * FROM [" + curTabl + "]", cnOle);
            }

            try // Пробуем выполнить ридер
            {
                rdrOle = cmdOle.ExecuteReader();
            }
            catch
            {
                txn.Rollback();
                cnSQL.Close();
                cnOle.Close();
                MessageBox.Show("Недопутимое имя таблицы!", "Сообщение");
                return;
            }

            while (rdrOle.Read())
            {
                cmdSQL = new SqlCommand("", cnSQL, txn);
                string fld = ""; // Текущее поле
                string param = ""; // Текущий параметр
                string listFields = ""; // Список полей
                string listParams = ""; // Список параметров

                for (int j = 0; j < rdrOle.FieldCount; j++)
                {
                    if (rdrOle[j] == null)
                    {
                    }
                    else
                    {
                        fld = "[" + rdrOle.GetName(j) + "]";
                        param = "@" + rdrOle.GetName(j);
                        listFields = listFields + fld + ", ";
                        listParams = listParams + param + ", ";
                        cmdSQL.Parameters.AddWithValue(param, rdrOle[j]);
                    }
                }
                listFields = listFields.Substring(0, listFields.Length - 2);
                listParams = listParams.Substring(0, listParams.Length - 2);
                strSQL = "INSERT INTO [" + curTabl + "] (" + listFields + ") VALUES (" + listParams + ")";
                cmdSQL.CommandText = strSQL;
                cmdSQL.ExecuteNonQuery();
            }
            rdrOle.Close();

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соединения
            cnSQL.Close();
            cnOle.Close();
            MessageBox.Show("Таблица успешно экспортирована!", "Сообщение");
        }

        private void execSQLCmd(string str) // Выполнить SQL-комманду
        {
            strSQL = str;
            cmdSQL = new SqlCommand(strSQL, cnSQL, txn);
            cmdSQL.ExecuteNonQuery();
        }

        private string comChoice() // Выбираем команду при выборе значения ComboBox'a
        {
            string Choice = "";
            switch (curTabl)
            {
                case "PriceBuh":
                    {
                        Choice = "CREATE TABLE [PriceBuh] ([Storage] INT, [N_Doc] INT, [Date] DATE, [Code] VARCHAR(255), [Name_IVC] VARCHAR(255), [Code_Buh] BIGINT, " +
                                "[Name_Buh] VARCHAR(255), [EI_Buh] VARCHAR(10), [EI_IVC] VARCHAR(10), [QT] DECIMAL(28,14), [Price] DECIMAL(28,2), [Mistake] VARCHAR(255), " +
                                "[Provider] VARCHAR(255))";
                    };
                    break;
                case "VedCenPr":
                    {
                        Choice = "CREATE TABLE [VedCenPr] ([Storage] INT, [Code_Buh] BIGINT, [Name_Buh] VARCHAR(255), [Provider_Code] INT, [Date] DATE, [EI_Buh] VARCHAR(10), " +
                                 "[Acc_Price] DECIMAL(28,2), [Pur_Price] DECIMAL(28,2), [Dev] DECIMAL(28,2))"; 
                    };
                    break;
                case "Prov":
                    {
                        Choice = "CREATE TABLE [Prov] ([Code] VARCHAR(255), [Name] VARCHAR(255), [Dep] VARCHAR(3), [EI] VARCHAR(3), [Norm] DECIMAL(28,14), " +
                            "[Price] DECIMAL(28,2), [Amount] DECIMAL(28,2), [Product] VARCHAR(255))";
                    };
                    break;
                case "Sheet4":
                    {
                        Choice = "CREATE TABLE [Sheet4] ([Code] VARCHAR(255), [Name_Sub] VARCHAR(255), [EI_Sub] VARCHAR(10), [Price_Sub] DECIMAL(28,2), [Date_Sub] DATE)";
                    }
                    break;
                case "Sheet5":
                    {
                        Choice = "CREATE TABLE [Sheet5] ([Code] VARCHAR(255), [Name] VARCHAR(255), [Name_Source] VARCHAR(255), [EI] VARCHAR(10), [Price_Fin] DECIMAL(28,2), " +
                            "[Date_Fin] DATE, [Obn] VARCHAR(1), [Code_Analog1] VARCHAR(255), [Code_Analog2] VARCHAR(255), [DocInf] VARCHAR(255), [Note] VARCHAR(255))";
                    }
                    break;
                case "Sheet6":
                    {
                        Choice = "CREATE TABLE [Sheet6] ([Code] VARCHAR(255), [Name] VARCHAR(255), [Dep] VARCHAR(3), [EI] VARCHAR(255), [Norm] DECIMAL(28,14), " +
                            "[Price] DECIMAL(28,2), [Amount] DECIMAL(28,2), [Product] VARCHAR(255), [N] TINYINT, [Price_Svod] DECIMAL(28,2), [Date_Svod] DATE, " +
                            "[Source] VARCHAR (255), [Price_Sub] DECIMAL(28,2), [Price_Fin] DECIMAL(28,2), [Ved] VARCHAR (1), [Price_Sub_Age] VARCHAR (1), " +
                            "[Code_Analog] VARCHAR(255), [Amount_Svod] DECIMAL(28,2), [Amount_Fin] DECIMAL(28,2), [Cost] DECIMAL(28,2), " +
                            "[Cost_Svod] DECIMAL(28,2), [Cost_Fin] DECIMAL(28,2), [Rate] DECIMAL(28,2), [Rate_Svod] DECIMAL(28,2), [Rate_Fin] DECIMAL(28,2), " +
                            "[Filter] VARCHAR (1), [Comp_6_13] VARCHAR (1), [Comp_9_13] VARCHAR (1), [EI_Text] VARCHAR (10))";
                    }
                    break;

                case "OKEI":
                    {
                        Choice = "CREATE TABLE [OKEI] ([EI_Text] VARCHAR(10), [EI] VARCHAR(3))";
                    }
                    break;
            }
            return Choice;
        }

        private void button2_Click(object sender, EventArgs e) // Кнопка "Обзор.."
        {
            openFileDialog1.InitialDirectory = @"C:\Мои документы\Дима\Проверка цен";
            openFileDialog1.FileName = "";
            switch (iOper)
            {
                case 0:
                    openFileDialog1.Filter = "Книга Excel|*.xls; *.xlsx";
                    break;
                case 1:
                    openFileDialog1.Filter = "База данных Access|*.mdb; *.accdb";
                    break;
            }
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e) // При выборе файла
        {
            oleFile = openFileDialog1.FileName; // Получаем путь к файлу
            checkBox1.Checked = true;
            if (iOper == 0)  // Формируем подключение на основе выбора 
            {
                strOleConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + oleFile + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
            }
            else
            {
                strOleConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + oleFile;
            }
            cnOle = new OleDbConnection(strOleConn);
        }

        private void checkBox1_Click(object sender, EventArgs e) // При нажатии на checkBox выводим сообщение с выбранным файлом
        {
            if (checkBox1.Checked == true)
                MessageBox.Show(openFileDialog1.FileName,"Выбранный файл");
        }

        private void button3_Click(object sender, EventArgs e) // Подг. лист2
        {
            // Открываем соединение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction(); 

            try // Пытаемся удалить таблицу Sheet2
            {
                execSQLCmd("DROP TABLE Sheet2");
            }
            catch
            {
            }

            // Из PriceBuh выбираем поля с ненулевой ценой и делаем группировку по максимальной дате
            // Результат записываем в таблицу Sheet21
            strSQL = "WITH Prom AS (SELECT * FROM [PriceBuh] WHERE [Price] > 0) " +
                     "SELECT [Code], MAX([Date]) AS [Dates] INTO [Sheet21] FROM Prom GROUP BY Code; " +

            // Из PriceBuh (и Sheet21) выбираем поля Код, Цена и Дата и делаем группировку по максимальной Цене и Дате
            // Результат записываем в таблицу Sheet22
            "WITH Prom AS (SELECT [PriceBuh].[Code], [PriceBuh].[Price], [PriceBuh].[Date] FROM [PriceBuh], [Sheet21] " +
            "WHERE [PriceBuh].[Code] = [Sheet21].[Code] AND [PriceBuh].[Date] = [Sheet21].[Dates]) " +
            "SELECT [Code], MAX([Price]) As [Prices], MAX([Date]) As [Dates] INTO [Sheet22] FROM Prom GROUP BY [Code]; " +

            // Из PriceBuh (и Sheet22) выбираем поля Код, Цена, Дата, N Документа и делаем группировку по максимальной Цене, Дате и N Док
            //Результат записываем в таблицу Sheet23
            "WITH Prom AS (SELECT [PriceBuh].[Code], [PriceBuh].[Price], [PriceBuh].[Date], [PriceBuh].[N_Doc] FROM [PriceBuh], [Sheet22] " +
            "WHERE [PriceBuh].[Code] = [Sheet22].[Code] AND [PriceBuh].[Price] = [Sheet22].[Prices] AND [PriceBuh].[Date] = [Sheet22].[Dates]) " +
            "SELECT [Code], MAX([Price]) As [Prices], MAX([Date]) As [Dates], MAX([N_Doc]) As [Docs] INTO [Sheet23] FROM Prom GROUP BY [Code]; " +

            // Из PriceBuh (и Sheet23) выбираем поля Код, Цена, Дата, N Документа, Кол-во и делаем группировку по максимальной Цене, Дате, N Док и Кол-ву
            //Результат записываем в таблицу Sheet24
            "WITH Prom AS (SELECT [PriceBuh].[Code], [PriceBuh].[Price], [PriceBuh].[Date], [PriceBuh].[N_Doc], [PriceBuh].[QT] FROM [PriceBuh], [Sheet23] " +
            "WHERE [PriceBuh].[Code] = [Sheet23].[Code] AND [PriceBuh].[Price] = [Sheet23].[Prices] AND [PriceBuh].[Date] = [Sheet23].[Dates] AND [PriceBuh].[N_Doc] = [Sheet23].[Docs]) " +
            "SELECT [Code], MAX([Price]) As [Prices], MAX([Date]) As [Dates], MAX([N_Doc]) As [Docs], MAX([QT]) As [QTs] INTO [Sheet24] FROM Prom GROUP BY [Code]; " +

            // Из PriceBuh (и Sheet24) выбираем уникальные значения из всех полей. Результат записываем в таблицу Sheet2
            "SELECT DISTINCT [PriceBuh].[Code], [PriceBuh].[Name_IVC], [PriceBuh].[EI_IVC], [PriceBuh].[Price] AS [Price_Base], [PriceBuh].[Date] AS [Date_Base], " +
            "[PriceBuh].[Code_Buh], [PriceBuh].[Name_Buh], [PriceBuh].[EI_Buh], [PriceBuh].[Storage], [PriceBuh].[N_Doc], [PriceBuh].[QT], [PriceBuh].[Provider] INTO [Sheet2] " +
            "FROM [PriceBuh], [Sheet24] WHERE [PriceBuh].[Code] = [Sheet24].[Code] AND [PriceBuh].[Price] = [Sheet24].[Prices] AND [PriceBuh].[Date] = [Sheet24].[Dates] AND " +
            "[PriceBuh].[N_Doc] = [Sheet24].[Docs] AND [PriceBuh].[QT] = [Sheet24].[QTs]; " +

            //Удаляем временные таблицы
            "DROP TABLE [Sheet21], [Sheet22], [Sheet23], [Sheet24]; " +

            //Добавляем строку Информация о документе и заполняем
            "ALTER TABLE [Sheet2] ADD [DocInf] VARCHAR(255); " +
            "UPDATE [Sheet2] SET [DocInf] = '№' + CAST([N_Doc] AS VARCHAR(255)) + ' от ' +  CONVERT(VARCHAR(255), [Date_Base], 104) + ', ' + Provider";

            try
            {
                // Выполняем полученую команду
                execSQLCmd(strSQL);
            }
            catch
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show("Таблица PriceBuh не существует!", "Сообщение");
                return;
            }


            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();

            MessageBox.Show("Лист2 обновлен!", "Сообщение");
        }

        private void button4_Click(object sender, EventArgs e) // Подг. лист3
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            try // Пытаемся удалить таблицу Sheet3
            {
                execSQLCmd("DROP TABLE Sheet3");
            }
            catch
            {
            }

            // Создаем временную таблицу Sheet31, куда добавляем все поля, где Цена покупная или Цена учетная ненулевые
            strSQL = "SELECT * INTO [Sheet31] FROM [VedCenPr] WHERE [Acc_Price] > 0 OR [Pur_Price] > 0; " +

            // В Sheet31 делаем группировку по максимальной дате. Результат записываем в таблицу Sheet32
            "SELECT [Code_Buh], MAX([Date]) As [Dates] INTO [Sheet32] FROM [Sheet31] GROUP BY [Code_Buh]; " +

            // В Sheet33 добавляем поля из Sheet31, соответсвующие критерию группировки
            "SELECT [Sheet31].* INTO [Sheet33] FROM [Sheet31], [Sheet32] WHERE [Sheet31].[Code_Buh] = [Sheet32].[Code_Buh] AND [Sheet31].[Date] = [Sheet32].[Dates]; " +

            // В таблицу Sheet33 добавляем поле Код, Цена базы, Цена группировки
            "ALTER TABLE [Sheet33] ADD [Code] VARCHAR(255), [Price_Base] DECIMAL(28,2), [Price_grup] DECIMAL(28,2); " +

            // Заполняем поле Код форматированием поля Код бух.
            "UPDATE [Sheet33] SET [Code] = CONVERT(VARCHAR(255), FORMAT ([Code_Buh], '000000-########')); " +

            //Заполняем поле Цена базы ценами из Sheet2, в пустые строчки ставим нули
            //"UPDATE [Sheet2], [Sheet33] SET [Sheet33].[Price_Base] = [Sheet2].[Price_Base] WHERE [Sheet33].[Code] = [Sheet2].[Code]; " +
            "UPDATE [Sheet33] SET [Sheet33].[Price_Base] = [Sheet2].[Price_Base] FROM [Sheet2] WHERE [Sheet33].[Code] = [Sheet2].[Code]; " +

            "UPDATE [Sheet33] SET [Price_Base] = 0 WHERE [Price_Base] IS NULL; " +

            //Заполняем поле Цена группировки, где Покупная или Учетная цена = Цене базы ставим Цену базы, иначе 0
            //в пустые строчки ставим нули
            "UPDATE [Sheet33] SET [Price_grup] = [Price_Base] WHERE [Acc_Price] = [Price_Base] OR [Pur_Price] = [Price_Base]; " +
            "UPDATE [Sheet33] SET [Price_grup] = 0 WHERE [Price_grup] IS NULL; " +

            //Создаем таблицу Sheet34 как группировку по всем полям
            "SELECT [Code], MAX([Storage]) AS [Storages], MAX([Code_Buh]) AS [Code_Buhs], MAX([Name_Buh]) AS [Name_Buhs], " +
            "MAX([Provider_Code]) AS [Provider_Codes], MAX([Date]) AS [Dates], MAX([EI_Buh]) AS [EI_Buhs], " +
            "MAX([Acc_Price]) AS [Acc_Prices], MAX([Pur_Price]) AS [Pur_Prices], MAX([Price_Base]) AS [Price_Bases], MAX([Price_grup]) AS [Price_grups] " +
            "INTO [Sheet34] FROM [Sheet33] GROUP BY [Code]; " +
            
            //В таблицу Sheet34 добавляет Цена бухг.
            "ALTER TABLE [Sheet34] ADD [Price_Buh] DECIMAL(28,2); " +

            //Заполняем поле Цена бухг

            //Где Цена группировки > 0, ставим в результирующую колонку Цену группировки
            "UPDATE [Sheet34] SET [Price_Buh] = [Price_grups] WHERE [Price_grups] > 0; " +

            //Где Цена группировки = 0, а Покупная цена > 0 ставим в результирующую колонку Покупную цену
            "UPDATE [Sheet34] SET [Price_Buh] = [Pur_Prices] WHERE [Price_grups] = 0 AND [Pur_Prices] > 0; " +

            //Где Цена группировки и Покупная цена = 0 ставим в результирующую колонку Учетную цену
            "UPDATE [Sheet34] SET [Price_Buh] = [Acc_Prices] WHERE [Price_grups] = 0 AND [Pur_Prices] = 0; " +

            //Создаем окончательную таблицу Sheet3
            "SELECT [Code], [Storages] AS [Storage], [Provider_Codes] AS [Provider_Code], [Code_Buhs] AS [Code_Buh], " +
            "[Name_Buhs] AS [Name_Buh], [EI_Buhs] AS [EI_Buh], [Acc_Prices] AS [Acc_Price], [Pur_Prices] AS [Pur_Price], " +
            "[Price_Buh], [Dates] AS [Date_Buh] INTO [Sheet3] FROM [Sheet34]; " +

            //Удаляем временные таблицы
            "DROP TABLE [Sheet31], [Sheet32], [Sheet33], [Sheet34]";

            try
            {
                // Выполняем полученую команду
                execSQLCmd(strSQL);
            }
            catch
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show("Таблица VedCenPr не существует!", "Сообщение");
                return;
            }


            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Лист3 обновлен!", "Сообщение");
        }

        private void button5_Click(object sender, EventArgs e) // Подг. пров.
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            //Выбираем, заменить ли существующие цены
            DialogResult vybor = MessageBox.Show("Перезаписать существующие значения?","Диалог",MessageBoxButtons.YesNoCancel);

            switch (vybor)
            {
                case DialogResult.Yes:
                    {
                        try
                        {
                            execSQLCmd("DROP TABLE [Prover]");
                        }
                        catch
                        {
                        }

                        strSQL = "SELECT * INTO [Prom] FROM [Prov]; ALTER TABLE [Prom] ADD [N] TINYINT, [Prom_Dep] VARCHAR(255), [Prom_EI] VARCHAR(3); " +
                            "UPDATE [Prom] SET [N] = 1, [Prom_Dep] = CONVERT(VARCHAR(3), FORMAT(CONVERT(INT, [Dep]), '000')), [Prom_EI] = CONVERT(VARCHAR(3), FORMAT(CONVERT(INT, [EI]), '000')); " +
                            "SELECT [Code], [Name], [Prom_Dep] AS [Dep], [Prom_EI] AS [EI], [Norm], [Price], [Amount], [Product], [N] INTO [Prover] FROM [Prom]; DROP TABLE [Prom]";

                        try
                        {
                            execSQLCmd(strSQL); // Выполняем команду
                        }
                        catch
                        {
                            txn.Rollback();
                            cnSQL.Close();
                            MessageBox.Show("Таблица Prov не существует!", "Сообщение");
                            return;
                        }                        
                    };
                    break;
                case DialogResult.No:
                    {
                        try
                        {
                            strSQL = "SELECT TOP 1 [Code] FROM [Prover]";
                            execSQLCmd(strSQL);


                            strSQL = "SELECT * INTO [Prom] FROM [Prov]; ALTER TABLE [Prom] ADD [Prom_EI] VARCHAR(255), [EI_Text] VARCHAR(255), [Prom_Dep] VARCHAR(3); " +
                                "UPDATE [Prom] SET [Prom_EI] = CONVERT(VARCHAR(3), FORMAT(CONVERT(INT, [EI]), '000')), [Prom_Dep] = CONVERT(VARCHAR(3), FORMAT(CONVERT(INT, [Dep]), '000')); " +
                                "INSERT INTO [Prover] ([Code], [Name], [Dep], [EI], [Norm], [Price], [Amount], [Product]) " +
                                "SELECT [Code], [Name], [Prom_Dep], [Prom_EI], [Norm], [Price], [Amount], [Product] FROM [Prom]; " +
                                "UPDATE [Prover] SET [N] = (SELECT MAX([N]) FROM [Prover]) + 1 WHERE [N] IS NULL; " +
                                "DROP TABLE [Prom]";

                            execSQLCmd(strSQL);
                        }

                        catch
                        {
                            strSQL = "SELECT * INTO [Prom] FROM [Prov]; ALTER TABLE [Prom] ADD [N] TINYINT, [Prom_Dep] VARCHAR(255), [Prom_EI] VARCHAR(3); " +
                            "UPDATE [Prom] SET [N] = 1, [Prom_Dep] = CONVERT(VARCHAR(3), FORMAT(CONVERT(INT, [Dep]), '000')), [Prom_EI] = CONVERT(VARCHAR(3), FORMAT(CONVERT(INT, [EI]), '000')); " +
                            "SELECT [Code], [Name], [Prom_Dep] AS [Dep], [Prom_EI] AS [EI], [Norm], [Price], [Amount], [Product], [N] INTO [Prover] FROM [Prom]; DROP TABLE [Prom]";

                            try
                            {
                                execSQLCmd(strSQL); // Выполняем команду
                            }
                            catch
                            {
                                txn.Rollback();
                                cnSQL.Close();
                                MessageBox.Show("Таблица Prov не существует!", "Сообщение");
                                return;
                            }
                        }
                    };
                    break;
                case DialogResult.Cancel:
                    {
                        txn.Rollback();
                        cnSQL.Close();
                        return;
                    };
                    break;
            }

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Список проверки обновлен!", "Сообщение");
        }

        private void button6_Click(object sender, EventArgs e) // Свод. табл.
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            try
            {
                Svod();
            }
            catch (Exception ex)
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }            

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Сводная таблица сформирована!", "Сообщение");
        }

        private void Svod () // Формирует сводную таблицу
        {
            try // Пытаемся удалить таблицу Svod и представление
            {                
                execSQLCmd("DROP VIEW [SvodView]");
            }
            catch
            {
            }

            try 
            {
                execSQLCmd("DROP TABLE [Svod]");
            }
            catch
            {
            }

            // Создаем сводную таблицу, берем все строки из Sheet5
            strSQL = "SELECT * INTO [Svod] FROM [Sheet5]; " +

            // Добавляем новые значения из Sheet2
            "INSERT INTO [Svod] ([Code], [Name], [EI], [Price_Fin], [Date_Fin]) " +
            "SELECT [Sheet2].[Code], [Sheet2].[Name_IVC], [Sheet2].[EI_IVC], [Sheet2].[Price_Base], [Sheet2].[Date_Base] " +
            "FROM [Sheet2] LEFT JOIN [Svod] ON [Sheet2].[Code] = [Svod].[Code] WHERE [Svod].[Code] IS NULL; " +
            // Проставляем источник наименования
            "UPDATE [Svod] SET [Name_Source] = 'база' WHERE [Name_Source] IS NULL; " +

            // Добавляем новые значения из Sheet3
            "INSERT INTO [Svod] ([Code], [Name], [EI], [Price_Fin], [Date_Fin]) " +
            "SELECT [Sheet3].[Code], [Sheet3].[Name_Buh], [Sheet3].[EI_Buh], [Sheet3].[Price_Buh], [Sheet3].[Date_Buh] " +
            "FROM [Sheet3] LEFT JOIN [Svod] ON [Sheet3].[Code] = [Svod].[Code] WHERE [Svod].[Code] IS NULL; " +
            // Проставляем источник наименования
            "UPDATE [Svod] SET [Name_Source] = 'ведомость' WHERE [Name_Source] IS NULL; " +

            // Добавляем новые значения из Sheet4
            "INSERT INTO [Svod] ([Code], [Name], [EI], [Price_Fin], [Date_Fin]) " +
            "SELECT [Sheet4].[Code], [Sheet4].[Name_Sub], [Sheet4].[EI_Sub], [Sheet4].[Price_Sub], [Sheet4].[Date_Sub] " +
            "FROM [Sheet4] LEFT JOIN [Svod] On [Sheet4].[Code] = [Svod].[Code] WHERE [Svod].[Code] IS NULL; " +
            // Проставляем источник наименования
            "UPDATE [Svod] SET [Name_Source] = 'подстановка' WHERE [Name_Source] IS NULL; " +

            // Добавляем новые поля Цена базы и Дата базы
            "ALTER TABLE [Svod] ADD [Price_Base] DECIMAL(28,2), [Date_Base] DATE; " +
            // Заполняем их
            "UPDATE [Svod] SET [Svod].[Price_Base] = [Sheet2].[Price_Base], [Svod].[Date_Base] = [Sheet2].[Date_Base] FROM [Sheet2] WHERE [Svod].[Code] = [Sheet2].[Code]; " +

            // Добавляем новые поля Цена бухг. и Дата бухг.
            "ALTER TABLE [Svod] ADD [Price_Buh] DECIMAL(28,2), [Date_Buh] DATE; " +
            // Заполняем их
            "UPDATE [Svod] SET [Svod].[Price_Buh] = [Sheet3].[Price_Buh], [Svod].[Date_Buh] = [Sheet3].[Date_Buh] FROM [Sheet3] WHERE [Svod].[Code] = [Sheet3].[Code]; " +

            // Добавляем новые поля Цена результирующая и Дата результирующая
            "ALTER TABLE [Svod] ADD [Price_Res] DECIMAL(28,2), [Date_Res] DATE; " +
            // Заполняем их
            // Если Цена базы = Цене бухг, выбираем эту цену и наибольшую дату
            "UPDATE [Svod] SET [Price_Res] = [Price_Base], [Date_Res] = [Date_Base] WHERE [Price_Base] = [Price_Buh] AND [Date_Base] > [Date_Buh]; " +
            "UPDATE [Svod] SET [Price_Res] = [Price_Buh], [Date_Res] = [Date_Buh] WHERE [Price_Base] = [Price_Buh] AND [Date_Base] <= [Date_Buh]; " +
            // Если Цена базы <> Цене бухг, берем дату и соответ. цену если Дата бух. на месяц больше
            "UPDATE [Svod] SET [Price_Res] = [Price_Base], [Date_Res] = [Date_Base] WHERE [Price_Base] <> [Price_Buh] AND DATEDIFF(d, [Date_Base], [Date_Buh]) <= 30 " +
            "UPDATE [Svod] SET [Price_Res] = [Price_Buh], [Date_Res] = [Date_Buh] WHERE [Price_Base] <> [Price_Buh] AND DATEDIFF(d, [Date_Base], [Date_Buh]) > 30; " +
            // Если одно из полей пустое, то берем другое поле
            "UPDATE [Svod] SET [Price_Res] = [Price_Base], [Date_Res] = [Date_Base] WHERE [Price_Buh] IS NULL; " +
            "UPDATE [Svod] SET [Price_Res] = [Price_Buh], [Date_Res] = [Date_Buh] WHERE [Price_Base] IS NULL; " +

            // Добавляем новые поля Цена подстановки и Дата подстановки
            "ALTER TABLE [Svod] ADD [Price_Sub] DECIMAL(28,2), [Date_Sub] DATE; " +
            // Заполняем их
            "UPDATE [Svod] SET [Svod].[Price_Sub] = [Sheet4].[Price_Sub], [Svod].[Date_Sub] = [Sheet4].[Date_Sub] FROM [Sheet4] WHERE [Svod].[Code] = [Sheet4].[Code]; " +

            // Добавляем новые поля Цена реальная и Дата реальная
            "ALTER TABLE [Svod] ADD [Price_Real] DECIMAL(28,2), [Date_Real] DATE; " +
            //Заполняем их
            // Результат выбираем данные из результирующих и подстановочных данных в соответствии с макс. датой
            "UPDATE [Svod] SET [Price_Real] = [Price_Res], [Date_Real] = [Date_Res] WHERE [Date_Res] >= [Date_Sub]; " +
            "UPDATE [Svod] SET [Price_Real] = [Price_Sub], [Date_Real] = [Date_Sub] WHERE [Date_Res] < [Date_Sub]; " +
            // Если одно из полей пустое, то берем другое поле
            "UPDATE [Svod] SET [Price_Real] = [Price_Res], [Date_Real] = [Date_Res] WHERE [Price_Sub] IS NULL; " +
            "UPDATE [Svod] SET [Price_Real] = [Price_Sub], [Date_Real] = [Date_Sub] WHERE [Price_Res] IS NULL; " +

            // Добавляем поля Дата и Цена для аналоговых значений
            "ALTER TABLE [Svod] ADD [Price_Analog1] DECIMAL(28,2), [Date_Analog1] DATE, [Price_Analog2] DECIMAL(28,2), [Date_Analog2] DATE, " +
            "[Price_Analog] DECIMAL(28,2), [Date_Analog] DATE, [DocInfBase] VARCHAR(255), [DocInfAn1] VARCHAR(255), [DocInfAn2] VARCHAR(255), [DocInfFin] VARCHAR(255); " +

            // Заполняем поле Инф о документе базов
            "UPDATE [Svod] SET [Svod].[DocInfBase] = [Sheet2].[DocInf] FROM [Sheet2] WHERE [Svod].[Code] = [Sheet2].[Code]; " +

            // Заполняем их ценами, датами и информацией о документе самой таблицы по коду аналога
            // используем временную таблицу
            "SELECT [Code], [Code_Analog1], [Price_Analog1], [Date_Analog1], [DocInfAn1] INTO [Svod1] FROM [Svod]; " +
            "UPDATE [Svod1] SET [Svod1].[Price_Analog1] = [Svod].[Price_Fin], [Svod1].[Date_Analog1] = [Svod].[Date_Fin], [Svod1].[DocInfAn1] = [Svod].[DocInf] " +
            "FROM [Svod] WHERE [Svod1].[Code_Analog1] = [Svod].[Code]; " +
            "UPDATE [Svod] SET [Svod].[Price_Analog1] = [Svod1].[Price_Analog1], [Svod].[Date_Analog1] = [Svod1].[Date_Analog1], [Svod].[DocInfAn1] = [Svod1].[DocInfAn1] " +
            "FROM [Svod1] WHERE [Svod1].[Code] = [Svod].[Code]; " +
            "DROP TABLE [Svod1]; " +

            // теперь то же самое со второй аналоговой ценой
            "SELECT [Code], [Code_Analog2], [Price_Analog2], [Date_Analog2], [DocInfAn2] INTO [Svod1] FROM [Svod]; " +
            "UPDATE [Svod1] SET [Svod1].[Price_Analog2] = [Svod].[Price_Fin], [Svod1].[Date_Analog2] = [Svod].[Date_Fin], [Svod1].[DocInfAn2] = [Svod].[DocInf] " +
            "FROM [Svod] WHERE [Svod1].[Code_Analog2] = [Svod].[Code]; " +
            "UPDATE [Svod] SET [Svod].[Price_Analog2] = [Svod1].[Price_Analog2], [Svod].[Date_Analog2] = [Svod1].[Date_Analog2], [Svod].[DocInfAn2] = [Svod1].[DocInfAn2] " +
            "FROM [Svod1] WHERE [Svod1].[Code] = [Svod].[Code]; " +
            "DROP TABLE [Svod1]; " +

            // Заполняем итоговые аналоговые данные

            // Сначала выбираем по максимальной дате
            "UPDATE [Svod] SET [Price_Analog] = [Price_Analog1], [Date_Analog] = [Date_Analog1] WHERE [Date_Analog1] > [Date_Analog2]; " +
            "UPDATE [Svod] SET [Price_Analog] = [Price_Analog2], [Date_Analog] = [Date_Analog2] WHERE [Date_Analog1] < [Date_Analog2]; " +
            // Если даты равны, то берем максимальную цену
            "UPDATE [Svod] SET [Price_Analog] = [Price_Analog1], [Date_Analog] = [Date_Analog1] WHERE [Date_Analog1] = [Date_Analog2] AND [Price_Analog1] >= [Price_Analog2]; " +
            "UPDATE [Svod] SET [Price_Analog] = [Price_Analog2], [Date_Analog] = [Date_Analog2] WHERE [Date_Analog1] = [Date_Analog2] AND [Price_Analog1] < [Price_Analog2]; " +
            // Если одно из полей пустое, то берем другое поле
            "UPDATE [Svod] SET [Price_Analog] = [Price_Analog1], [Date_Analog] = [Date_Analog1] WHERE [Price_Analog2] IS NULL; " +
            "UPDATE [Svod] SET [Price_Analog] = [Price_Analog2], [Date_Analog] = [Date_Analog2] WHERE [Price_Analog1] IS NULL; " +


            // Обнуляем старые подстановочные данные если Дата подстановки слишком старая
            "UPDATE [Svod] SET [Price_Analog] = NULL, [Date_Analog] = NULL WHERE YEAR(GETDATE()) -  YEAR([Date_Analog]) > 1 OR DATEDIFF(d, [Date_Analog], [Date_Fin]) > 30; " +

            // Добавляем новые поля Цена окончательная и Дата окончательная
            "ALTER TABLE [Svod] ADD [Price_Last] DECIMAL(28,2), [Date_Last] DATE; " +

            // Заполняем их
            // Если Дата реальная меньше чем на месяц Даты аналоговой, то берем ее, иначе берем аналоговые данные
            "UPDATE [Svod] SET [Price_Last] = [Price_Real], [Date_Last] = [Date_Real] WHERE DATEDIFF(d, [Date_Real], [Date_Analog]) <= 30; " +
            "UPDATE [Svod] SET [Price_Last] = [Price_Analog], [Date_Last] = [Date_Analog] WHERE DATEDIFF(d, [Date_Real], [Date_Analog]) > 30; " +
            // Если одно из полей пустое, то берем другое поле
            "UPDATE [Svod] SET [Price_Last] = [Price_Real], [Date_Last] = [Date_Real] WHERE [Price_Analog] IS NULL; " +
            "UPDATE [Svod] SET [Price_Last] = [Price_Analog], [Date_Last] = [Date_Analog] WHERE [Price_Real] IS NULL; " +

            // Добавляем поля Сравнение, Источник цены, Возраст цены, Информация о цене
            "ALTER TABLE [Svod] ADD [Compar] VARCHAR (2), [Price_Source] VARCHAR (15), [Price_Age] VARCHAR (5), [Price_Inf] VARCHAR (30); " +

            // Заполняем их
            // Если Дата окончательная больше, чем Дата первоначальная, то ставим + (то есть имеем новую цену)
            "UPDATE [Svod] SET [Compar] = '+' WHERE [Date_Fin] < [Date_Last]; " +
            // Если Дата окончательная меньше, чем Дата первоначальная, то ставим - (то есть имеем старую цену)
            "UPDATE [Svod] SET [Compar] = '-' WHERE [Date_Fin] >= [Date_Last]; " +
            // Если Дата окончательная = Дате первоначальной, то ставим тж (то есть имеем ту же цену) - перезаписывает предыдущие при равенстве цен
            "UPDATE [Svod] SET [Compar] = 'тж' WHERE [Price_Fin] = [Price_Last]; " +

            // Проставляем Источник цены (каждая следующая подстановка может перезаписывать предыдущие) по умолчанию ставим "подстановка"
            "UPDATE [Svod] SET [Price_Source] = 'подстановка'; " +
            // Если Цена первоначальная = Цене бухг., ставим "ведомость"
            "UPDATE [Svod] SET [Price_Source] = 'ведомость' WHERE Price_Fin = Price_Buh; " +
            // Если Цена первоначальная = Цене базы, ставим "база"
            "UPDATE [Svod] SET [Price_Source] = 'база' WHERE [Price_Fin] = [Price_Base]; " +

            // Проставляем Возраст цены, если Дата первоначальная старше Текущей даты на полгода, ставим "стар"
            "UPDATE [Svod] SET [Price_Age] = 'стар' WHERE DATEDIFF(d, [Date_Fin], CONVERT (DATE, GETDATE ())) > 180; " +

            // Проставляем Информацию о цене
            // Там, где поле Возраст цены пустое, берем только Источник цены
            "UPDATE [Svod] SET [Price_Inf] = [Price_Source] WHERE [Price_Age] IS NULL; " +
            // В противном случае к Источнику цены прибавляем (стар)
            "UPDATE [Svod] SET [Price_Inf] = [Price_Source] + ' (' + [Price_Age] + ')' WHERE [Price_Age] IS NOT NULL; " +

            // Возвращаемся к заполнению Информацию о док
            "UPDATE [Svod] SET [DocInfFin] = [DocInfAn2] WHERE [Price_Last] = [Price_Analog2] AND [DocInfAn2] IS NOT NULL; " +
            "UPDATE [Svod] SET [DocInfFin] = [DocInfAn1] WHERE [Price_Last] = [Price_Analog1] AND [DocInfAn1] IS NOT NULL; " +
            "UPDATE [Svod] SET [DocInfFin] = [DocInfBase] WHERE [Price_Last] = [Price_Base] AND [DocInfBase] IS NOT NULL; " +

            // Добавляем новые поля Дата и цена для копирования
            "ALTER TABLE [Svod] ADD [Price_Copy] DECIMAL(28,2), [Date_Copy] DATETIME, [DocInf_Copy] VARCHAR(255); " +

            // Заполняем их
            // по умолчанию ставим их равными первоначальными Дате и Цене
            "UPDATE [Svod] SET [Price_Copy] = [Price_Fin], [Date_Copy] = [Date_Fin], [DocInf_Copy] = [DocInf]; " +
            // там, где Цена последняя не пустая, а поле Обн не стоит е, ставим Цену и Дату последние
            "UPDATE [Svod] SET [Price_Copy] = [Price_Last], [Date_Copy] = [Date_Last] " +
            "WHERE [Price_Last] IS NOT NULL AND [Obn] IS NULL AND YEAR(GETDATE ()) - YEAR(Date_Last) < 2; " +

            // Если будем использовать старые данные
            "UPDATE [Svod] SET [DocInf_Copy] = [DocInfAn1] WHERE [DocInfAn1] IS NOT NULL AND [Price_Copy] = [Price_Analog1]; " +
            "UPDATE [Svod] SET [DocInf_Copy] = [DocInfFin] WHERE [DocInfFin] IS NOT NULL AND [Price_Copy] = [Price_Last]; " +

            // Информация о документах
            "ALTER TABLE [Svod] ADD [DT] DATE, [DTOtkl] INT, [DTComp] VARCHAR(255); " +
            "UPDATE [Svod] SET [DT] = CONVERT(DATE, SUBSTRING([DocInf_Copy], CHARINDEX (' от ', [DocInf_Copy]) + 4, 10)); " +
            "UPDATE [Svod] SET [DTOtkl] = DATEDIFF(d, [DT], [Date_Copy]) WHERE [DT] IS NOT NULL; " +
            "UPDATE [Svod] SET [DTComp] = '+' WHERE [Date_Base] = [DT] OR [Date_Buh] = [DT]";

            // Выполняем полученую команду
            execSQLCmd(strSQL);


            // Создаем представление сводной таблицы
            strSQL = "CREATE VIEW [SvodView] AS SELECT [Code], [Name], [EI], [Price_Fin], [Date_Fin], [Obn], [Code_Analog1], [Code_Analog2], [Note], " +
            "[Price_Last], [Date_Last], [Price_Copy], [Date_Copy], [Compar], [Price_Base], [Date_Base], [Price_Buh], [Date_Buh], [Price_Sub], [Date_Sub], " +
            "[Price_Real], [Date_Real], [Price_Analog1], [Date_Analog1], [Price_Analog2], [Date_Analog2], [Price_Analog], [Date_Analog] FROM [Svod]";
            // Выполняем полученую команду
            execSQLCmd(strSQL);
        }

        private void button7_Click(object sender, EventArgs e) // Кнопка Обнов.
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            try
            {
                Obnov();
                Obnov();
                Obnov();
                Obnov();
            }
            catch (Exception ex)
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Обновление базы данных выполнено!", "Сообщение");

        }
        private void Obnov() // Обновление сводной таблицы
        {
            // Вызываем формирование сводной таблицы
            Svod();

            // Копируем Цену и Дату для копирования в Цену и Дату первоначальную
            strSQL = "UPDATE [Svod] SET [Price_Fin] = [Price_Copy], [Date_Fin] = [Date_Copy], [DocInf] = [DocInf_Copy]; " +

            // Удаляем Лист 5 и заменяем его данными из сводной таблицы
            "DROP TABLE [Sheet5]; " +
            "SELECT [Code], [Name], [Name_Source], [EI], [Price_Fin], [Date_Fin], [Obn], [Code_Analog1], [Code_Analog2], [DocInf], [Note] INTO [Sheet5] FROM [Svod] ORDER BY [Code]";

            // Выполняем полученую команду
            execSQLCmd(strSQL);

            // Вызываем формирование сводной таблицы
            Svod();
        }

        private void button8_Click(object sender, EventArgs e) // Кнопка "Первый пересчет"
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            try
            {
                FirstObnov();
            }
            catch
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show("Таблица Sheet6 не существует!", "Сообщение");
                return;
            }

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Первый пересчет выполнен!", "Сообщение");
        }

        private void FirstObnov() // Первый пересчет обновления
        {
            try // Пытаемся удалить Лист 6 и его представление
            {
                execSQLCmd("DROP VIEW [Sheet6View]");              
            }
            catch
            {
            }
            try
            {
                execSQLCmd("DROP TABLE [Sheet6]");

            }
            catch
            {
            }


            // Формируем новый лист6
            strSQL = "SELECT [Prover].*, [Svod].[Price_Fin] AS [Price_Svod] , [Svod].[Date_Fin] AS [Date_Svod], [Svod].[Price_Inf] AS [Source] " +
                "INTO [Sheet6] FROM [Prover] LEFT JOIN [Svod] ON [Svod].[Code] = [Prover].[Code]; " +


            // Добавляем поля Цена подстан, Цена_последняя, Возраст цены подастановки, Код аналога
            "ALTER TABLE [Sheet6] ADD [Price_Sub] DECIMAL(28,2), [Price_Fin] DECIMAL(28,2), [Ved] VARCHAR (1), [Price_Sub_Age] VARCHAR (1), [Code_Analog] VARCHAR(255); " +

            // Заполняем поле Цена последняя
            // Где Цена сводная не пустая, берем ее, иначе берем первоначальную Цену
            "UPDATE [Sheet6] SET [Price_Fin] = [Price_Svod] WHERE Price_Svod IS NOT NULL; " +
            "UPDATE [Sheet6] SET [Price_Fin] = [Price] WHERE [Price_Fin] IS NULL; " +

            // Добавляем поля Сумма сводная и Сумма окончательная, Стоимость, Стоимость сводная, Стоимость окончательная
            "ALTER TABLE [Sheet6] ADD [Amount_Svod] DECIMAL(28,2), [Amount_Fin] DECIMAL(28,2), [Cost] DECIMAL(28,2), [Cost_Svod] DECIMAL(28,2), [Cost_Fin] DECIMAL(28,2); " +

            // Заполняем их

            // Сумму свода рассчитываем как Округление (Норма * Цена свода)
            "UPDATE [Sheet6] SET [Amount_Svod] = ROUND ([Norm] * [Price_Svod], 2), [Amount_Fin] = ROUND ([Norm] * [Price_Fin], 2); " +

            // Заполняем поля Стоимость, Стоимость сводная, Стоимость окончательная
            // Для этого создаем временную группировочную таблицу
            "SELECT [Product], SUM ( ROUND ([Amount], 2)) AS [Cost], SUM ([Amount_Svod]) AS [Cost_Svod], " +
            "SUM ([Amount_Fin]) AS [Cost_Fin] INTO [Sheet61] FROM [Sheet6] GROUP BY [Product]; " + 
            // Берем из нее нужные данные
            "UPDATE [Sheet6] SET [Sheet6].[Cost] = [Sheet61].[Cost], [Sheet6].[Cost_Svod] = [Sheet61].[Cost_Svod], [Sheet6].[Cost_Fin] = [Sheet61].[Cost_Fin] " +
            "FROM [Sheet61] WHERE [Sheet6].[Product] = [Sheet61].[Product]; " +
            // Удаляем временную таблицу
            "DROP TABLE [Sheet61]; " +

            // Добавляем поля Доля, Доля сводная и Доля окончательная
            "ALTER TABLE [Sheet6] ADD [Rate] DECIMAL(28,2), [Rate_Svod] DECIMAL(28,2), [Rate_Fin] DECIMAL(28,2); " +

            // Долю, Долю сводную и Долю окончательную рассчитываем как Округление (Соответствующая Сумма / Соответсвующая стоимость)
            "UPDATE [Sheet6] SET [Rate] = ROUND ([Amount] / [Cost] * 100, 2), " +
            "[Rate_Svod] = ROUND ([Amount_Svod] / [Cost_Svod] * 100, 2), [Rate_Fin] = ROUND ([Amount_Fin] / [Cost_Fin] * 100, 2); " +

            // Добавляем поля Фильтр, Сравнение 6/13, Сравнение 9/13, ЕИ_Текст
            "ALTER TABLE [Sheet6] ADD [Filter] VARCHAR (1), [Comp_6_13] VARCHAR (1), [Comp_9_13] VARCHAR (1), [EI_Text] VARCHAR (10); " +

            // Заполняем их
            // В Фильтре ставим +, когда Доля, Доля сводная или Доля окончательная >=1
            "UPDATE [Sheet6] SET [Filter] = '+' WHERE [Rate] >= 1 OR [Rate_Svod] >= 1 OR [Rate_Fin] >= 1; " +

            // В Сравнении 6/13 сравниваем Цену первоначальную и Цену окончательную. Когда они равны, ставим "-", иначе "+"
            "UPDATE [Sheet6] SET [Comp_6_13] = '+' WHERE [Price] <> [Price_Fin]; " +
            "UPDATE [Sheet6] SET [Comp_6_13] = '-' WHERE [Price] = [Price_Fin]; " +

            // В Сравнении 9/13 сравниваем Цену свода и Цену окончательную. Когда они равны, ставим "-", иначе "+"
            "UPDATE [Sheet6] SET [Comp_9_13] = '+' WHERE [Price_Svod] <> [Price_Fin]; " +
            "UPDATE [Sheet6] SET [Comp_9_13] = '-' WHERE [Price_Svod] = [Price_Fin]; " +

            // В поле ЕИ_Текст проставляем единицы измерения в текстовом виде, которые берем из таблицы OKEI
            "UPDATE [Sheet6] SET [Sheet6].[EI_Text] = [OKEI].[EI_Text] FROM [OKEI] WHERE [Sheet6].[EI] = [OKEI].[EI]";

            // Выполняем полученую команду
            execSQLCmd(strSQL);

            // Создаем представление для Листа 6
            strSQL = "CREATE VIEW [Sheet6View] AS SELECT [Code], [Name], [Dep], [EI_Text], [Norm], [Price], [Amount], [Product], [Price_Svod], [Date_Svod], " +
                "[Source], [Price_Sub], [Price_Fin], [Ved], [Price_Sub_Age], [Code_Analog], [Comp_6_13], [Comp_9_13], [Amount_Fin], [Cost_Fin], [Filter], [N] FROM [Sheet6]";

            // Выполняем полученую команду
            execSQLCmd(strSQL);
        }

        private void button9_Click(object sender, EventArgs e) // Кнопка Обн./сохр. цены
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            try
            {
                // Вызываем обновление Листа 6 при вводе подстановочных цен
                ObnovProv();
            }
            catch
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show("Таблица Sheet6 не существует!", "Сообщение");
                return;
            }

            try 
            {
                execSQLCmd("DROP TABLE [PromPrice]");
            }
            catch
            {

            }

            // Создаем временную группировочную таблицу
            strSQL = "SELECT [Code], MAX ([EI_Text]) AS [EI], MAX ([Price_Sub]) AS [Price], MAX ([Price_Sub_Age]) AS [Age], MAX([Code_Analog]) AS [Analog], MAX([Ved]) AS [Vedom] " +
            "INTO [PromPrice1] FROM [Sheet6] WHERE [Price_Sub] IS NOT NULL OR [Ved] IS NOT NULL GROUP BY [Code]; " +

            // Добавляем в нее поле Дата и заполняем
            "ALTER TABLE [PromPrice1] ADD [Date_Sub] DATE; " +
            // Где поле Age не заполнено, ставим текущую дату
            "UPDATE [PromPrice1] SET [Date_Sub] = GETDATE () WHERE [Age] IS NULL; " +
            // Иначе ставим 01.01 года, предшествующего текущему
            "UPDATE [PromPrice1] SET [Date_Sub] = CONVERT (DATE, '01.01.' + CONVERT(VARCHAR(255), YEAR(GETDATE ()) - 1)) WHERE [Age] IS NOT NULL; " +

            // Добавляем в нее поле Наименование и заполняем
            "ALTER TABLE [PromPrice1] ADD [Name] VARCHAR(255); " +
            "UPDATE [PromPrice1] SET [PromPrice1].[Name] = [Sheet6].[Name] FROM [Sheet6] WHERE [PromPrice1].[Code] = [Sheet6].[Code]; " +

            // Добавляем в окончательную таблицу все поля из временной таблицы в другом порядке и с другими названиями
            "SELECT [Code], [Name], [EI], [Price], [Date_Sub] AS [Date], [Age], [Analog], [Vedom] AS [Ved] INTO [PromPrice] FROM [PromPrice1]; " +

            // Удаляем временную таблицу
            "DROP TABLE [PromPrice1]; " +

            // Подставляем на Лист 6 данные из полученной таблицы
            "UPDATE [Sheet6] SET [Sheet6].[Price_Sub] = [PromPrice].[Price], [Sheet6].[Price_Sub_Age] = [PromPrice].[Age], [Sheet6].[Code_Analog] = [PromPrice].[Analog] " +
            "FROM [PromPrice] WHERE [Sheet6].[Code] = [PromPrice].[Code]";

            // Выполняем полученую команду
            execSQLCmd(strSQL);

            // Вызываем обновление Листа 6 при вводе подстановочных цен
            ObnovProv();

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Обновление/сохранение цен выполнено!", "Сообщение");
        }

        private void ObnovProv() // Обновляем Лист 6 при вводе подстановочных цен
        {
            // Пересчитываем поле Цена последняя
            // Если все цены пустые, ставим пустоту
            strSQL = "UPDATE [Sheet6] SET [Price_Fin] = NULL WHERE [Price] IS NULL AND [Price_Svod] IS NULL AND [Price_Sub] IS NULL; " +
            // Если Цена первоначальная не пустая, ставим ее
            "UPDATE [Sheet6] SET [Price_Fin] = [Price] WHERE [Price] IS NOT NULL; " +
            // Если Цена сводная не пустая, ставим ее
            "UPDATE [Sheet6] SET [Price_Fin] = [Price_Svod] WHERE [Price_Svod] IS NOT NULL; " +
            // Если Цена подстановочная не пустая, ставим ее
            "UPDATE [Sheet6] SET [Price_Fin] = [Price_Sub] WHERE [Price_Sub] IS NOT NULL; " +
            // Пересчитываем поле Сумма последняя
            "UPDATE [Sheet6] SET [Amount_Fin] = ROUND ([Norm] * [Price_Fin], 2); " +

            // Пересчитываем поле Стоимость последняя
            // Для этого создаем временную группировочную таблицу
            "SELECT [Product], SUM ([Amount_Fin]) AS [Cost_Fin] INTO [Sheet61] FROM [Sheet6] GROUP BY [Product]; " +
            // Берем из нее данные о Стоимости
            "UPDATE [Sheet6] SET [Sheet6].[Cost_Fin] = [Sheet61].[Cost_Fin] FROM [Sheet61] WHERE [Sheet6].[Product] = [Sheet61].[Product]; " +
            // Удаляем временную таблицу
            "DROP TABLE [Sheet61]; " +

            // Пересчитываем поле Доля последняя
            "UPDATE [Sheet6] SET [Rate_Fin] = ROUND ([Amount_Fin] / [Cost_Fin] * 100, 2); " +

            // Пересчитываем поля Фильтр, Сравнение 6/13, Сравнение 9/13

            // Сначала обнуляем старые данные
            "UPDATE [Sheet6] SET [Filter] = NULL, [Comp_6_13] = NULL, [Comp_9_13] = NULL; " +

            // В Фильтре ставим +, когда Доля, Доля сводная или Доля окончательная >=1
            "UPDATE [Sheet6] SET [Filter] = '+' WHERE [Rate] >= 1 OR [Rate_Svod] >= 1 OR [Rate_Fin] >= 1; " +

            // В Сравнении 6/13 сравниваем Цену первоначальную и Цену окончательную
            // Когда они равны, ставим "-", иначе "+"
            "UPDATE [Sheet6] SET [Comp_6_13] = '+' WHERE [Price] <> [Price_Fin]; " +
            "UPDATE [Sheet6] SET [Comp_6_13] = '-' WHERE [Price] = [Price_Fin]; " +

            // В Сравнении 9/13 сравниваем Цену свода и Цену окончательную
            // Когда они равны, ставим "-", иначе "+"
            "UPDATE [Sheet6] SET [Comp_9_13] = '+' WHERE [Price_Svod] <> [Price_Fin]; " +
            "UPDATE [Sheet6] SET [Comp_9_13] = '-' WHERE [Price_Svod] = [Price_Fin]";

            // Выполняем полученую команду
            execSQLCmd(strSQL);
        }

        private void button10_Click(object sender, EventArgs e) // Кнопка Повт. пересчет
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();


            try // Выполняем действия для проверки существования
            {
                strSQL = "SELECT TOP 1 [Code] INTO [Vr] FROM [PromPrice]; " +
                    "DROP TABLE Vr";
                execSQLCmd(strSQL);
            }
            catch
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show("Таблица PromPrice не существует!", "Сообщение");
                return;
            }

            try
            {
                // Вызываем первый пересчет обновления
                FirstObnov();
            }
            catch
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show("Таблица Sheet6 не существует!", "Сообщение");
                return;
            }


            // Берем сохраненные данные о проверке из таблицы PromPrice
            strSQL = "UPDATE [Sheet6] SET [Sheet6].[Price_Sub] = [PromPrice].[Price], [Sheet6].[Price_Sub_Age] = [PromPrice].[Age], " +
                "[Sheet6].[Code_Analog] = [PromPrice].[Analog], [Sheet6].[Ved] = [PromPrice].[Ved] FROM [PromPrice] WHERE [Sheet6].[Code] = [PromPrice].[Code]";

            // Выполняем полученую команду
            execSQLCmd(strSQL);

            // Вызываем обновление Листа 6 при вводе подстановочных цен
            ObnovProv();

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Повторный пересчет цен выполнен!", "Сообщение");
        }

        private void button11_Click(object sender, EventArgs e) // Кнопка Законч. пров
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            //Выбираем, закончить ли проверку
            DialogResult vybor = MessageBox.Show("Закончить проверку?", "Диалог", MessageBoxButtons.YesNo);

            switch (vybor)
            {
                case DialogResult.Yes:
                    {
                        try
                        {
                            execSQLCmd("DROP TABLE [PromCode]");
                        }
                        catch
                        {
                        }

                        try
                        {
                            execSQLCmd("DROP TABLE [PromPrice]");
                        }
                        catch
                        {
                        }

                        try
                        {
                            // Вызываем первый пересчет обновления
                            FirstObnov();
                        }
                        catch
                        {
                            txn.Rollback();
                            cnSQL.Close();
                            MessageBox.Show("Таблица Sheet6 не существует!", "Сообщение");
                            return;
                        }
                    };
                    break;
                case DialogResult.No:
                    {
                        txn.Rollback();
                        cnSQL.Close();
                        return;
                    };
                    break;
            }

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Проверка цен закончена!", "Сообщение");
        }
        private void button12_Click(object sender, EventArgs e) // Кнопка Cформ. ведом.
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            strSQL = "SELECT [N] FROM [Sheet6] GROUP BY N ORDER BY N";
            cmdSQL = new SqlCommand(strSQL, cnSQL, txn);
            rdrSQL = cmdSQL.ExecuteReader();
            lstN1 = new List<string>();
            lstN2 = new List<string>();
            while (rdrSQL.Read())
            {
                lstN1.Add(rdrSQL[0].ToString());
                lstN2.Add(rdrSQL[0].ToString());
            }
            rdrSQL.Close();

            try
            {
                execSQLCmd("DROP TABLE [OMTS]");
            }
            catch { }

            // Выбираем отмеченные поля во временную таблицу
            strSQL = "SELECT [Code] AS [Код], [Name] AS [Наименование], [EI] AS [ЕИ], [Price] AS [Цена без НДС] INTO [OMTS1] FROM [PromPrice] WHERE [Ved] IS NOT NULL; " +

            // Обнуляем цены
            "UPDATE [OMTS1] SET [Цена без НДС] = NULL; " +

            // Добавляем поле Изделие и заполняем
            "ALTER TABLE [OMTS1] ADD [Изделие] VARCHAR(255); " +
            "UPDATE [OMTS1] SET [OMTS1].[Изделие] = [Sheet6].[Product] FROM [Sheet6] WHERE [OMTS1].[Код] = [Sheet6].[Code]; ";

            DialogResult res = MessageBox.Show("Выбрать определенные позиции?", "Диалог", MessageBoxButtons.YesNo);
            switch (res)
            {
                case DialogResult.Yes: 
                    {
                        strSQL = strSQL + "SELECT DISTINCT [Код], [Наименование], [ЕИ], [Цена без НДС], [Изделие] INTO [OMTS] FROM ([OMTS1] LEFT JOIN [Sheet6] " +
                        "ON [OMTS1].[Код] = [Sheet6].[Code]) ";
                        Form7 frm7 = new Form7();
                        frm7.lst1 = this.lstN1;
                        frm7.lst2 = this.lstN2;
                        frm7.pf = this;
                        frm7.ShowDialog(this);
                        strSQL = strSQL + dobavka + "; DROP TABLE [OMTS1]";
                    };
                    break;

                case DialogResult.No:
                    {
                        strSQL = strSQL + "SELECT * INTO [OMTS] FROM [OMTS1]; DROP TABLE [OMTS1]; ";
                    };
                    break;
            }

            try
            {
                execSQLCmd(strSQL);
            }
            catch (Exception ex)
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Таблица для проверки цен сформирована!", "Сообщение");

        }

        private void button13_Click(object sender, EventArgs e) // Кнопка Экспорт в Excel
        {
            // Создаем объект Excel
            Excel.Application exc = new Excel.Application();            
            // Добавляем новую книгу
            Excel.Workbook wb = exc.Workbooks.Add();
            // Добавляем новую книгу
            Excel.Worksheet ws = wb.Worksheets[1];
            // Устанавливаем для нее альбомную ориентацию
            ws.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Создаем рекордсет, куда помещаем целевую таблицу
            strSQL = "SELECT * FROM OMTS";
            cmdSQL = new SqlCommand(strSQL, cnSQL);
            int i = 2; // Номер строки
            int m = 0; // Количество столбцов
            ws.Cells[1, 2].Value = "69-62 Дмитрий";
            ws.Cells[1, 2].Font.Name = "Arial";
            ws.Cells[1, 2].Font.Size = 12;

            rdrSQL = cmdSQL.ExecuteReader();
            while (rdrSQL.Read())
            {
                for (int k = 0; k < rdrSQL.FieldCount; k++)
                {
                    if (k == 0)
                    {
                        i++;
                    }

                    if (i == 3)
                    {
                        ws.Cells[i - 1, k + 1].Value = rdrSQL.GetName(k);
                        ws.Cells[i - 1, k + 1].Borders.LineStyle = true;
                        ws.Cells[i - 1, k + 1].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                        ws.Cells[i - 1, k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        ws.Cells[i - 1, k + 1].Font.Name = "Times New Roman";
                        ws.Cells[i - 1, k + 1].Font.Size = 12;

                        ws.Cells[i, k + 1].Value = rdrSQL[k];
                        ws.Cells[i, k + 1].Borders.LineStyle = true;
                        ws.Cells[i, k + 1].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                        ws.Cells[i, k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        ws.Cells[i, k + 1].Font.Name = "Times New Roman";
                        ws.Cells[i, k + 1].Font.Size = 12;
                    }
                    else
                    {
                        ws.Cells[i, k + 1].Value = rdrSQL[k];
                        ws.Cells[i, k + 1].Borders.LineStyle = true;
                        ws.Cells[i, k + 1].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                        ws.Cells[i, k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        ws.Cells[i, k + 1].Font.Name = "Times New Roman";
                        ws.Cells[i, k + 1].Font.Size = 12;
                    }
                }
            }
            m = rdrSQL.FieldCount;
            rdrSQL.Close();

            ws.Range[ws.Cells[2, 1], ws.Cells[i + 1, m]].Columns.Autofit(); // Автоподбор ширины столбцов
                                                                                // Устанавливаем выравнивание в ячейках
            ws.Range[ws.Cells[3, 1], ws.Cells[i + 2, 1]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ws.Range[ws.Cells[3, 2], ws.Cells[i + 2, 2]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ws.Range[ws.Cells[3, 3], ws.Cells[i + 2, 3]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ws.Range[ws.Cells[3, 5], ws.Cells[i + 2, 5]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            // Устанавливаем ширину последнего столбца: если его длина > 31, то ставим 31 и делаем автоподбор ширины текста
            if (ws.Range[ws.Cells[3, 5], ws.Cells[i + 2, 5]].Columns.ColumnWidth > 31)
            {
                ws.Range[ws.Cells[3, 5], ws.Cells[i + 2, 5]].Columns.ColumnWidth = 31;
                ws.Range[ws.Cells[3, 5], ws.Cells[i + 2, 5]].ShrinkToFit = true;
            }
                
            // Устанавливаем ширину столбца с наименованиями
            int Dlin = 0; // Длина всей строки
            int Ost = 0; // Длина без второй строки
            for (int l = 1; l < m; l++)
            {
                Dlin = Dlin + Convert.ToInt32(ws.Range[ws.Cells[3, l], ws.Cells[i + 2, l]].Columns.ColumnWidth);

                if (l != 2)
                {
                     Ost = Ost + Convert.ToInt32(ws.Range[ws.Cells[3, l], ws.Cells[i + 2, l]].Columns.ColumnWidth);
                }

            }

            if (Dlin > 96)
            {
                ws.Range[ws.Cells[2, 2], ws.Cells[i + 2, 2]].Columns.ColumnWidth = 96 - Ost;
            }

            ws.Range[ws.Cells[2, 2], ws.Cells[i + 2, 2]].WrapText = true;
            ws.Range[ws.Cells[2, 2], ws.Cells[i + 2, 2]].Rows.Autofit();

            exc.Visible = true;
            // Сохраняем в файл
            // objExcel.ActiveWorkbook.SaveAs FileName:= CurrentProject.Path + "\Èñòî÷íèêè\" + "Ïðîâåðêà.xls", FileFormat:=-4143  'Ñîîòâåòñòâóåò xlWorkbookNormal (Êíèãà Excel 97-03)

            // Выходим из приложения и обнуляем объект
            exc.Quit();
        }
        

        private void button14_Click(object sender, EventArgs e) // Кнопка Зап. изм. без кода
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            // Добавляем новые позиции из PromPrice на Лист 4
            strSQL = "INSERT INTO [Sheet4] ([Code], [Name_Sub], [EI_Sub]) SELECT [PromPrice].[Code], [PromPrice].[Name], [PromPrice].[EI] " +
                    "FROM [PromPrice] LEFT JOIN [Sheet4] ON [PromPrice].[Code] = [Sheet4].[Code] " +
                    "WHERE [Sheet4].[Code] IS NULL AND [PromPrice].[Analog] IS NULL AND [PromPrice].[Price] IS NOT NULL; " +

            // Добавляем новые Цену и Время из PromPrice на Лист 4
            "UPDATE [Sheet4] SET [Sheet4].[Price_Sub] = [PromPrice].[Price], [Sheet4].[Date_Sub] = [PromPrice].[Date] " +
            "FROM [PromPrice] WHERE [PromPrice].[Code] = [Sheet4].[Code] AND [PromPrice].[Analog] IS NULL AND [PromPrice].[Price] IS NOT NULL";

            try 
            {
                // Выполняем полученую команду
                execSQLCmd(strSQL);

                // Вызываем обновление сводной таблицы
                Obnov();
            }
            catch (Exception ex)
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Изменения без кода записаны!", "Сообщение");
        }

        private void button15_Click(object sender, EventArgs e)  // Кнопка Проверка кодов
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            try // Пытаемся удалить таблицу PromCode
            {
                execSQLCmd("DROP TABLE [PromCode]");
            }
            catch
            {
            }

            // Берем все поля с Листа 5
            strSQL = "SELECT * INTO [PromCode] FROM [Sheet5]; " + 

            // Добавляем новые данные из PromPrice
            "INSERT INTO [PromCode] ([Code], [Name], [EI]) SELECT [PromPrice].[Code], [PromPrice].[Name], [PromPrice].[EI] " +
            "FROM [PromPrice] LEFT JOIN [PromCode] ON [PromPrice].[Code] = [PromCode].[Code] WHERE [PromCode].[Code] IS NULL AND [PromPrice].[Analog] IS NOT NULL; " +

            // В Источник наименования ставим "подстановка"
            "UPDATE [PromCode] SET [Name_Source] = 'подстановка' WHERE [Name_Source] IS NULL; " +

            // Создаем поля ставнения и заполняем
            // Где Код аналога из PromPrice = Кодам аналога из данной таблицы, ставим "-", иначе "+"

            "ALTER TABLE [PromCode] ADD [Analog] VARCHAR (255), [Comp1] VARCHAR (1), [Comp2] VARCHAR (1); " +
            "UPDATE [PromCode] SET [PromCode].[Analog] = [PromPrice].[Analog] FROM [PromPrice] WHERE [PromCode].[Code] = [PromPrice].[Code]; " +
            "UPDATE [PromCode] SET [Comp1] = '+' WHERE [Code_Analog1] <> [Analog]; " +
            "UPDATE [PromCode] SET [Comp1] = '+' WHERE [Code_Analog1] IS NULL AND [Analog] IS NOT NULL; " +
            "UPDATE [PromCode] SET [Comp1] = '-' WHERE [Code_Analog1] = [Analog]; " +
            "UPDATE [PromCode] SET [Comp2] = '+' WHERE [Code_Analog2] <> [Analog]; " +
            "UPDATE [PromCode] SET [Comp2] = '-' WHERE [Code_Analog2] = [Analog]";

            try
            {
                // Выполняем полученую команду
                execSQLCmd(strSQL);
            }
            catch (Exception ex)
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Таблица PromCode сформирована!", "Сообщение");
        }

        private void button16_Click(object sender, EventArgs e) // Кнопка Пересч. пров. кодов
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            try
            {
                RefreshCode();
            }
            catch
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show("Таблица PromCode не существует!", "Сообщение");
                return;
            }

            // Фиксируем транзакцию
            txn.Commit();

            cnSQL.Close();
            MessageBox.Show("Таблица PromCode пересчитана!", "Сообщение");
        }

        private void RefreshCode() // Пересчет таблицы для проверки кодов
        {
            // Пересчитаываем поля сравнения

            // Для начала обнуляем поля сравнения
            strSQL = "UPDATE [PromCode] SET [Comp1] = NULL; " +
                    "UPDATE [PromCode] SET [Comp2] = NULL; " +

            // Где Код аналога из PromPrice = Кодам аналога из данной таблицы, ставим "-", иначе "+
            "UPDATE [PromCode] SET [Comp1] = '+' WHERE [Code_Analog1] <> [Analog]; " +
            "UPDATE [PromCode] SET [Comp1] = '+' WHERE [Code_Analog1] IS NULL AND [Analog] IS NOT NULL; " +
            "UPDATE [PromCode] SET [Comp1] = '-' WHERE [Code_Analog1] = [Analog]; " +
            "UPDATE [PromCode] SET [Comp2] = '+' WHERE [Code_Analog2] <> [Analog]; " +
            "UPDATE [PromCode] SET [Comp2] = '-' WHERE [Code_Analog2] = [Analog]";

            // Выполняем полученую команду
            execSQLCmd(strSQL);
        }

        private void button17_Click(object sender, EventArgs e) // Кнопка Законч. пров. кодов
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            try
            {
                execSQLCmd("DROP TABLE [Sheet5]");
            }
            catch
            {
            }

            // Добавляем полученные данные из PromCode на Лист 5
            strSQL = "SELECT [Code], [Name], [Name_Source], [EI], [Price_Fin], [Date_Fin], [Obn], [Code_Analog1], [Code_Analog2], [DocInf], [Note] INTO [Sheet5] FROM [PromCode]";
            // Выполняем полученую команду

            try
            {
                execSQLCmd(strSQL);
            }
            catch
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show("Таблица PromCode не существует!", "Сообщение");
                return;
            }            

            // Вызываем обновление сводной таблицы
            Obnov();
            Obnov();

            // Удаляем таблицу PromCode, которая больше не нужна
            execSQLCmd("DROP TABLE [PromCode]");

            // Фиксируем транзакцию
            txn.Commit();

            cnSQL.Close();
            MessageBox.Show("Проверка кодов закончена!", "Сообщение");
        }

        private void button18_Click(object sender, EventArgs e) // Кнопка Пров. наим.
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            try // Пытаемся удалить таблицу PromCode
            {
                execSQLCmd("DROP TABLE [PromName]");
            }
            catch
            {
            }
            // Выбираем Код, Наименование, Источник наименования из Листа 5
            strSQL = "SELECT [Code], [Name], [Name_Source] INTO [PromName] FROM [Sheet5]; " +

            // Добавляем поля Наименование ИВЦ, Сравнение 1, Наименование бух., Сравнение 2,
            // Наименование подстановки, Сравнение 3, Наименование проверки, Сравнение 4,
            // Наименование для копирования, Сравнение 5
            "ALTER TABLE [PromName] ADD [Name_IVC] VARCHAR(255), [Comp1] VARCHAR(1), [Name_Buh] VARCHAR(255), [Comp2] VARCHAR(1), " +
            "[Name_Sub] VARCHAR(255), [Comp3] VARCHAR(1), [Name_Ver] VARCHAR(255), [Comp4] VARCHAR(1), " +
            "[Name_Copy] VARCHAR(255), [Name_Copy_Source] VARCHAR(255), [Comp5] VARCHAR(1); " +

            // Заполняем поле Наименование ИВЦ данными из Листа2
            "UPDATE [PromName] SET [PromName].[Name_IVC] = [Sheet2].[Name_IVC] FROM [Sheet2] WHERE [PromName].[Code] = [Sheet2].[Code]; " +

            // Заполняем поле Сравнение 1
            // Если Наименование первоначальное = Наименованию ИВЦ, то ставим +, иначе -
            "UPDATE [PromName] SET [Comp1] = '-' WHERE [Name] = [Name_IVC]; " +
            "UPDATE [PromName] SET [Comp1] = '+' WHERE [Name] <> [Name_IVC]; " +

            // Заполняем поле Наименование бух. данными из Листа3
            "UPDATE [PromName] SET [PromName].[Name_Buh] = [Sheet3].[Name_Buh] FROM [Sheet3] WHERE [PromName].[Code] = [Sheet3].[Code]; " +

            // Заполняем поле Сравнение 2
            // Если Наименование первоначальное = Наименованию бух., то ставим +, иначе -
            "UPDATE [PromName] SET [Comp2] = '-' WHERE [Name] = [Name_Buh]; " +
            "UPDATE [PromName] SET [Comp2] = '+' WHERE [Name] <> [Name_Buh]; " +

            // Заполняем поле Наименование подстановки. данными из Листа4
            "UPDATE [PromName] SET [PromName].[Name_Sub] = [Sheet4].[Name_Sub] FROM [Sheet4] WHERE [PromName].[Code] = [Sheet4].[Code]; " +

            // Заполняем поле Сравнение 3
            // Если Наименование первоначальное = Наименованию подст., то ставим +, иначе -
            "UPDATE [PromName] SET [Comp3] = '-' WHERE [Name] = [Name_Sub]; " +
            "UPDATE [PromName] SET [Comp3] = '+' WHERE [Name] <> [Name_Sub]";

            try
            {
                // Выполняем полученую команду
                execSQLCmd(strSQL);
            }
            catch (Exception ex)
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            try
            {
                strSQL = "UPDATE [PromName] SET [PromName].[Name_Ver] = [Sheet6].[Name] FROM [Sheet6] WHERE [PromName].[Code] = [Sheet6].[Code]; " +
                // Заполняем поле Сравнение 4
                // Если Наименование первоначальное = Наименованию пров., то ставим +, иначе -
                "UPDATE [PromName] SET [Comp4] = '-' WHERE [Name] = [Name_Ver]; " +
                "UPDATE [PromName] SET [Comp4] = '+' WHERE [Name] <> [Name_Ver]";
                // Выполняем полученую команду
                execSQLCmd(strSQL);
            }
            catch
            {
            }

            // Заполняем поле Наименование для копирования

            // По умолчанию оставляем то же самое Наименование и Источник цены
            strSQL = "UPDATE [PromName] SET [Name_Copy] = [Name], Name_Copy_Source = Name_Source; " +

            // Далее последовательно выбираем Наименование бух., ИВЦ, подстан., проверки,
            // которое будет перезаписывать предыдущие данные

            "UPDATE [PromName] SET [Name_Copy] = [Name_Buh], [Name_Copy_Source] = 'ведомость' WHERE [Name_Buh] IS NOT NULL AND [Name_Source] = 'ведомость'; " +
            "UPDATE [PromName] SET [Name_Copy] = [Name_IVC], [Name_Copy_Source] = 'база' WHERE [Name_IVC] IS NOT NULL AND [Name_Source] <> 'подстановка'; " +
            "UPDATE [PromName] SET [Name_Copy] = [Name_Sub], [Name_Copy_Source] = 'подстановка' WHERE [Name_Sub] IS NOT NULL; " +
            "UPDATE [PromName] SET [Name_Copy] = [Name_Ver], [Name_Copy_Source] = 'подстановка' WHERE [Name_Ver] IS NOT NULL; " +

            // Заполняем поле Сравнение 5
            // Если Наименование первоначальное = Наименованию для копир., то ставим +, иначе -
            "UPDATE [PromName] SET [Comp5] = '-' WHERE [Name] = [Name_Copy]; " +
            "UPDATE [PromName] SET [Comp5] = '+' WHERE [Name] <> [Name_Copy]";
            // Выполняем полученую команду
            execSQLCmd(strSQL);

            // Фиксируем транзакцию
            txn.Commit();

            cnSQL.Close();
            MessageBox.Show("Таблица Проверки наименований сформирована!", "Сообщение");
        }

        private void button19_Click(object sender, EventArgs e) // Кнопка Пересч. пров. наим.
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            try
            {
                ProvName();
            }
            catch
            {
                txn.Rollback();
                cnSQL.Close();
                MessageBox.Show("Таблица PromName не существует!", "Сообщение");
                return;
            }

            // Фиксируем транзакцию
            txn.Commit();

            cnSQL.Close();
            MessageBox.Show("Пересчет Проверки наименований закончен!", "Сообщение");            
        }

        private void ProvName() // Пересчет проверки наименований
        {
            // Обнуляем текущие значения
            strSQL = "UPDATE [PromName] SET [Comp1] = NULL, [Comp2] = NULL, [Comp3] = NULL, [Comp4] = NULL, [Comp5] = NULL; " +

            // Заполняем поле Сравнение 1
            // Если Наименование первоначальное = Наименованию ИВЦ, то ставим +, иначе -
            "UPDATE [PromName] SET [Comp1] = '-' WHERE [Name] = [Name_IVC]; " +
            "UPDATE [PromName] SET [Comp1] = '+' WHERE [Name] <> [Name_IVC]; " +

            // Заполняем поле Сравнение 2
            // Если Наименование первоначальное = Наименованию бух., то ставим +, иначе -
            "UPDATE [PromName] SET [Comp2] = '-' WHERE [Name] = [Name_Buh]; " +
            "UPDATE [PromName] SET [Comp2] = '+' WHERE [Name] <> [Name_Buh]; " +

            // Заполняем поле Сравнение 3
            // Если Наименование первоначальное = Наименованию подст., то ставим +, иначе -
            "UPDATE [PromName] SET [Comp3] = '-' WHERE [Name] = [Name_Sub]; " +
            "UPDATE [PromName] SET [Comp3] = '+' WHERE [Name] <> [Name_Sub]; " +

            // Заполняем поле Сравнение 4
            // Если Наименование первоначальное = Наименованию пров., то ставим +, иначе -
            "UPDATE [PromName] SET [Comp4] = '-' WHERE [Name] = [Name_Ver]; " +
            "UPDATE [PromName] SET [Comp4] = '+' WHERE [Name] <> [Name_Ver]; " +

            // Заполняем поле Наименование для копирования

            // По умолчанию оставляем то же самое Наименование и Источник цены
            "UPDATE [PromName] SET [Name_Copy] = [Name], [Name_Copy_Source] = [Name_Source]; " +

            // Далее последовательно выбираем Наименование бух., ИВЦ, подстан., проверки,
            // которое будет перезаписывать предыдущие данные

            "UPDATE [PromName] SET [Name_Copy] = [Name_Buh], [Name_Copy_Source] = 'ведомость' WHERE [Name_Buh] IS NOT NULL AND [Name_Source] = 'ведомость'; " +
            "UPDATE [PromName] SET [Name_Copy] = [Name_IVC], [Name_Copy_Source] = 'база' WHERE [Name_IVC] IS NOT NULL AND [Name_Source] <> 'подстановка'; " +
            "UPDATE [PromName] SET [Name_Copy] = [Name_Sub], [Name_Copy_Source] = 'подстановка' WHERE [Name_Sub] IS NOT NULL; " +
            "UPDATE [PromName] SET [Name_Copy] = [Name_Ver], [Name_Copy_Source] = 'подстановка' WHERE [Name_Ver] IS NOT NULL; " +

            // Заполняем поле Сравнение 5
            // Если Наименование первоначальное = Наименованию для копир., то ставим +, иначе -
            "UPDATE [PromName] SET [Comp5] = '-' WHERE [Name] = [Name_Copy]; " +
            "UPDATE [PromName] SET [Comp5] = '+' WHERE [Name] <> [Name_Copy]";

            // Выполняем полученую команду
            execSQLCmd(strSQL);
        }
        private void button20_Click(object sender, EventArgs e) // Кнопка Скопир. все наим.
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            //Выбираем, скопировать ли все наименования
            DialogResult vybor = MessageBox.Show("Скопировать все наименования?", "Диалог", MessageBoxButtons.YesNo);

            switch (vybor)
            {
                case DialogResult.Yes:
                    {
                        try
                        {
                            // Вызываем копирование всех позиций
                            CopyAll();
                        }
                        catch
                        {
                            txn.Rollback();
                            cnSQL.Close();
                            MessageBox.Show("Таблица PromName не существует!", "Сообщение");
                            return;
                        }
                    };
                    break;
                case DialogResult.No:
                    {
                        txn.Rollback();
                        cnSQL.Close();
                        return;
                    };
                    break;
            }

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Все наименования скопированы!", "Сообщение");
        }

        private void CopyAll() // Скопировать все позиции в проверке наименований
        {
            // Копируем Наименование для копирования и Источник наименований для копирования в соответствующие первоначальные позиции
            strSQL = "UPDATE [PromName] SET [Name] = [Name_Copy], [Name_Source] = [Name_Copy_Source]";
            // Выполняем полученую команду
            execSQLCmd(strSQL);
            // Вызываем пересчет проверки наименований
            ProvName();
        }

        private void button21_Click(object sender, EventArgs e) // Законч. пров. наим.
        {
            // Открываем соедиение
            try
            {
                cnSQLOpen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Сообщение");
                return;
            }

            // Начинаем транзакцию
            txn = cnSQL.BeginTransaction();

            //Выбираем, скопировать ли все наименования
            DialogResult vybor = MessageBox.Show("Закончить проверку наименований?", "Диалог", MessageBoxButtons.YesNo);

            switch (vybor)
            {
                case DialogResult.Yes:
                    {
                        try
                        {
                            // Вызываем завершение проверки наименований
                            FinishProvName();
                        }
                        catch (Exception ex)
                        {                            
                            txn.Rollback();
                            cnSQL.Close();
                            MessageBox.Show(ex.Message, "Сообщение");
                            return;
                        }
                    };
                    break;
                case DialogResult.No:
                    {
                        txn.Rollback();
                        cnSQL.Close();
                        return;
                    };
                    break;
            }

            // Фиксируем транзакцию
            txn.Commit();

            // Закрываем соедиение
            cnSQL.Close();
            MessageBox.Show("Все наименования скопированы!", "Сообщение");
        }


        private void FinishProvName() // Закончить проверку наименований
        {
            // Копируем Наименование и Источник наименований из проверки на лист 5
            strSQL = "UPDATE [Sheet5] SET [Sheet5].[Name] = [PromName].[Name], [Sheet5].[Name_Source] = [PromName].[Name_Source] FROM [PromName] WHERE [Sheet5].[Code] = [PromName].[Code]";
            // Выполняем полученую команду
            execSQLCmd(strSQL);

            // Формирование сводной таблицы
            Svod();

            // Удаляем таблицу проверки
            strSQL = "DROP TABLE [PromName]";
            // Выполняем полученую команду
            execSQLCmd(strSQL);
        }
    }
}

