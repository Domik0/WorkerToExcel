using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Core;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace WorkerToExel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const int WIN_1252_CP = 1252; // Windows ANSI codepage 1252
        private List<Worker> workers = new List<Worker>();
        public Worker selectWorker = new Worker();
        Excel.Application excelApp;
        Excel._Worksheet workSheet;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = selectWorker;
        }

        /// <summary>
        /// Save data in Excel(File format CSV(UTF-8))
        /// </summary>
        private void Save(object sender, RoutedEventArgs e)
        {
            if(excelApp != null)
                CloseExcel();
            CreateExcel();
            AddData();

            string path = GetPath();
            if (path != null)
            {
                //Я честно не знаю почему не работает, но вот эта строчка должна устанавливать дефолтную кодировку utf-8.
                excelApp.DefaultWebOptions.Encoding = MsoEncoding.msoEncodingUTF8;
                try
                {
                    workSheet.SaveAs(path, Excel.XlFileFormat.xlCSV);
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message,
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            CloseExcel();

            //string path = "C:\\Users\\dafed\\Desktop\\123.csv";
            string data;
            if(!string.IsNullOrEmpty(path))
            {
                using (StreamReader sr = new StreamReader(path, Encoding.UTF8))
                {
                    data = sr.ReadToEnd();
                }

                using (StreamWriter sw = new StreamWriter(path, true, Encoding.UTF8))
                {
                    sw.Write("");
                }

                using (StreamReader sr = new StreamReader(path, Encoding.UTF8))
                {
                    data = sr.ReadToEnd();
                }
            }
            //File.Replace(path, new StreamWriter(path, false, Encoding.GetEncoding(WIN_1252_CP)));
        }

        /// <summary>
        /// Close excel file.
        /// </summary>
        private void CloseExcel()
        {
            excelApp.Workbooks.Close();
            excelApp.Quit();
            workSheet = null;
            excelApp = null;
        }

        /// <summary>
        /// Create Excel file and add header
        /// </summary>
        private void CreateExcel()
        {
            excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            CreateHeader();
        }

        /// <summary>
        /// Create header naming for database
        /// </summary>
        void CreateHeader()
        {
            int numberColumnHeader = 1;
            SetDataOnCell(numberColumnHeader, "A", "id");
            SetDataOnCell(numberColumnHeader, "B", "email");
            SetDataOnCell(numberColumnHeader, "C", "lname");
            SetDataOnCell(numberColumnHeader, "D", "fname");
            SetDataOnCell(numberColumnHeader, "E", "mname");
            SetDataOnCell(numberColumnHeader, "F", "gender");
            SetDataOnCell(numberColumnHeader, "G", "city");
            SetDataOnCell(numberColumnHeader, "H", "phone");
            SetDataOnCell(numberColumnHeader, "I", "position");
            SetDataOnCell(numberColumnHeader, "J", "manager_");
            SetDataOnCell(numberColumnHeader, "K", "login");
            SetDataOnCell(numberColumnHeader, "L", "password");
            SetDataOnCell(numberColumnHeader, "M", "my_field");
        }

        /// <summary>
        /// Method for adding data
        /// </summary>
        /// <param name="row">Number row</param>
        /// <param name="column">In Latin, like columns in Excel</param>
        /// <param name="data">Input string data</param>
        void SetDataOnCell(int row, string column, string data) => workSheet.Cells[row, column] = data;

        /// <summary>
        /// Entering all workers from the list in Excel
        /// </summary>
        void AddData()
        {
            if(workers.Count != 0)
            {
                int row = 1;
                foreach (var worker in workers)
                {
                    row++;
                    SetDataOnCell(row, "B", worker.Email);
                    SetDataOnCell(row, "C", worker.LastName);
                    SetDataOnCell(row, "D", worker.FirstName);
                    SetDataOnCell(row, "L", worker.Password);
                }

                for (int i = 1; i <= 13; i++)
                {
                    workSheet.Columns[i].AutoFit();
                    ((Excel.Range) workSheet.Columns[i]).AutoFit();
                }
            }
        }

        /// <summary>
        /// Method for getting the name and path of saving the file
        /// </summary>
        /// <returns>File save path</returns>
        string GetPath()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "CSV Files(*.csv)|*.csv|All(*.*)|*";
            dialog.RestoreDirectory = true;
            if (dialog.ShowDialog() == true)
                return dialog.FileName;
            return null;
        }

        /// <summary>
        /// Method for adding entered data to List<Worker>
        /// </summary>
        private void Add(object sender, RoutedEventArgs e)
        {
            if (ChekValidation() && ChekNull())
            {
                workers.Add(new Worker()
                {
                    Email = selectWorker.Email,
                    LastName = selectWorker.LastName,
                    FirstName = selectWorker.FirstName,
                    Password = selectWorker.Password,
                });
                ClearTextBox();
                selectWorker = new Worker();
            }
            else
            {
                MessageBox.Show("Поля должны быть заполнены кореектными данные и не должны быть пусты.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Clearing the entered data
        /// </summary>
        private void ClearTextBox()
        {
            TextBoxLastName.Text = "";
            TextBoxFirstName.Text = "";
            TextBoxEmail.Text = "";
            TextBoxPassword.Text = "";
        }

        /// <summary>
        /// Data validation check
        /// </summary>
        bool ChekValidation()
        {
            return !Validation.GetHasError(TextBoxLastName) && !Validation.GetHasError(TextBoxFirstName) &&
                   !Validation.GetHasError(TextBoxEmail) && !Validation.GetHasError(TextBoxPassword);
        }

        /// <summary>
        /// Checking for the absence of empty fields
        /// </summary>
        bool ChekNull()
        {
            return TextBoxLastName.Text != "" && TextBoxFirstName.Text != "" &&
                   TextBoxEmail.Text != "" && TextBoxPassword.Text != "";
        }

        /// <summary>
        /// Show List<Worker> in Excel
        /// </summary>
        private void Show(object sender, RoutedEventArgs e)
        {
            CreateExcel();
            AddData();
        }

        private void TextBoxEmail_Error(object sender, ValidationErrorEventArgs e)
        {
            if (Validation.GetHasError(sender as TextBox))
            {
                errorEmailText.Text = e.Error.ErrorContent.ToString();
                errorEmailText.Visibility = Visibility.Visible;
            }
            else
            {
                errorEmailText.Visibility = Visibility.Hidden;
            }
        }

        private void TextBoxLastName_Error(object sender, ValidationErrorEventArgs e)
        {
            if (Validation.GetHasError(sender as TextBox))
            {
                errorLnameText.Text = e.Error.ErrorContent.ToString();
                errorLnameText.Visibility = Visibility.Visible;
            }
            else
            {
                errorLnameText.Visibility = Visibility.Hidden;
            }
        }

        private void TextBoxFirstName_Error(object sender, ValidationErrorEventArgs e)
        {
            if (Validation.GetHasError(sender as TextBox))
            {
                errorFnameText.Text = e.Error.ErrorContent.ToString();
                errorFnameText.Visibility = Visibility.Visible;
            }
            else
            {
                errorFnameText.Visibility = Visibility.Hidden;
            }
        }

        private void TextBoxPassword_Error(object sender, ValidationErrorEventArgs e)
        {
            if (Validation.GetHasError(sender as TextBox))
            {
                errorPasswordText.Text = e.Error.ErrorContent.ToString();
                errorPasswordText.Visibility = Visibility.Visible;
            }
            else
            {
                errorPasswordText.Visibility = Visibility.Hidden;
            }
        }
    }
}
