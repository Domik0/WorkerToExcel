using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
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
        private List<Worker> workers = new List<Worker>();
        public Worker SelectWorker = new Worker();
        Excel.Application excelApp;
        Excel._Worksheet workSheet;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = SelectWorker;
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            CreateExcel();
            AddData();
            
            string path = GetPath();
            if (path != null)
            {
                //Я честно не знаю почему, но вот эта строчка должна реализовывать сохранение excel
                //в кодировке utf-8.
                excelApp.DefaultWebOptions.Encoding = MsoEncoding.msoEncodingUTF8;
                workSheet.SaveAs(path, Excel.XlFileFormat.xlCSV);
            }
        }

        void CreateExcel()
        {
            excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = "id";
            workSheet.Cells[1, "B"] = "email";
            workSheet.Cells[1, "C"] = "lname";
            workSheet.Cells[1, "D"] = "fname";
            workSheet.Cells[1, "E"] = "mname";
            workSheet.Cells[1, "F"] = "gender";
            workSheet.Cells[1, "G"] = "city";
            workSheet.Cells[1, "H"] = "phone";
            workSheet.Cells[1, "I"] = "position";
            workSheet.Cells[1, "J"] = "manager_";
            workSheet.Cells[1, "K"] = "login";
            workSheet.Cells[1, "L"] = "password";
            workSheet.Cells[1, "M"] = "my_field";
        }

        void AddData()
        {
            var row = 1;
            foreach (var worker in workers)
            {
                row++;
                workSheet.Cells[row, "B"] = worker.email;
                workSheet.Cells[row, "C"] = worker.lname;
                workSheet.Cells[row, "D"] = worker.fname;
                workSheet.Cells[row, "L"] = worker.password;
            }

            for (int i = 1; i <= 13; i++)
            {
                workSheet.Columns[i].AutoFit();
                ((Excel.Range)workSheet.Columns[i]).AutoFit();
            }
        }

        string GetPath()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "CSV Files(*.csv)|*.csv|All(*.*)|*";
            dialog.RestoreDirectory = true;
            if (dialog.ShowDialog() == true)
                return dialog.FileName;
            return null;
        }

        private void Add(object sender, RoutedEventArgs e)
        {
            if (ChekValidation() && ChekNull())
            {
                workers.Add(new Worker()
                {
                    email = TextBoxEmail.Text,
                    lname = TextBoxLname.Text,
                    fname = TextBoxFname.Text,
                    password = TextBoxPassword.Text,
                });
                TextBoxLname.Text = "";
                TextBoxFname.Text = "";
                TextBoxEmail.Text = "";
                TextBoxPassword.Text = "";
                SelectWorker = new Worker();
            }
            else
            {
                MessageBox.Show("Поля должны быть заполнены кореектными данные и не должны быть пусты.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        bool ChekValidation()
        {
            return !Validation.GetHasError(TextBoxLname) && !Validation.GetHasError(TextBoxFname) &&
                   !Validation.GetHasError(TextBoxEmail) && !Validation.GetHasError(TextBoxPassword);
        }

        bool ChekNull()
        {
            return TextBoxLname.Text != "" && TextBoxFname.Text != "" &&
                   TextBoxEmail.Text != "" && TextBoxPassword.Text != "";
        }

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

        private void TextBoxLname_Error(object sender, ValidationErrorEventArgs e)
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

        private void TextBoxFname_Error(object sender, ValidationErrorEventArgs e)
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
