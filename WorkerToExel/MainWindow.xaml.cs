using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

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

            var row = 1;
            foreach (var worker in workers)
            {
                row++;
                workSheet.Cells[row, "B"] = Encoding.UTF8.GetString(worker.email);
                workSheet.Cells[row, "C"] = Encoding.UTF8.GetString(worker.lname);
                workSheet.Cells[row, "D"] = Encoding.UTF8.GetString(worker.fname);
                workSheet.Cells[row, "L"] = Encoding.UTF8.GetString(worker.password);
            }

            for (int i = 1; i <= 13; i++)
            {
                workSheet.Columns[i].AutoFit();
                ((Excel.Range)workSheet.Columns[i]).AutoFit();
            }
            
            string path = GetPath();
            if (path != null)
            {
                excelApp.DefaultWebOptions.Encoding = MsoEncoding.msoEncodingUTF8;
                workSheet.SaveAs(path, Excel.XlFileFormat.xlCSV);
            }
        }

        public string GetPath()
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
            if (TextBoxLname.Text != "" && TextBoxFname.Text != ""
                                        && TextBoxEmail.Text != "" && TextBoxPassword.Text != "")
            {
                workers.Add(new Worker()
                {
                    email = Encoding.UTF8.GetBytes(TextBoxEmail.Text),
                    lname = Encoding.UTF8.GetBytes(TextBoxLname.Text),
                    fname = Encoding.UTF8.GetBytes(TextBoxFname.Text),
                    password = Encoding.UTF8.GetBytes(TextBoxPassword.Text),
                });
                TextBoxLname.Text = "";
                TextBoxFname.Text = "";
                TextBoxEmail.Text = "";
                TextBoxPassword.Text = "";
            }
            else
            {
                MessageBox.Show("Все поля должны быть заполнены.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
