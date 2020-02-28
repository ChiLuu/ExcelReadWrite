using Microsoft.Win32;
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
using Microsoft.Office.Interop.Excel;

namespace ExcelReadWrite
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        Excel excelFile1;
        Excel excelFile2;
        List<string> workSheet1ColumnNames = new List<string>();
        List<string> workSheet2ColumnNames = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            string filePath;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                lbWorkSheets1.Items.Clear();
                filePath = openFileDialog.FileName;
                txtPath.Text = openFileDialog.SafeFileName;

                excelFile1 = new Excel(filePath);

                foreach (Worksheet worksheet in excelFile1.Workbook.Worksheets)
                {
                    lbWorkSheets1.Items.Add(worksheet.Name);
                }
            }
                
        }

        private void btnOpenFile2_Click(object sender, RoutedEventArgs e)
        {
            string filePath2;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                lbWorkSheets2.Items.Clear();
                filePath2 = openFileDialog.FileName;
                txtPath2.Text = openFileDialog.SafeFileName;

                excelFile2 = new Excel(filePath2);

                foreach (Worksheet worksheet in excelFile2.Workbook.Worksheets)
                {
                    lbWorkSheets2.Items.Add(worksheet.Name);
                }
            }
        }

        private void lbWorkSheets1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cbWorkBook1.ItemsSource = "";
            txtWS1ColumnNameStartOnRow.Text = "";
            excelFile1.Worksheet = excelFile1.Workbook.Worksheets[lbWorkSheets1.SelectedIndex + 1];
            txtWS1ColumnNameStartOnRow.IsEnabled = true;
        }

        private void lbWorkSheets2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cbWorkBook2.Items.Clear();
            txtWS2ColumnNameStartOnRow.Text = "";
            excelFile2.Worksheet = excelFile2.Workbook.Worksheets[lbWorkSheets2.SelectedIndex];
            txtWS2ColumnNameStartOnRow.IsEnabled = true;
        }

        private void txtWS1ColumnNameStartOnRow_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(Validator.IsInt32(txtWS1ColumnNameStartOnRow, "Row Number"))
            {
                Worksheet activeWorkSheet = (Worksheet)excelFile1.Worksheet;
                int column = 1;
                int row = Convert.ToInt32(txtWS1ColumnNameStartOnRow.Text);
                int lastColumnNumber = activeWorkSheet.UsedRange.Columns.Count;
                cbWorkBook1.ItemsSource = "";

                if (activeWorkSheet.Cells[row, column].Value != null)
                {
                    string currentCell = activeWorkSheet.Cells[row, column].Value.ToString();

                    while (!string.IsNullOrEmpty(currentCell) && column <= lastColumnNumber)
                    {
                        workSheet1ColumnNames.Add(currentCell);
                        column++;
                        if (activeWorkSheet.Cells[row, column].Value != null)
                            currentCell = activeWorkSheet.Cells[row, column].Value.ToString();
                    }
                }
                else
                {
                    MessageBox.Show("Column at this row is empty. " +
                        "\nPlease make sure column " + column + ", row " + row +
                        "\nis not empty in this Worksheet");
                }
                

                cbWorkBook1.ItemsSource = workSheet1ColumnNames;
            }
        }

        
    }
}
