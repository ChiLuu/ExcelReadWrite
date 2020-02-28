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
using TextBox = System.Windows.Controls.TextBox;

namespace ExcelReadWrite
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        Excel excelFile1 = new Excel(); // First Excel file
        Excel excelFile2 = new Excel(); // Second Excel file
        List<string> workSheet1ColumnNames = new List<string>(); // List to hold all column names for first worksheet
        List<string> workSheet2ColumnNames = new List<string>(); // List to hold all column names for second worksheet

        /// <summary>
        /// When applicat starts up
        /// </summary>
        /// <returns>No returns</returns>
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Set an Excel object's properties and set textbox to display name of file selected.
        /// </summary>
        /// <param name="excelFile">Excel object</param>
        /// <param name="filePathTextBox">TextBox that will display file name</param>
        /// <returns>True if file select, false cancel selection</returns>
        private bool setExcelFile(Excel excelFile, TextBox filePathTextBox)
        {
            string path;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;" +
                                    "*.xls;*.xlt;*.xls;*.xml;*.xml;*.xlam;*.xla;*.xlw;*.xlr;";

            if (openFileDialog.ShowDialog() == true)
            {
                path = openFileDialog.FileName; // Absolute path of file selected.
                filePathTextBox.Text = openFileDialog.SafeFileName; // File name with extension.
                excelFile.InitializeExcel(path); // Set properties.
                return true;
            }
            return false;
        }

        /// <summary>
        /// Opens up a OpenFileDialog window to allows user to select first Excel File.
        /// </summary>
        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            // If user selected an Excel file, add worksheets to list box's display.
            if (setExcelFile(excelFile1, txtPath1))
            {
                lbWorkSheets1.Items.Clear();
                foreach (Worksheet worksheet in excelFile1.Workbook.Worksheets)
                {
                    lbWorkSheets1.Items.Add(worksheet.Name);
                }
            }
        }

        /// <summary>
        /// Opens up a OpenFileDialog window to allows user to select second Excel File.
        /// </summary>
        private void btnOpenFile2_Click(object sender, RoutedEventArgs e)
        {
            // If user selected an Excel file, add worksheets to list box's display.
            if (setExcelFile(excelFile2, txtPath2))
            {
                lbWorkSheets2.Items.Clear();
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
        }

        private void lbWorkSheets2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cbWorkBook2.ItemsSource = "";
            txtWS2ColumnNameStartOnRow.Text = "";
            excelFile2.Worksheet = excelFile2.Workbook.Worksheets[lbWorkSheets2.SelectedIndex + 1];
        }

        private void DisplayColumnNames(Excel excelFile, TextBox rowStart, List<String> columnNameList)
        {
                Worksheet activeWorkSheet = (Worksheet)excelFile.Worksheet;
                int column = 1;
                int row = Convert.ToInt32(rowStart.Text);
                int lastColumnNumber = activeWorkSheet.UsedRange.Columns.Count;
                columnNameList.Clear();

                if (activeWorkSheet.Cells[row, column].Value != null)
                {
                    string currentCell = activeWorkSheet.Cells[row, column].Value.ToString();

                    while (!string.IsNullOrEmpty(currentCell) && column <= lastColumnNumber)
                    {
                        columnNameList.Add(currentCell);
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
                    rowStart.Focus();
                }
        }
        private void clearColumnList(ComboBox comboBox, List<string> columnList)
        {
            comboBox.ItemsSource = "";
            columnList.Clear();
        }
        private void txtWS1ColumnNameStartOnRow_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (Validator.IsPresent(txtWS1ColumnNameStartOnRow, "Row Number") &&
                Validator.IsInt32(txtWS1ColumnNameStartOnRow, "Row Number"))
            {
                clearColumnList(cbWorkBook1, workSheet1ColumnNames);
                btnGetColumn1.IsEnabled = true;
            }
            else
            {
                clearColumnList(cbWorkBook1, workSheet1ColumnNames);
                btnGetColumn1.IsEnabled = false;
            }
                
        }

        private void txtWS2ColumnNameStartOnRow_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (Validator.IsPresent(txtWS2ColumnNameStartOnRow, "Row Number") &&
                Validator.IsInt32(txtWS2ColumnNameStartOnRow, "Row Number"))
            {
                clearColumnList(cbWorkBook2, workSheet2ColumnNames);
                btnGetColumn2.IsEnabled = true;
            }
            else
            {
                clearColumnList(cbWorkBook2, workSheet2ColumnNames);
                btnGetColumn2.IsEnabled = false;
            }
                
        }

        private void btnGetColumn1_Click(object sender, RoutedEventArgs e)
        {
            DisplayColumnNames(excelFile1, txtWS1ColumnNameStartOnRow, workSheet1ColumnNames);

            if (workSheet1ColumnNames.Count > 0)
            {
                cbWorkBook1.ItemsSource = workSheet1ColumnNames;
                cbWorkBook1.SelectedIndex = 0;
            }
        }

        private void btnGetColumn2_Click(object sender, RoutedEventArgs e)
        {
            cbWorkBook2.ItemsSource = "";
            DisplayColumnNames(excelFile2, txtWS2ColumnNameStartOnRow, workSheet2ColumnNames);

            if (workSheet2ColumnNames.Count > 0)
            {
                cbWorkBook2.ItemsSource = workSheet2ColumnNames;
                cbWorkBook2.SelectedIndex = 0;
            }
        }
    } // End class
} // End Namespace
