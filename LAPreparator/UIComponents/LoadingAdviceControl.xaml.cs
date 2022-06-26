using ExcelDataExchanger;
using ExcelReader;
using ExcelWriter;
using LAPreparator.BusinessLogic;
using LAPreparator.Utilities;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
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

namespace LAPreparator.UIComponents
{
    /// <summary>
    /// Interaction logic for LoadingAdviceControl.xaml
    /// </summary>
    public partial class LoadingAdviceControl : UserControl
    {
        private readonly IExcelDataReader excelDataReader;
        private readonly ILaCreator laCreator;
        private List<IGrouping<string, DataRow>> _groupData;

        public DataTable DataTableCollection { get; set; }


        public LoadingAdviceControl(
            IExcelDataReader excelDataReader,
            ILaCreator laCreator
            )
        {
            InitializeComponent();
            this.excelDataReader = excelDataReader;
            this.laCreator = laCreator;
        }

        private async void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName = FileTextBox.Text;
            string sheetName = SheetNameTextBox.Text;
            string range = RangeTextBox.Text;
            try
            {
                if (string.IsNullOrEmpty(fileName))
                {
                    MessageBox.Show("File name cannot be null.");
                }
                else
                {

                    DataTableCollection = await Task.Run(() => excelDataReader.GetData(fileName, sheetName, range));
                    LoadGroupStack(DataTableCollection);
                    LoadGridView(DataTableCollection);
                }
            }catch(Exception exception)
            {
                MessageBox.Show(exception.Message);
                StreamWriter writer = new StreamWriter(DateTime.Today.ToString("yyyy-MM-dd-hh-mm") + ".txt", true);
                
                writer.WriteLine(exception.StackTrace);
                writer.Flush();
                writer.Close();
            } 

        }

        private void LoadGroupStack(DataTable data)
        {
            GroupByPanel.Children.Clear();
            foreach (DataColumn column in data.Columns)
            {
                RadioButton radioButton = new RadioButton();
                radioButton.Name = "GroupByRadio";
                radioButton.Content = column.ColumnName;
                radioButton.Checked += RadioButton_Checked;
                GroupByPanel.Children.Add(radioButton);
            }
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var radio = (RadioButton)sender;
            if (radio.IsChecked != null && radio.IsChecked == true)
            {
                 _groupData = GroupDataGrid(radio.Content.ToString());
           //     CreateFile(groupData);
            }
        }
    
        private List<IGrouping<string, DataRow>> GroupDataGrid(string content)
        {


            var groupedData = DataTableCollection.AsEnumerable().GroupBy(a => a.Field<string>(content)).ToList();

            return groupedData;

        }

        private void LoadGridView(DataTable data)
        {
            DataGrid.Columns.Clear();
            DataGrid.ItemsSource = null;
            foreach (DataColumn item in data.Columns)
            {
                var col = new DataGridTextColumn();
                col.Header = item.ColumnName;
                col.Binding = new Binding($"[{item.ColumnName}]");
                DataGrid.Columns.Add(col);
            }
            var dataView = new DataView(data);
            DataGrid.ItemsSource = dataView;
        }
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel| *.xls;*.xlsx;";
            openFileDialog.FileOk += OpenFileDialog_FileOk;
            openFileDialog.ShowDialog();


        }

        private void OpenFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            FileTextBox.Text = ((OpenFileDialog)sender).FileName;

        }

        private void DataGrid_Sorting(object sender, DataGridSortingEventArgs e)
        {
            e.Column.SortMemberPath = e.Column.SortMemberPath.Trim('[', ']');

        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            string sourceFileName = FileTextBox.Text;
            string sourceSheetName = SheetNameTextBox.Text;
            string range = RangeTextBox.Text;
            if (_groupData != null)
            {
                try
                {

                    var las = laCreator.CreateLAs(_groupData, sourceFileName, sourceSheetName, range);
                    MessageBox.Show("Done! Opening Log");
                    ReportGenerator.CreateReport(las);
                }catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Select Group");
            }
        }
    }
}
