using DataDrivenCustomMailer.MessageReaders;
using EmailService;
using ExcelDataExchanger;
using ExcelReader;
using ExcelWriter;
using LAPreparator.BusinessLogic;
using LAPreparator.DataAccess;
using LAPreparator.Serivices;
using LAPreparator.Services;
using LAPreparator.UIComponents;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

namespace LAPreparator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string resourceDirectory;
        LoadingAdviceControl laControl;
        public MainWindow()
        {
            InitializeComponent();
            resourceDirectory = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "Resources");
        }

        private void ConfigurationMenuButton_Click(object sender, RoutedEventArgs e)
        {
            ConfigurationMenuStack.Visibility = ConfigurationMenuStack.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;


        }

        private void EmailMenuButton_Click(object sender, RoutedEventArgs e)
        {
            EmailMenuStack.Visibility = EmailMenuStack.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
        }

        private void SendLaButton_Click(object sender, RoutedEventArgs e)
        {
            IExcelDataReader excelDataReader = new ExcelDataReader();
            IExcelDataWriter excelDataWriter = new ExcelDataWriter(); ;
            IExchanger exchanger = new ExcelDataExchanger.ExcelDataExchanger();
            ExcelEmailAddressService addressService = new ExcelEmailAddressService();
            ITemplateReader msgReader = new MsgTemplateReader();
            EmailContractor emailContractor = new EmailContractor(addressService, msgReader);
            EmailService.IEmailService emailService = new EmailService.EmailService();
            ILaCreator laCreator = new LaCreator(excelDataWriter, exchanger, emailContractor, emailService);
            if (laControl is null)
            {

                laControl = new LoadingAdviceControl(excelDataReader, laCreator);
                BodyPanel.Children.Add(laControl);
            }
        }


        private void MessageBodyButton_Click(object sender, RoutedEventArgs e)
        {
            var templatePath = System.IO.Path.Combine(resourceDirectory, "template.msg");
            try
            {
                TemplateViewer.DisplayMessage(templatePath);
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void AddressBookButton_Click(object sender, RoutedEventArgs e)
        {
            var templatePath = System.IO.Path.Combine(resourceDirectory, "AddressBook.xlsx");
            ProcessStartInfo pInfo = new ProcessStartInfo()
            {
                UseShellExecute = true,
                FileName = templatePath

            };
            Process.Start(pInfo);
        }
    }
}
