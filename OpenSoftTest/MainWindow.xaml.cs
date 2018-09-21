using System;
using System.Collections.Generic;
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
using System.Windows.Forms;


namespace OpenSoftTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btn_browse_Click(object sender, RoutedEventArgs e)
        {
            using(OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel Files (*.xlsx;*.xlsm;*.xls;*.xlsb)|*.xlsx;*.xlsm;*.xls;*.xlsb";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) tbx_path.Text = dialog.FileName;
            }
        }

        private void btn_start_Click(object sender, RoutedEventArgs e)
        {
            Task.Factory.StartNew((path) => performExcelTask(path.ToString()), tbx_path.Text);
        }

        public void performExcelTask(string path)
        {
            using (ExcelScanner scanner = new ExcelScanner())
            {
                scanner.paintWords(path, "ReD", Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed);
            }
        }

    }
}
