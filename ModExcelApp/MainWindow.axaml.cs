using Avalonia.Controls;
using OfficeOpenXml;

namespace ModExcelApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void OnStart_ButtonStart(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            var vp_mdf = pMDF.Text;

            var dialog = new ModalDialog();
            await dialog.ShowDialog(this);

        }
    }
}