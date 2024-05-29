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


        public double ctoCol(double ctoRMB, double tasa, double cbmCaja, double tasa_m3, double unds_caja)
        {
            return ctoRMB * tasa * 1.05 + cbmCaja * tasa_m3 / unds_caja;
        }

        public double cto_place(double cto_col, string bodega)
        {
            double cto = 1;
            switch(bodega)
            {
                case "MDH":
                    cto = 0.7;
                    break;
                case "TIN":
                    cto = 0.75;
                    break;
                case "LIKE":
                    cto= 0.65;
                    break;
                case "ALMA":
                    cto= 0.7;
                    break;
            }

            return cto_col*2*cto;
        }        
    }
}