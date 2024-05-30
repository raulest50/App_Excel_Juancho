using Avalonia.Controls;
using OfficeOpenXml;
using System.IO;
using System.Linq;

using System.Threading.Tasks;

namespace ModExcelApp
{


    /*
     * Marca: A
     * Foto: C-G
     * CTNS: U
     * Cantidad:V
     * Precio RMB:W
     * 
     * Total RMB: AD
     * CMB Caja: AI
     * CMB Total: AJ
     * 
     * --------
     * 
     * cto: AO
     * sugerido: AP
     * 
     */



    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            InitDefaultAllTextBoxes();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private async void HandlerStart(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            string fpath = tb_file_path.Text;
            await ProcesarExcel(fpath);
        }

        private async void HandlerBrowse(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filters.Add(new FileDialogFilter() { Name = "Excel Files", Extensions = { "xlsx", "xls" } });

            var result = await openFileDialog.ShowAsync(this);
            if (result != null && result.Length > 0)
            {
                tb_file_path.Text = result[0];
            }
        }


        // funcion principal
        public async Task ProcesarExcel( string file_path)
        {
            using (var package = new ExcelPackage(new FileInfo(file_path)))
            {
                ExcelWorkbook wb = package.Workbook;
                DeleteAllExceptDespacho(wb);

                ExcelWorksheet mdh, tin, like, alma;

                mdh = CreateSheetIfNotExist(wb, "MDH");
                tin = CreateSheetIfNotExist(wb, "TIN");
                like = CreateSheetIfNotExist(wb, "LIKE");
                alma = CreateSheetIfNotExist(wb, "ALMA");

                // Save the changes to the file
                await Task.Run(() => package.Save());
            }

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

        public ExcelWorkbook LoadWorkbook(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                // Handle error: file path is invalid or file does not exist
                return null;
            }

            var package = new ExcelPackage(new FileInfo(filePath));
            return package.Workbook;
        }


        public static void DeleteAllExceptDespacho(ExcelWorkbook workbook)
        {
            // Disable display alerts
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Collect all sheets except the one named "Despacho"
            var sheetsToDelete = workbook.Worksheets
                .Where(sheet => sheet.Name != "Despacho")
                .ToList();

            // Delete collected sheets
            foreach (var sheet in sheetsToDelete)
            {
                workbook.Worksheets.Delete(sheet);
            }
        }
        
        public static ExcelWorksheet CreateSheetIfNotExist(ExcelWorkbook workbook, string sheetName)
        {
            // Check if the sheet already exists
            var existingSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);
            if (existingSheet != null)
            {
                return existingSheet;
            }

            // If the sheet does not exist, create it
            var newSheet = workbook.Worksheets.Add(sheetName);
            return newSheet;
        }

        private void InitDefaultAllTextBoxes()
        {
            tb_col_marca.Text = "A";
            tb_col_foto_i.Text = "C";
            tb_col_foto_e.Text = "G";
            tb_col_ctns.Text = "U";
            tb_col_cantidad.Text = "V";
            tb_col_precio_rmb.Text = "W";
            tb_col_total_rmb.Text = "AD";
            tb_cbm_caja.Text = "AI";
            tb_col_cbm_total.Text = "AJ";
            //tb_cto_mdh.Text = "";
            //tb_cto_tin.Text ="";
            //tb_cto_like.Text = "";
            //tb_cto_alma.Text = "";
        }



    }
}