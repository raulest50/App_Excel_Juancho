using Avalonia.Controls;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using ModExcelApp;

using System.Threading.Tasks;
using Avalonia.Controls.Notifications;
using System;
using Avalonia.Threading;
using OfficeOpenXml.Style;
using System.Drawing;

using System.Diagnostics;
using OfficeOpenXml.Drawing;

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

        private async void Test(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            string fpath = tb_file_path.Text;

            try
            {
                using (var package = new ExcelPackage(new FileInfo(fpath)))
                {
                    ExcelWorkbook wb = package.Workbook;
                    ExcelWorksheet hojaMain = wb.Worksheets["Despacho"];

                    var pics = hojaMain.Drawings.Where(x => x.DrawingType == eDrawingType.Picture).Select(x => x.As.Picture);

                    Debug.WriteLine("Test Output");

                    var pic = pics.ElementAtOrDefault(3);

                    Debug.WriteLine($"name: {pic.Name}");
                    Debug.WriteLine($"tipo: {pic.Image.Type}");
                    Debug.WriteLine($"image bounds: {pic.Image.Bounds.Width}" );
                    Debug.WriteLine($" pic position:  {pic.Border.ToString()} ");

                }
            } catch(Exception ex)
            {

            }
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
            try
            {
                using (var package = new ExcelPackage(new FileInfo(file_path)))
                {
                    
                    ExcelWorkbook wb = package.Workbook;
                    Rutinas.DeleteAllExceptDespacho(wb);

                    ExcelWorksheet mdh, tin, like, alma;

                    mdh = Rutinas.CreateSheetIfNotExist(wb, "MDH");
                    tin = Rutinas.CreateSheetIfNotExist(wb, "TIN");
                    like = Rutinas.CreateSheetIfNotExist(wb, "LIKE");
                    alma = Rutinas.CreateSheetIfNotExist(wb, "ALMA");

                    ExcelWorksheet hojaMain = wb.Worksheets["Despacho"];
                    Rutinas.SetupColTitles(hojaMain, mdh);
                    Rutinas.SetupColTitles(hojaMain, tin);
                    Rutinas.SetupColTitles(hojaMain, like);
                    Rutinas.SetupColTitles(hojaMain, alma);



                    
                    bool not_finished = true;

                    // indices para cada hoja
                    int n = 4;
                    int m = 2;
                    int t = 2;
                    int l = 2;
                    int a = 2;

                    int nc;
                    


                    
                    while (not_finished)
                    {
                        if (Rutinas.IsCellEmpty(hojaMain.Cells[n, 1]))
                        {
                            not_finished = false;
                        }
                        else // si la celda no esta vacia
                        {
                            string bodega_name = hojaMain.Cells[n, 1].Value.ToString();

                            nc = Rutinas.CountMergedCells(hojaMain, n, 1);

                            if (nc > 1) // si las celdas son merged
                            {
                                for (int i = n; i < n + nc; i++)
                                {
                                    switch (bodega_name)
                                    {
                                        case "MDH":
                                            Rutinas.CopiarRecord(hojaMain, mdh, i, m);
                                            m++;
                                            break;
                                        case "TIN":
                                            Rutinas.CopiarRecord(hojaMain, tin, i, t);
                                            t++;
                                            break;
                                        case "LIKE":
                                            Rutinas.CopiarRecord(hojaMain, like, i, l);
                                            l++;
                                            break;
                                        case "ALMA BEAUTY":
                                            Rutinas.CopiarRecord(hojaMain, alma, i, a);
                                            a++;
                                            break;
                                    }
                                }

                                switch (bodega_name)
                                {
                                    case "MDH":
                                        Rutinas.MergeCellsInColumnA(m - nc - 1, m - 2, mdh);
                                        break;
                                    case "TIN":
                                        Rutinas.MergeCellsInColumnA(t - nc - 1, t - 2, tin);
                                        break;
                                    case "LIKE":
                                        Rutinas.MergeCellsInColumnA(l - nc - 1, l - 2, like);
                                        break;
                                    case "ALMA BEAUTY":
                                        Rutinas.MergeCellsInColumnA(a - nc - 1, a - 2, alma);
                                        break;
                                }

                                n += nc;
                            }
                            else // si es una sola celda
                            {
                                switch (bodega_name.ToUpper())
                                {
                                    case "MDH":
                                        Rutinas.CopiarRecord(hojaMain, mdh, n, m);
                                        m++;
                                        break;
                                    case "TIN":
                                        Rutinas.CopiarRecord(hojaMain, tin, n, t);
                                        t++;
                                        break;
                                    case "LIKE":
                                        Rutinas.CopiarRecord(hojaMain, like, n, l);
                                        l++;
                                        break;
                                    case "ALMA BEAUTY":
                                        Rutinas.CopiarRecord(hojaMain, alma, n, a);
                                        a++;
                                        break;
                                }

                                n++;
                            }
                        }
                    }
                    

                    // Save the changes to the file
                    await Task.Run(() => package.Save());
                }
            } catch (Exception ex)
            {
                
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

