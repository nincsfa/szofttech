using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
namespace T0KW0U_12GYAK
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp; // A Microsoft Excel alkalmazás
        Excel.Workbook xlWB;     // A létrehozott munkafüzet
        Excel.Worksheet xlSheet; // Munkalap a munkafüzeten belül

        Models.HajosContext hajosContext = new Models.HajosContext();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                // Excel elindítása és az applikáció objektum betöltése
                xlApp = new Excel.Application();

                // Új munkafüzet
                xlWB = xlApp.Workbooks.Add(Missing.Value);

                // Új munkalap
                xlSheet = xlWB.ActiveSheet;

                // Tábla létrehozása
                CreateTable(); // Ennek megírása a következõ feladatrészben következik

                // Control átadása a felhasználónak
                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex) // Hibakezelés a beépített hibaüzenettel
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                // Hiba esetén az Excel applikáció bezárása automatikusan
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }
        void CreateTable()
        {
            string[] fejlécek = new string[] {
        "Kérdés",
        "1. válasz",
        "2. válaszl",
        "3. válasz",
        "Helyes válasz",
        "kép"};
            Models.HajosContext hajosContext = new Models.HajosContext();
            var mindenKérdés = hajosContext.Questions.ToList();

            object[,] adatTömb = new object[mindenKérdés.Count(),6];

            for (int i = 0; i < mindenKérdés.Count(); i++)
            {
                adatTömb[i, 0] = mindenKérdés[i].Question1;
                adatTömb[i, 1] = mindenKérdés[i].Answer1;
                adatTömb[i, 2] = mindenKérdés[i].Answer2;
                adatTömb[i, 3] = mindenKérdés[i].Answer3;
                adatTömb[i, 4] = mindenKérdés[i].CorrectAnswer;
                adatTömb[i, 5] = mindenKérdés[i].Image;
            }
            int sorokSzáma = adatTömb.GetLength(0);
            int oszlopokSzáma = adatTömb.GetLength(1);
            Excel.Range adatRange = xlSheet.get_Range("A2", Type.Missing).get_Resize(sorokSzáma, oszlopokSzáma);
            adatRange.Value2 = adatTömb;
            adatRange.Columns.AutoFit();

            Excel.Range fejlécRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(1, 6);
            fejlécRange.Font.Bold = true;
            fejlécRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            fejlécRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            fejlécRange.EntireColumn.AutoFit();
            fejlécRange.RowHeight = 40;
            fejlécRange.Interior.Color = Color.Fuchsia;
            fejlécRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            int lastRowID = xlSheet.UsedRange.Rows.Count;

            adatRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range elsooszlopRange = xlSheet.get_Range("A2", Type.Missing).get_Resize(lastRowID, 1);
            elsooszlopRange.Font.Bold = true;
            elsooszlopRange.Interior.Color = Color.LightYellow;

            Excel.Range utolsooszlopRange = xlSheet.get_Range("F2", Type.Missing).get_Resize(lastRowID, 1);
            utolsooszlopRange.Interior.Color = Color.LimeGreen;

            Excel.Range utolsoelottioszlopRange = xlSheet.get_Range("E2", Type.Missing).get_Resize(lastRowID, 1);
            utolsoelottioszlopRange.NumberFormat = "0.00";
        }
    }
}