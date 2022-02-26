using Avtorizaci.Class;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Application = System.Windows.Forms.Application;

namespace Avtorizaci.Forms
{
    public partial class DOCExidAO_1List2 : Form
    {
        static public Excel.Worksheet xlSheet;//Лист
        static public Excel.Range xlSheetRange;//Выделеная область
        static public int count2 = 0;
        static public int count3 = 63;
        Excel.Range excelCells999;
        Excel.Range excelCells998;

        public DOCExidAO_1List2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = String.Empty;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = saveFileDialog1.FileName;
                string path = Path.Combine(Application.StartupPath, @"ao-1.xls");
                Excel.Application exApp = new Excel.Application();
                Workbook book = exApp.Workbooks.Open(path);
                Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                xlSheet = (Excel.Worksheet)exApp.Sheets[1];

                Excel.Range excelCells39 = xlSheet.Range[xlSheet.Cells[7, 1], xlSheet.Cells[7, 1]];
                xlSheet.Cells[7, 1] = "ООО СМП-708 СК г.Мурманск";
                excelCells39.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                Excel.Range excelCells2 = xlSheet.Range[xlSheet.Cells[13, 21], xlSheet.Cells[13, 21]];
                xlSheet.Cells[13, 21] = Class1.NomerFormAO1;
                excelCells2.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                Excel.Range excelCells3 = xlSheet.Range[xlSheet.Cells[1, 1], xlSheet.Cells[1, 1]];
                xlSheet.Cells[13, 26] = Class1.DateFormAO1.ToString("dd.MM.yyyy");

                Excel.Range excelCells4 = xlSheet.Range[xlSheet.Cells[19, 16], xlSheet.Cells[19, 16]];
                xlSheet.Cells[19, 16] = Class1.StrukturnoePodrazdelenie;

                Excel.Range excelCells5 = xlSheet.Range[xlSheet.Cells[20, 11], xlSheet.Cells[20, 11]];
                xlSheet.Cells[20, 11] = Class1.PodotchetnoeFace;

                Excel.Range excelCell6 = xlSheet.Range[xlSheet.Cells[22, 14], xlSheet.Cells[22, 14]];
                xlSheet.Cells[22, 14] = Class1.Proffecie;

                Excel.Range excelCell7 = xlSheet.Range[xlSheet.Cells[22, 35], xlSheet.Cells[22, 35]];
                xlSheet.Cells[22, 35] = Class1.Naznacenieavanca;

                Excel.Range excelCell8 = xlSheet.Range[xlSheet.Cells[27, 18], xlSheet.Cells[27, 18]];
                xlSheet.Cells[27, 18] = Class1.avans;

                Excel.Range excelCell9 = xlSheet.Range[xlSheet.Cells[31, 18], xlSheet.Cells[31, 18]];
                xlSheet.Cells[31, 18] = Class1.avans;

                Excel.Range excelCell10 = xlSheet.Range[xlSheet.Cells[32, 18], xlSheet.Cells[32, 18]];
                xlSheet.Cells[32, 18] = Class1.izracxodovano;

                if (Convert.ToDouble(Class1.avans) >= Convert.ToDouble(Class1.izracxodovano))
                {
                    Class1.itogSumFormAO1 = Convert.ToDouble(Class1.avans) - Convert.ToDouble(Class1.izracxodovano);
                    Class1.Ostatok = Convert.ToDouble(Class1.itogSumFormAO1);
                    Class1.Pereracvod = 0;
                }
                else if (Convert.ToDouble(Class1.avans) <= Convert.ToDouble(Class1.izracxodovano))
                {
                    Class1.itogSumFormAO1 = Convert.ToDouble(Class1.izracxodovano) - Convert.ToDouble(Class1.avans);
                    Class1.Pereracvod = Convert.ToDouble(Class1.itogSumFormAO1);
                    Class1.Ostatok = 0;
                }
                Excel.Range excelCell11 = xlSheet.Range[xlSheet.Cells[33, 18], xlSheet.Cells[33, 18]];
                xlSheet.Cells[33, 18] = Class1.Ostatok;

                Excel.Range excelCell12 = xlSheet.Range[xlSheet.Cells[34, 18], xlSheet.Cells[34, 18]];
                xlSheet.Cells[34, 18] = Class1.Pereracvod;

                int i = 0, j = 0;
                for (i = 0; i <= dataGridView1.RowCount - 2; i++)
                {
                    count2++;
                    count3++;
                    for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                    {
                        wsh.Cells[i + 63, j + 6] = dataGridView1[0, i].Value.ToString();
                        if (dataGridView1.ColumnCount > 2)
                            wsh.Cells[i + 63, j + 11] = dataGridView1[1, i].Value.ToString();
                        if (dataGridView1.ColumnCount > 2)
                            wsh.Cells[i + 63, j + 16] = dataGridView1[2, i].Value.ToString();
                        if (dataGridView1.ColumnCount > 3)
                            wsh.Cells[i + 63, j + 25] = dataGridView1[3, i].Value.ToString();
                        if (dataGridView1.ColumnCount > 4)
                            wsh.Cells[i + 63, j + 31] = dataGridView1[4, i].Value.ToString();


                        excelCells999 = xlSheet.Range[xlSheet.Cells[i + 63, 1], xlSheet.Cells[i + 63, 1]];
                        xlSheet.Cells[i + 63, 1] = count2;
                        excelCells999 = (Excel.Range)wsh.get_Range("A" + i + 63, "E" + i + 63).Cells;
                        excelCells999.Merge(Type.Missing);
                        wsh.Cells[i + 63, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        excelCells999.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        excelCells999.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        excelCells999.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        excelCells999.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        excelCells998 = xlSheet.Range[xlSheet.Cells[i + 63, 6], xlSheet.Cells[i + 63, 6]];
                        excelCells998 = (Excel.Range)wsh.get_Range("F" + i + 63, "J" + i + 63).Cells;
                        excelCells998.Merge(Type.Missing);

                        wsh.Cells[i + 63, j + 6].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 6].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 6].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 6].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                        excelCells998.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        excelCells998.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        excelCells998.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        excelCells998.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        Excel.Range excelCells997 = xlSheet.Range[xlSheet.Cells[i + 63, 11], xlSheet.Cells[i + 63, 11]];
                        excelCells997 = (Excel.Range)wsh.get_Range("K" + i + 63, "O" + i + 63).Cells;
                        excelCells997.Merge(Type.Missing);

                        wsh.Cells[i + 63, j + 11].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 11].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 11].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 11].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                        excelCells997.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        excelCells997.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        excelCells997.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        excelCells997.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        Excel.Range excelCells996 = xlSheet.Range[xlSheet.Cells[i + 63, 16], xlSheet.Cells[i + 63, 16]];
                        excelCells996 = (Excel.Range)wsh.get_Range("P" + i + 63, "X" + i + 63).Cells;
                        excelCells996.Merge(Type.Missing);

                        wsh.Cells[i + 63, j + 16].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 16].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 16].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 16].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        wsh.Cells[i + 63, j + 20].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 20].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 20].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 20].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                        excelCells996.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        excelCells996.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        excelCells996.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        excelCells996.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        Excel.Range excelCells995 = xlSheet.Range[xlSheet.Cells[i + 63, 25], xlSheet.Cells[i + 63, 25]];
                        excelCells995 = (Excel.Range)wsh.get_Range("Y" + i + 63, "AD" + i + 63).Cells;
                        excelCells995.Merge(Type.Missing);

                        wsh.Cells[i + 63, j + 25].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 25].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 25].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 25].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        wsh.Cells[i + 63, j + 28].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 28].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 28].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 28].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                        excelCells995.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        excelCells995.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        excelCells995.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        excelCells995.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        Excel.Range excelCells994 = xlSheet.Range[xlSheet.Cells[i + 63, 31], xlSheet.Cells[i + 63, 31]];
                        excelCells994 = (Excel.Range)wsh.get_Range("AE" + i + 63, "AJ" + i + 63).Cells;
                        excelCells994.Merge(Type.Missing);

                        wsh.Cells[i + 63, j + 31].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 31].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 31].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 31].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        wsh.Cells[i + 63, j + 34].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 34].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 34].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 34].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        excelCells994.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        excelCells994.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        excelCells994.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        excelCells994.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        Excel.Range excelCells10002 = xlSheet.Range[xlSheet.Cells[i + 63, 37], xlSheet.Cells[i + 63, 37]];
                        excelCells10002 = (Excel.Range)wsh.get_Range("AK" + i + 63, "AP" + i + 63).Cells;
                        excelCells10002.Merge(Type.Missing);

                        wsh.Cells[i + 63, j + 37].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 37].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 37].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 37].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        wsh.Cells[i + 63, j + 40].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 40].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 40].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 40].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        excelCells10002.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        excelCells10002.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        excelCells10002.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        excelCells10002.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        Excel.Range excelCells10003 = xlSheet.Range[xlSheet.Cells[i + 63, 43], xlSheet.Cells[i + 63, 43]];
                        excelCells10003 = (Excel.Range)wsh.get_Range("AQ" + i + 63, "AV" + i + 63).Cells;
                        excelCells10003.Merge(Type.Missing);

                        wsh.Cells[i + 63, j + 43].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 43].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 43].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 43].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        wsh.Cells[i + 63, j + 46].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 46].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 46].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 46].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                        excelCells10003.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        excelCells10003.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        excelCells10003.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        excelCells10003.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        Excel.Range excelCells10004 = xlSheet.Range[xlSheet.Cells[i + 63, 49], xlSheet.Cells[i + 63, 49]];
                        excelCells10004 = (Excel.Range)wsh.get_Range("AW" + i + 63, "BC" + i + 63).Cells;
                        excelCells10004.Merge(Type.Missing);

                        wsh.Cells[i + 63, j + 49].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 49].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 49].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 49].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        wsh.Cells[i + 63, j + 51].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        wsh.Cells[i + 63, j + 51].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        wsh.Cells[i + 63, j + 51].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        wsh.Cells[i + 63, j + 51].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                        excelCells10004.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                        excelCells10004.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                        excelCells10004.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                        excelCells10004.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                    }
                }

                Excel.Range excelCell1000 = xlSheet.Range[xlSheet.Cells[i + 63, 16], xlSheet.Cells[i + 63, 16]];
                xlSheet.Cells[i + 63, 16] = "Итого:";
                excelCell1000 = (Excel.Range)wsh.get_Range("P" + i + 63, "X" + i + 63).Cells;
                excelCell1000.Merge(Type.Missing);

                string AD = "Y";
                string noner = Convert.ToString(count2 + 62);
                string itog = AD + noner;
                Excel.Range excelCell1001 = xlSheet.Range[xlSheet.Cells[i + 63, 25], xlSheet.Cells[i + 63, 25]];
                xlSheet.Cells[i + 63, 25] = "=" + "SUM" + "(Y63:" + itog;
                excelCell1001 = (Excel.Range)wsh.get_Range("P" + i + 63, "X" + i + 63).Cells;
                excelCell1001.Merge(Type.Missing);
                for (int i1 = 0; i1 < 24; i1++)
                {
                    wsh.Cells[i + 63, 25 + i1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                    wsh.Cells[i + 63, 25 + i1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                    wsh.Cells[i + 63, 25 + i1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                    wsh.Cells[i + 63, 25 + i1].Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                    excelCell1000.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                    excelCell1000.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                    excelCell1000.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                    excelCell1000.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                }
                int position = i + 65;
                int position2 = i + 66;
                Excel.Range excelCell1003 = xlSheet.Range[xlSheet.Cells[i + 65, 1], xlSheet.Cells[i + 65, 1]];
                xlSheet.Cells[i + 65, 1] = "Подотчетное лицо";

                Excel.Range excelCell1066 = xlSheet.Range[xlSheet.Cells[i + 65, 1], xlSheet.Cells[i + 65, 1]];
                excelCell1066 = (Excel.Range)wsh.get_Range("A" + position + ":E" + position).Cells;
                excelCell1066.UnMerge();

                Excel.Range excelCell1067 = xlSheet.Range[xlSheet.Cells[i + 65, 6], xlSheet.Cells[i + 65, 6]];
                excelCell1067 = (Excel.Range)wsh.get_Range("E" + position + ":J" + position).Cells;
                excelCell1067.UnMerge();

                Excel.Range excelCell10553 = xlSheet.Range[xlSheet.Cells[i + 65, 1], xlSheet.Cells[i + 65, 1]];
                excelCell10553 = (Excel.Range)wsh.get_Range("A" + position, "J" + position).Cells;
                excelCell10553.Merge(Type.Missing);


                Excel.Range excelCell10756 = xlSheet.Range[xlSheet.Cells[i + 65, 11], xlSheet.Cells[i + 65, 11]];
                excelCell10756 = (Excel.Range)wsh.get_Range("K" + position + ":O" + position).Cells;
                excelCell10756.UnMerge();

                Excel.Range excelCell10674 = xlSheet.Range[xlSheet.Cells[i + 65, 16], xlSheet.Cells[i + 65, 16]];
                excelCell10674 = (Excel.Range)wsh.get_Range("P" + position + ":X" + position).Cells;
                excelCell10674.UnMerge();

                Excel.Range excelCell1075 = xlSheet.Range[xlSheet.Cells[i + 65, 11], xlSheet.Cells[i + 65, 11]];
                xlSheet.Cells[i + 65, 11] = "";
                excelCell1075 = (Excel.Range)wsh.get_Range("K" + position, "X" + position).Cells;
                excelCell1075.Merge(Type.Missing);
                excelCell1075.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                Excel.Range excelCell10754 = xlSheet.Range[xlSheet.Cells[i + 66, 11], xlSheet.Cells[i + 66, 11]];
                excelCell10756 = (Excel.Range)wsh.get_Range("K" + position2 + ":O" + position2).Cells;
                excelCell10756.UnMerge();

                Excel.Range excelCell10672 = xlSheet.Range[xlSheet.Cells[i + 66, 16], xlSheet.Cells[i + 66, 16]];
                excelCell10674 = (Excel.Range)wsh.get_Range("P" + position2 + ":X" + position2).Cells;
                excelCell10674.UnMerge();

                Excel.Range excelCell10765 = xlSheet.Range[xlSheet.Cells[i + 66, 11], xlSheet.Cells[i + 66, 11]];
                xlSheet.Cells[i + 66, 11] = "";
                excelCell10765 = (Excel.Range)wsh.get_Range("K" + position2, "X" + position2).Cells;
                excelCell10765.Merge(Type.Missing);

                Excel.Range excelCell10033 = xlSheet.Range[xlSheet.Cells[i + 66, 11], xlSheet.Cells[i + 66, 11]];
                xlSheet.Cells[i + 66, 11] = "(подпись)";

                Excel.Range excelCell107 = xlSheet.Range[xlSheet.Cells[i + 65, 31], xlSheet.Cells[i + 65, 31]];
                xlSheet.Cells[i + 65, 31] = Class1.PodotchetnoeFace;
                excelCell107 = (Excel.Range)wsh.get_Range("AE" + position, "AP" + position).Cells;
                excelCell107.Merge(Type.Missing);
                excelCell107.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                Excel.Range excelCell102 = xlSheet.Range[xlSheet.Cells[i + 66, 31], xlSheet.Cells[i + 66, 31]];
                excelCell102 = (Excel.Range)wsh.get_Range("AE" + position2 + ":AJ" + position2).Cells;
                excelCell102.UnMerge();

                Excel.Range excelCell10671 = xlSheet.Range[xlSheet.Cells[i + 66, 37], xlSheet.Cells[i + 66, 37]];
                excelCell10671 = (Excel.Range)wsh.get_Range("AK" + position2 + ":AP" + position2).Cells;
                excelCell10671.UnMerge();

                Excel.Range excelCell10760 = xlSheet.Range[xlSheet.Cells[i + 66, 31], xlSheet.Cells[i + 66, 31]];
                xlSheet.Cells[i + 66, 31] = "";
                excelCell10760 = (Excel.Range)wsh.get_Range("AE" + position2, "AP" + position2).Cells;
                excelCell10760.Merge(Type.Missing);

                Excel.Range excelCell100312 = xlSheet.Range[xlSheet.Cells[i + 66, 31], xlSheet.Cells[i + 66, 31]];
                xlSheet.Cells[i + 66, 31] = "(расшифровка подписи)";
                path = saveFileDialog1.FileName;
                book.SaveAs(saveFileDialog1.FileName);
                book.Close(@"ao-1.xls");
                Process[] ps2 = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (Process p2 in ps2)
                {
                    p2.Kill();

                }
            }
        }
           
        private void button5_Click(object sender, EventArgs e)
        {
            DOCExidAO_1 f = new DOCExidAO_1();
            f.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add(dateTimePicker1.Text, textBox3.Text, comboBox2.Text, textBox1.Text, textBox2.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {

            try
            {
                    int a = dataGridView1.CurrentRow.Index;
                dataGridView1.Rows.Remove(dataGridView1.Rows[a]);
                

            }
            catch
            {
                DialogResult result = MessageBox.Show(
                "Все записи удалены", "Предупреждение!",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1);
            }
        }

        private void DOCExidAO_1List2_Load(object sender, EventArgs e)
        {
           
            Class1.Get();
            
            Class1.Getdocument();
            comboBox2.DataSource = Class1.dtspdocumen;
            comboBox2.ValueMember = "Nomer_doc";
            comboBox2.DisplayMember = "Documen";

        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }
    }
}
