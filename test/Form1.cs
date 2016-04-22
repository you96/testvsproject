using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

using System.Runtime.InteropServices;
using System.Collections;
using System.IO;
using System.Xml;
using System.Threading;
using System.Diagnostics;

namespace test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        //flush all Excel object dans memoire
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private int getColsLimit(Excel._Worksheet NpWorkSheet)
        {
            int colcount = 24;
            Excel.Range rangeuse = NpWorkSheet.UsedRange;
            object[,] value2 = (object[,])rangeuse.Value;
            int cols = rangeuse.Columns.Count;
            for (int i = 1; i < cols; i++)
            {
                if ((value2[1, i] != null && (value2[1, i].ToString() == "1000" || value2[1, i].ToString() == "2000")) || (value2[2, i] != null && (value2[2, i].ToString() == "1000" || value2[2, i].ToString() == "2000")))
                {
                    colcount = i;
                    break;
                }
            }
            return colcount;
        }
        public string getColIndexText(int numb)
        {
            string text = "";
            if (numb < 27)
            {
                text = Convert.ToChar(numb + 64).ToString();
            }
            else
            {
                int temp = numb / 26;
                int temp2 = numb % 26;
                if (numb != 0)
                {
                    text = Convert.ToChar(temp + 64).ToString() + Convert.ToChar(temp2 + 64).ToString();
                }
                else
                {
                    text = Convert.ToChar(temp - 1 + 64).ToString() + "Z";
                }
            }
            return text;
        }
        private int getRowsLimit(Excel._Worksheet NpWorkSheet)
        {
            int rowcount = 24;
            Excel.Range rangeuse = NpWorkSheet.UsedRange;
            object[,] value2 = (object[,])rangeuse.Value;
            int rows = rangeuse.Rows.Count - 1;
            for (int i = 1; i < rows; i++)
            {
                if ((value2[i, 1] != null && (value2[i, 1].ToString() == "1000" || value2[i, 1].ToString() == "2000")) || (value2[i, 2] != null && (value2[i, 2].ToString() == "1000" || value2[i, 2].ToString() == "2000")))
                {
                    rowcount = i;
                    break;
                }
            }
            return rowcount;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.AppendText("Start decoupage");
            string pathnotapme = @"D:\ptw\notepme";
            //pathstylerfinal =  @textBox12.Text + "\\changeStyle\\divi\\final";

            string openfilex = @"D:\ptw\Histo.xlsx";

            ////////////////open excel///////////////////////////////////////
            Thread.Sleep(3000);
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Workbook xlWorkBookx1;
            Excel.Workbook xlWorkBooknewx1;
            object misValue = System.Reflection.Missing.Value;
            //////////creat modele histox.xls pour fichier diviser////////////////////////////////
            Excel.Application xlAppRef;
            Excel.Workbook xlWorkBookRef;
            xlAppRef = new Excel.ApplicationClass();
            xlAppRef.Visible = false; ; xlAppRef.DisplayAlerts = false; xlAppRef.ScreenUpdating = false;
            xlAppRef.DisplayAlerts = false;
            xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            Excel.Worksheet xlWorkSheetRef = (Excel.Worksheet)xlWorkBookRef.Worksheets.get_Item("Historique");
            Excel.Range rangeRefall = xlWorkSheetRef.get_Range("A1", "W" + getRowsLimit(xlWorkSheetRef));
            object[,] valuess = (object[,])rangeRefall.Value2;
            //bug : le seul moyen pour supprimer la dernière colonne est de chnager la largeur de toutes les colonnes (on ne sait pas pourquoi) !!!

            // xlWorkSheetRef.Cells.ColumnWidth = 20;
            int rowcount = rangeRefall.Rows.Count;
            int colcount = rangeRefall.Columns.Count;

            Excel.Range rangeRef = xlWorkSheetRef.get_Range("A" + getRowsLimit(xlWorkSheetRef), "W" + getRowsLimit(xlWorkSheetRef));
            rangeRef.EntireRow.Copy(misValue);

            rangeRef.EntireRow.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
            Excel.Range rangeRefdel = xlWorkSheetRef.UsedRange.get_Range("X1", xlWorkSheetRef.Cells[1, xlWorkSheetRef.UsedRange.Columns.Count - 1]) as Excel.Range;
            rangeRefdel.EntireColumn.ClearContents();
            rangeRefdel.EntireColumn.ClearFormats();
            rangeRefdel.EntireColumn.Clear();
            try
            {
                rangeRefdel = xlWorkSheetRef.UsedRange.get_Range("A" + (1 + getRowsLimit(xlWorkSheetRef)), "A" + xlWorkSheetRef.UsedRange.Rows.Count) as Excel.Range;
            }
            catch (Exception rx)
            {
                textBox1.AppendText(rx.ToString());
            }
            rangeRefdel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

            rangeRefdel = xlWorkSheetRef.UsedRange.get_Range("A1", "W" + (getRowsLimit(xlWorkSheetRef) - 1)) as Excel.Range;
            rangeRefdel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            Excel.Range rangeA1 = xlWorkSheetRef.Cells[1, 1] as Excel.Range;
            rangeA1.Activate();

            xlWorkSheetRef.SaveAs(@"D:\ptw\Histo.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBookRef.Close(true, misValue, misValue);
            xlAppRef.Quit();
            //////////////////////////////////////////////////////////////////////////////////
            Thread.Sleep(3000);
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = false; xlApp.DisplayAlerts = false; xlApp.ScreenUpdating = false;
            xlApp.DisplayAlerts = false;
            xlApp.Application.DisplayAlerts = false;

            //MessageBox.Show(openfilex);//D:\ptw\Histo.xls
            string remplacehisto8 = "[" + openfilex.Substring(7, 9) + "]";
            //xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.get_Range("A1", "AB2062");
            object[,] values = (object[,])range.Value2;



            int rCnt = 0;
            int rowx = xlWorkSheet.get_Range("A1", "AB2062").Rows.Count;
            int cCnt = 0;
            int col = 0;
            int col3000 = 0;
            int col4000 = 0;
            int col5000 = 0;
            int col8000 = 0;
            int col83000 = 0;
            rCnt = xlWorkSheet.get_Range("A1", "AB2062").Rows.Count;


            for (cCnt = 1; cCnt <= xlWorkSheet.get_Range("A1", "AB2062").Columns.Count - 1; cCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "3000"))
                {
                    col3000 = cCnt;
                }
                if (Regex.Equals(valuecellabs, "4000"))
                {
                    col4000 = cCnt;
                }
                if (Regex.Equals(valuecellabs, "5000"))
                {
                    col5000 = cCnt;
                }
                if (Regex.Equals(valuecellabs, "8000"))
                {
                    col8000 = cCnt;
                }
                if (Regex.Equals(valuecellabs, "10000"))
                {
                    col = cCnt;
                }
                if (Regex.Equals(valuecellabs, "83000"))
                {
                    col83000 = cCnt;
                    break;
                }
            }
            int fileflag = 0;
            for (int row = 25; row <= 2061; row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "-1"))
                {
                    textBox1.AppendText("Row number:" + row);
                    Thread.Sleep(3000);
                    xlWorkBookx1 = xlApp.Workbooks.Open(@"D:\ptw\Histo.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    // xlWorkBookx1 = xlApp.Workbooks.Open( @textBox12.Text + "\\Histox.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    Excel.Worksheet xlWorkSheetx1 = (Excel.Worksheet)xlWorkBookx1.Worksheets.get_Item("Historique");
                    string[] namestable = { "S-ACT.xlsx", "S-PAS.xlsx", "S-CR.xlsx", "S-ANN3.xlsx", "S-ANN4.xlsx", "S-ANN5.xlsx" };

                    string divisavenom = pathnotapme + "\\" + namestable[fileflag];
                    System.IO.Directory.CreateDirectory(pathnotapme);//////////////cree repertoire

                    xlWorkSheetx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    xlWorkBookx1.Close(true, misValue, misValue);
                    ////////////Grande titre "-1"/////////////////////////////////////////////////////////////////
                    if (Regex.Equals(Convert.ToString(values[25, col]), "-1"))
                    {
                        Excel.Range rangegtitre = xlWorkSheet.Cells[25, col] as Excel.Range;
                        Excel.Range rangePastegtitre = xlWorkSheet.UsedRange.Cells[24, 1] as Excel.Range;
                        rangegtitre.EntireRow.Cut(rangePastegtitre.EntireRow);

                        Excel.Range rangegtitreblank = xlWorkSheet.Cells[25, col] as Excel.Range;
                        rangegtitreblank.EntireRow.Delete(misValue);
                        row--;// point important, pour garder l'ordre de ligne ne change pas
                    }

                    ////////////////////insertion///////////////////////////////////////////////////////////////////
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row, col] as Excel.Range;
                    Excel.Range rangediviser = xlWorkSheet.UsedRange.get_Range("A1", xlWorkSheet.Cells[row - 1, col - 1]) as Excel.Range;
                    Excel.Range rangedelete = xlWorkSheet.UsedRange.get_Range("A25", xlWorkSheet.Cells[row - 1, col]) as Excel.Range;
                    rangediviser.EntireRow.Select();
                    rangediviser.EntireRow.Copy(misValue);
                    //MessageBox.Show(row.ToString());

                    xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    xlApp.DisplayAlerts = false;
                    xlApp.Application.DisplayAlerts = false;


                    Excel.Worksheet xlWorkSheetnewx1 = (Excel.Worksheet)xlWorkBooknewx1.Worksheets.get_Item("Historique");
                    /*Jintao: clearcontents for the worksheet
                    Excel.Range erals = xlWorkSheetnewx1.UsedRange;
                    erals.ClearContents();*/
                    //xlWorkBooknewx1.set_Colors(misValue, xlWorkBook.get_Colors(misValue));
                    Excel.Range rangenewx1 = xlWorkSheetnewx1.Cells[1, 1] as Excel.Range;
                    rangenewx1.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, misValue);

                    xlWorkSheetnewx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    //modifier lien pour effacer cross file reference!!!!!!!!!!!!!!2003-2010
                    xlWorkBooknewx1.ChangeLink(openfilex, divisavenom);
                    xlWorkBooknewx1.Close(true, misValue, misValue);

                    ////////////////////replace formulaire contient ptw/histo8.xls///////////////////
                    Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    xlApp.DisplayAlerts = false;
                    xlApp.Application.DisplayAlerts = false;



                    Excel.Worksheet xlWorkSheetremplace = (Excel.Worksheet)xlWorkBookremplace.Worksheets.get_Item("Historique");
                    Excel.Range rangeremplace = xlWorkSheetremplace.get_Range("A1", "W" + getRowsLimit(xlWorkSheetremplace));
                    rangeremplace.Cells.Replace(remplacehisto8, "", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);//NB remplacehisto8 il faut ameliorer pour adapder tous les cas
                    ////////delete col8000 "-2"//////////////////////////////////////////////////
                    object[,] values8000 = (object[,])rangeremplace.Value2;

                    for (int rowdel = 1; rowdel <= rangeremplace.Rows.Count; rowdel++)
                    {
                        string valuedel = Convert.ToString(values8000[rowdel, col8000]);
                        if (Regex.Equals(valuedel, "-2"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowdel, col8000] as Excel.Range;
                            rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                            rangeremplace = xlWorkSheetremplace.UsedRange;
                            values8000 = (object[,])rangeremplace.Value2;
                            rowdel--;
                        }
                    }
                    ///////////////row hide "-5"////////////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                        if (Regex.Equals(valuedel, "-5"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                            rangeDely.EntireRow.Hidden = true;
                        }
                    }
                    ///////////////row supprimer "-6"////////////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                        if (Regex.Equals(valuedel, "-6"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                            rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                            rangeremplace = xlWorkSheetremplace.UsedRange;
                            values8000 = (object[,])rangeremplace.Value2;
                            rowhide--;
                        }
                    }
                    ///////////////Hide -1 pour col 83000/////////////////////////////////////////////
                    //for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    //{
                    //    string valuedel = Convert.ToString(values8000[rowhide, col83000]);
                    //    if (Regex.Equals(valuedel, "-1"))
                    //    {
                    //        Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col83000] as Excel.Range;
                    //        rangeDely.EntireRow.Hidden = true;
                    //    }
                    //}
                    /////////////////////////////////////////////////////////////////////////////////
                    object[,] valuesNX = (object[,])rangeremplace.Value2;
                    //string valueNX = Convert.ToString(valuesNX[row, col]);
                    for (int row3000 = 1; row3000 <= rangeremplace.Rows.Count; row3000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-3"))
                        {
                            rangeprey.Locked = false;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row3000, col8000]) != "-7")//-7 non zero
                        {
                            rangeDely.Value2 = 0;
                        }
                    }
                    for (int row4000 = 1; row4000 <= rangeremplace.Rows.Count; row4000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-3"))
                        {
                            rangeprey.Locked = false;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row4000, col8000]) != "-7")//-7 non zero
                        {
                            rangeDely.Value2 = 0;
                        }
                    }
                    for (int row5000 = 1; row5000 <= rangeremplace.Rows.Count; row5000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-3"))
                        {
                            rangeprey.Locked = false;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row5000, col8000]) != "-7")//-7 non zero
                        {
                            rangeDely.Value2 = 0;
                        }
                    }

                    int countr = xlWorkSheetremplace.UsedRange.Rows.Count;
                    int countc = xlWorkSheetremplace.UsedRange.Columns.Count;
                    object[,] valuesx = (object[,])xlWorkSheetremplace.UsedRange.Value2;


                    Excel.Range rangeDeletex = xlWorkSheetremplace.UsedRange.get_Range("N1", xlWorkSheetremplace.Cells[1, xlWorkSheetremplace.UsedRange.Columns.Count]) as Excel.Range;
                    Excel.Range rangeDelete2 = xlWorkSheetremplace.get_Range(xlWorkSheetremplace.Cells[xlWorkSheetremplace.UsedRange.Rows.Count, 1], xlWorkSheetremplace.Cells[xlWorkSheetremplace.UsedRange.Rows.Count, 1]);

                    rangeDeletex.EntireColumn.Hidden = true;
                    rangeDelete2.EntireRow.Hidden = true;

                    ////////////////////////////////////////////////////////////////////////////
                    xlApp.ActiveWindow.SplitRow = 0;
                    xlApp.ActiveWindow.SplitColumn = 0;
                    xlWorkBookremplace.Save();
                    xlWorkBookremplace.Close(true, misValue, misValue);
                    




                    rangedelete.Copy(misValue);
                    rangedelete.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);



                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    row = 25;//important remise le ligne commencer apres action delete 1:)25ligne
                    xlWorkSheet.Activate();
                    fileflag++;
                }
            }
            xlApp.Quit();

            MessageBox.Show("jobs done");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (File.Exists("d:\\Jintao\\test.tdfc"))
            {
                try
                {
                    File.Move("d:\\Jintao\\test.tdfc", "d:\\Jintao\\test.txt");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    MessageBox.Show("OK");
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            double tempv = 0;
            string sssda =textBox2.Text.ToString();
            if (sssda.Replace("%", "").Contains(","))
            {
                string[] x = sssda.Replace("%", "").Split(',');
                tempv = (double.Parse(x[0]) + double.Parse(x[1]) / (10 * Convert.ToInt32(x[1]).ToString().Length)) / 100;
            }
            else
            {
                tempv = double.Parse(sssda.Replace("%", "")) / 100;
            }
            textBox3.Text = tempv.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string value = textBox2.Text;
            int valuelength = value.Length;
            for (int i = 3; i < valuelength; i += 3)
            {
                value = value.Substring(0, value.Length - i - i / 3 + 1) + " " + value.Substring(value.Length - i - i / 3 + 1);
            }
            textBox3.Text = value;
        }
    }
}
