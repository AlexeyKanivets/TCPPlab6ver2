using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Correlation_analisis
{
    public partial class Form1 : Form
    {
        int z = 0;
        private double R;
        private double x;
        private double y;
        string str;
        int rCnt;
        int cCnt;
        string filename;
        int page = 3;
        double[] averageArray = new double[7];
        double[] Fisher = new double[6];
        double[] CorrDet = new double[6];
        double[] FiVi = new double[6];
        double[] tmpV = new double[6];
        double[] tmpG = new double[12];
        double[,] CorrM = new double[6, 6];
        double[,] Matr = new double[6, 12];

        List<double> Styd = new List<double>();


        public Form1()
        {
            InitializeComponent();
        }

        private double correlator1(int firstColumn, int secondColumn)
        {
            double r1 = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                z++;
                r1 += (Convert.ToDouble(dataGridView1.Rows[i].Cells[firstColumn].Value) - averageArray[firstColumn]) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[secondColumn].Value) - averageArray[secondColumn]);
            }
            return r1;
        }

        private double correlator2(int firstColumn)
        {
            double r3 = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                r3 += (Convert.ToDouble(dataGridView1.Rows[i].Cells[firstColumn].Value) - averageArray[firstColumn]) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[firstColumn].Value) - averageArray[firstColumn]);
                z++;
            }
            return r3;
        }

        private void placeholredSecondGrid()
        {
            double[] corel2 = new double[7];
            for (int i = 1; i < 7; i++)
            {
                corel2[i] = correlator2(i);
            }
            for (int i = 1; i < 7; i++)
            {
                for (int j = 1; j < 7; j++)
                {
                    CorrM[i - 1, j - 1] = (correlator1(i, j) / Math.Sqrt(corel2[i] * corel2[j]));
                    dataGridView2.Rows[i - 1].Cells[j].Value = CorrM[i - 1, j - 1].ToString();                   
                }
            }
            //  MessageBox.Show(z.ToString());
        }

        private double getAverageColumn(int numberOfColumn)
        {
            double Sum = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                Sum += Convert.ToDouble(dataGridView1.Rows[i].Cells[numberOfColumn].Value);
            }
            Sum = Sum / (dataGridView1.RowCount - 1);
            return Sum;
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {


            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Excel (*.XLS; *.XLSX)|*.XLS; *.XLSX ";
            opf.ShowDialog();
            System.Data.DataTable tb = new System.Data.DataTable();
            filename = opf.FileName;



            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExcelRange;

            ExcelWorkBook = ExcelApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false,
                false, 0, true, 1, 0);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(2);

            ExcelRange = ExcelWorkSheet.UsedRange;
            dataGridView1.Rows.Clear();
            firstCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 1] as Microsoft.Office.Interop.Excel.Range).Text;
            secondCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 2] as Microsoft.Office.Interop.Excel.Range).Text;
            thirdCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 3] as Microsoft.Office.Interop.Excel.Range).Text;
            fourthCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 4] as Microsoft.Office.Interop.Excel.Range).Text;
            fifthCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 5] as Microsoft.Office.Interop.Excel.Range).Text;
            sixthCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 6] as Microsoft.Office.Interop.Excel.Range).Text;
            seventhCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 7] as Microsoft.Office.Interop.Excel.Range).Text;

            secondCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 2] as Microsoft.Office.Interop.Excel.Range).Text;
            thirdCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 3] as Microsoft.Office.Interop.Excel.Range).Text;
            fourthCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 4] as Microsoft.Office.Interop.Excel.Range).Text;
            fifthCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 5] as Microsoft.Office.Interop.Excel.Range).Text;
            sixthCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 6] as Microsoft.Office.Interop.Excel.Range).Text;
            seventhCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 7] as Microsoft.Office.Interop.Excel.Range).Text;

            dataGridView2.Rows.Add(7);

            dataGridView2.Rows[0].Cells[0].Value = secondCellTable2.HeaderText;
            dataGridView2.Rows[1].Cells[0].Value = thirdCellTable2.HeaderText;
            dataGridView2.Rows[2].Cells[0].Value = fourthCellTable2.HeaderText;
            dataGridView2.Rows[3].Cells[0].Value = fifthCellTable2.HeaderText;
            dataGridView2.Rows[4].Cells[0].Value = sixthCellTable2.HeaderText;
            dataGridView2.Rows[5].Cells[0].Value = seventhCellTable2.HeaderText;
            for (rCnt = 1; rCnt <= ExcelRange.Rows.Count - 3; rCnt++)
            {
                dataGridView1.Rows.Add(1);
                for (cCnt = 1; cCnt <= 7; cCnt++)
                {

                    str = (string)(ExcelRange.Cells[rCnt + 3, cCnt] as Microsoft.Office.Interop.Excel.Range).Text;
                    dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                }
            }
            ExcelWorkBook.Close(true, null, null);
            ExcelApp.Quit();

            for (int i = 1; i < 7; i++)
            {
                averageArray[i] = getAverageColumn(i);
            }

            placeholredSecondGrid();
            gridColor();

        }

        private void correlPage(object sender, EventArgs e, int page)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExcelRange;

            ExcelWorkBook = ExcelApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false,
                false, 0, true, 1, 0);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(page);

            ExcelRange = ExcelWorkSheet.UsedRange;
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();

            firstCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 1] as Microsoft.Office.Interop.Excel.Range).Text;
            secondCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 2] as Microsoft.Office.Interop.Excel.Range).Text;
            thirdCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 3] as Microsoft.Office.Interop.Excel.Range).Text;
            fourthCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 4] as Microsoft.Office.Interop.Excel.Range).Text;
            fifthCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 5] as Microsoft.Office.Interop.Excel.Range).Text;
            sixthCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 6] as Microsoft.Office.Interop.Excel.Range).Text;
            seventhCellTable1.HeaderText = (string)(ExcelRange.Cells[3, 7] as Microsoft.Office.Interop.Excel.Range).Text;

            secondCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 2] as Microsoft.Office.Interop.Excel.Range).Text;
            thirdCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 3] as Microsoft.Office.Interop.Excel.Range).Text;
            fourthCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 4] as Microsoft.Office.Interop.Excel.Range).Text;
            fifthCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 5] as Microsoft.Office.Interop.Excel.Range).Text;
            sixthCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 6] as Microsoft.Office.Interop.Excel.Range).Text;
            seventhCellTable2.HeaderText = (string)(ExcelRange.Cells[3, 7] as Microsoft.Office.Interop.Excel.Range).Text;

            dataGridView2.Rows.Add(7);

            dataGridView2.Rows[0].Cells[0].Value = secondCellTable2.HeaderText;
            dataGridView2.Rows[1].Cells[0].Value = thirdCellTable2.HeaderText;
            dataGridView2.Rows[2].Cells[0].Value = fourthCellTable2.HeaderText;
            dataGridView2.Rows[3].Cells[0].Value = fifthCellTable2.HeaderText;
            dataGridView2.Rows[4].Cells[0].Value = sixthCellTable2.HeaderText;
            dataGridView2.Rows[5].Cells[0].Value = seventhCellTable2.HeaderText;

            for (rCnt = 1; rCnt <= ExcelRange.Rows.Count - 3; rCnt++)
            {
                dataGridView1.Rows.Add(1);
                for (cCnt = 1; cCnt <= 7; cCnt++)
                {

                    str = (string)(ExcelRange.Cells[rCnt + 3, cCnt] as Microsoft.Office.Interop.Excel.Range).Text;
                    dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                }
            }

            for (int i = 1; i < 7; i++)
            {
                averageArray[i] = getAverageColumn(i);
            }

            ExcelWorkBook.Close(true, null, null);
            ExcelApp.Quit();

            placeholredSecondGrid();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            page = 2;

            correlPage(sender, e, page);
            gridColor();
            
        }
        private void but2_Click(object sender, EventArgs e)
        {

            page = 3;
            correlPage(sender, e, page);
            gridColor();
        }
        private void button3_Click(object sender, EventArgs e)
        {

            page = 4;
            correlPage(sender, e, page);
            gridColor();
        }

        private void gridColor()
        {
            for (int i = 0; i < 6; i++)
            {
                for (int j = 1; j < 7; j++)
                {
                    if (Convert.ToDouble(dataGridView2.Rows[i].Cells[j].Value) < 1.01)
                        dataGridView2.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.Red;

                    if (Convert.ToDouble(dataGridView2.Rows[i].Cells[j].Value) < 0.75)
                        dataGridView2.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.Pink;

                    if (Convert.ToDouble(dataGridView2.Rows[i].Cells[j].Value) < 0.51)
                        dataGridView2.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.Yellow;

                    if (Convert.ToDouble(dataGridView2.Rows[i].Cells[j].Value) < 0.3)
                        dataGridView2.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.Green;
                }
            }
        }

        private void pohM()
        {
            for (int i = 0; i < Matr.GetLength(0); i++)
            {
                for (int j = 0; j < Matr.GetLength(1); j++)
                {
                    Matr[i, j] = 0;
                }
            }
            Matr[0, 6] = 1;
            Matr[1, 7] = 1;
            Matr[2, 8] = 1;
            Matr[3, 9] = 1;
            Matr[4, 10] = 1;
            Matr[5, 11] = 1;
            for (int i = 0; i < CorrM.GetLength(0); i++)
            {
                for (int j = 0; j < CorrM.GetLength(0); j++)
                {                    
                        Matr[i, j] = CorrM[i,j];
                }
            }

            for (int k = 0; k < Matr.GetLength(0); k++)
            {
                for (int i = 0; i < Matr.GetLength(0); i++)
                {
                    tmpV[i] = Matr[i, k];
                }
                for (int i = 0; i < Matr.GetLength(1); i++)
                {
                    tmpG[i] = Matr[k, i];
                }
                for (int i = 0; i < Matr.GetLength(0); i++)
                {
                    if (i==k)
                    {
                        for (int j = 0; j < Matr.GetLength(1); j++)
                        {
                            Matr[i, j] = Matr[i, j] / tmpV[k];
                            Console.Write("  {0}", Matr[i, j]);
                        }
                    }
                    else
                    {
                        for (int j = 0; j < Matr.GetLength(1); j++)
                        {
                            Matr[i, j] = Matr[i, j] - tmpV[i]*tmpG[j]/ tmpV[k];
                            Console.Write("  {0}", Matr[i, j]);
                        }
                    }
                    Console.WriteLine(" ");
                }
            }

            Console.WriteLine(" ");
            for (int i = 0; i < Matr.GetLength(0); i++)
            {
                for (int j = 0; j < Matr.GetLength(1); j++)
                {
                   // Console.Write("  {0}", Matr[i, j]);
                }
                Console.WriteLine(" ");
            }                                       
        }

        private double det(double[,] A)
        {
            double det = 1 ;
            double m;
            for (int k = 1; k < A.GetLength(0); k++)
            {
                for (int j = k; j < A.GetLength(0); j++)
                {
                    m = A[j,k - 1] / A[k - 1,k - 1];
                    for (int i = 0; i < A.GetLength(1); i++)
                    {
                        A[j,i] = A[j,i] - m * A[k - 1,i];
                    }
                }
            }
            Console.WriteLine(" ");
            for (int i = 0; i < A.GetLength(0); i++)
            {
                for (int j = 0; j < A.GetLength(1); j++)
                {
                    Console.Write("  {0}", A[i, j]);
                }
                Console.WriteLine(" ");
            }
            for (int i = 0; i < A.GetLength(0); i++)
            {
                det =det*A[i, i];
            }
            Console.WriteLine("det {0}",det);
            return det;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            pohM();
            double rkj;
            double xi;
            double styd= 2.04;
            double xit = 25;
            bool fish = false;

            xi = (-22.167) * Math.Log(det(CorrM));            
            if (xi > xit)
            {
                MessageBox.Show("Мультиколінеарність існує по хі2");
            }
            else
            {
                MessageBox.Show("Мультиколінеарністі немає");
                return;
            }
            Console.Write("Fisher: ");
            for (int i = 0; i < CorrDet.Length; i++)
            {
                CorrDet[i] = 1 - 1 / Matr[i, i + 6];
                Fisher[i] = ((CorrDet[i]) / 5) / ((1 - CorrDet[i]) / 20);
                if (Fisher[i] > 2.7)
                {
                    fish = true;
                }
                Console.Write("  {0}", Fisher[i]);
            }

            if (fish)
            {
                MessageBox.Show("Мультиколінеарність існує по fisher");
            }
            else
            {
                MessageBox.Show("Мультиколінеарністі немає");
                return;
            }

            dataGridView2.Rows.Clear();
            dataGridView2.Rows.Add(7);
            dataGridView2.Rows[0].Cells[0].Value = secondCellTable2.HeaderText;
            dataGridView2.Rows[1].Cells[0].Value = thirdCellTable2.HeaderText;
            dataGridView2.Rows[2].Cells[0].Value = fourthCellTable2.HeaderText;
            dataGridView2.Rows[3].Cells[0].Value = fifthCellTable2.HeaderText;
            dataGridView2.Rows[4].Cells[0].Value = sixthCellTable2.HeaderText;
            dataGridView2.Rows[5].Cells[0].Value = seventhCellTable2.HeaderText;


            for (int i = 0; i < Matr.GetLength(0); i++)
            {
                for (int j = 0; j < Matr.GetLength(0); j++)
                {
                    rkj = -Matr[i, j + 6] / Math.Sqrt(Matr[j, j + 6] * Matr[i, i + 6]);
                    Styd.Add(Math.Abs(rkj) * Math.Sqrt(26 - 6) / Math.Sqrt(1 - rkj * rkj));
                    if (Math.Abs(rkj) * Math.Sqrt(26 - 6) / Math.Sqrt(1 - rkj * rkj) > styd)
                    {
                        dataGridView2.Rows[i].Cells[j + 1].Value = Math.Abs(rkj) * Math.Sqrt(26 - 6) / Math.Sqrt(1 - rkj * rkj);
                    }
                }
            }

        }
    }
}
    