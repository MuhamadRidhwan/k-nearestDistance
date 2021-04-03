using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;       // EXCEL APPLICATION.
using System.Drawing;
using System.Runtime.InteropServices;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Data.OleDb;
using System.Collections.ObjectModel;

namespace NearestDistanceSystem
{
    public partial class Form1 : Form
    {
        private string fileName = "";
        private string process = "";
        private string outFileName = "";

        // CREATE EXCEL OBJECTS.
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;

        List<string> lsCapacity = new List<string>();
        List<string> lsPopulation = new List<string>();
        List<string> lsDistance = new List<string>();

        //List<KeyValuePair<decimal, decimal>> lsTempBigList = new List<KeyValuePair<decimal, decimal>>();
        List<TempDistanceClass> lsTempBigList = new List<TempDistanceClass>();
        List<TempDistanceClass2> lsTempBigList2 = new List<TempDistanceClass2>();

        List<DistanceClass> DistanceList = new List<DistanceClass>();
        List<CapacityClass> CapacityList = new List<CapacityClass>();
        List<PopulationClass> PopulationList = new List<PopulationClass>();

        public Form1()
        {
            InitializeComponent();
        }

        private void BtnAddPopulationExcel_Click(object sender, EventArgs e)
        {
            if (dgvPopulation.Rows.Count != 0)
            {
                btnAddPopulationExcel.Enabled = false;
            }
            else
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";

                string sFileName;

                sFileName = openFileDialog.FileName;

                dgvPopulation.AllowUserToAddRows = false;

                //if (openFileDialog.ShowDialog() == DialogResult.OK)
                //{
                fileName = openFileDialog.FileName;
                txtPopulationExcel.Text = fileName;

                //if (fileName.Trim() != "")
                //{
                ExcelCallPopulation(fileName);

                foreach (DataGridViewRow dr in dgvPopulation.Rows)
                {
                    //Create object of your list type pl
                    PopulationClass pl = new PopulationClass();
                    pl.pProperty1 = Convert.ToDecimal(dr.Cells[0].Value);

                    //Add pl to your List  
                    PopulationList.Add(pl);
                }

                btnAddPopulationExcel.Enabled = false;

                if (dgvDistance.Rows.Count != 0 && dgvCapacity.Rows.Count != 0)
                {
                    btnReCalculate.Enabled = true;
                    btnCalculate.Enabled = true;
                }
            }

            #region comment
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.InitialDirectory = "c:\\";
            //openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";

            //if (openFileDialog.ShowDialog() == DialogResult.OK)
            //{
            //    fileName = openFileDialog.FileName;
            //    txtPopulationExcel.Text = fileName;
            #endregion

        }

        private void BtnAddCapacityExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";

            string sFileName;

            sFileName = openFileDialog.FileName;

            dgvCapacity.AllowUserToAddRows = false;

            //if (openFileDialog.ShowDialog() == DialogResult.OK)
            //{
            fileName = openFileDialog.FileName;
            txtCapacityExcel.Text = fileName;

            //if (fileName.Trim() != "")
            //{
            ExcelCallCapacity(fileName);

            foreach (DataGridViewRow dr in dgvCapacity.Rows)
            {
                //Create object of your list type pl
                CapacityClass pl = new CapacityClass();
                //pl.cProperty1 = Convert.ToDecimal(dr.Cells[0].Value);
                pl.cProperty1 = Convert.ToInt16(dr.Cells[0].Value);

                //Add pl to your List  
                CapacityList.Add(pl);
            }

            btnAddCapacityExcel.Enabled = false;

            if (dgvDistance.Rows.Count != 0 && dgvPopulation.Rows.Count != 0)
            {
                btnReCalculate.Enabled = true;
                btnCalculate.Enabled = true;
            }
        }

        private void BtnAddDistanceExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "All files (*.*)|*.*";

            string sFileName;

            sFileName = openFileDialog.FileName;

            dgvDistance.AllowUserToAddRows = false;

            //if (openFileDialog.ShowDialog() == DialogResult.OK)
            //{
            fileName = openFileDialog.FileName;
            txtDistanceExcel.Text = fileName;

            //if (fileName.Trim() != "")
            //{
            ExcelCallDistance(fileName);

            foreach (DataGridViewRow dr in dgvDistance.Rows)
            {
                //Create object of your list type pl
                DistanceClass pl = new DistanceClass();
                pl.dProperty1 = dr.Cells[0].Value;
                pl.dProperty2 = dr.Cells[1].Value;
                pl.dProperty3 = dr.Cells[2].Value;
                pl.dProperty4 = dr.Cells[3].Value;
                pl.dProperty5 = dr.Cells[4].Value;
                pl.dProperty6 = dr.Cells[5].Value;
                pl.dProperty7 = dr.Cells[6].Value;
                pl.dProperty8 = dr.Cells[7].Value;
                pl.dProperty9 = dr.Cells[8].Value;
                pl.dProperty10 = dr.Cells[9].Value;
                pl.dProperty11 = dr.Cells[10].Value;
                pl.dProperty12 = dr.Cells[11].Value;
                pl.dProperty13 = dr.Cells[12].Value;
                pl.dProperty14 = dr.Cells[13].Value;
                pl.dProperty15 = dr.Cells[14].Value;
                pl.dProperty16 = dr.Cells[15].Value;
                pl.dProperty17 = dr.Cells[16].Value;
                pl.dProperty18 = dr.Cells[17].Value;
                pl.dProperty19 = dr.Cells[18].Value;
                pl.dProperty20 = dr.Cells[19].Value;
                pl.dProperty21 = dr.Cells[20].Value;
                pl.dProperty22 = dr.Cells[21].Value;
                pl.dProperty23 = dr.Cells[22].Value;
                pl.dProperty24 = dr.Cells[23].Value;
                pl.dProperty25 = dr.Cells[24].Value;
                pl.dProperty26 = dr.Cells[25].Value;
                pl.dProperty27 = dr.Cells[26].Value;
                pl.dProperty28 = dr.Cells[27].Value;
                pl.dProperty29 = dr.Cells[28].Value;

                //Add pl to your List  
                DistanceList.Add(pl);
            }

            btnAddDistanceExcel.Enabled = false;

            if (dgvPopulation.Rows.Count != 0 && dgvCapacity.Rows.Count != 0)
            {
                btnReCalculate.Enabled = true;
                btnCalculate.Enabled = true;
            }
        }

        OpenFileDialog openFileDialog1 = new OpenFileDialog();

        Dictionary<decimal, decimal> myDict = new Dictionary<decimal, decimal>();

        public void Button1_Click(object sender, EventArgs e)
        {
            txtPopulationExcel.Text = string.Empty;

            #region comment

            //myDict.Clear();

            //var distance1 = new decimal[] { 2.95186101M,  4.379658932M, 5.416592903M, 5.854756264M, 5.972246991M, 6.331351431M, 7.120540195M, 8.336009981M, 8.868575713M, 9.077406235M, 
            //                                9.245042592M, 9.436866549M, 9.712565433M, 9.855485688M, 10.93336699M, 11.88228854M, 13.3352104M,  13.88906741M, 14.1835416M,  15.70016346M,
            //                                15.73392311M, 16.4166177M,  16.41661771M,  16.58940685M, 17.9950047M,  18.33234635M, 19.27633018M, 19.50598813M, 20.63478634M
            //                                };          

            //var population = new decimal[] { 2075, 800, 950, 1817, 1200, 361, 993, 1030, 522, 1402, 1051, 670, 1350, 850, 950, 1276, 1050, 1645, 1125, 1194, 870, 1870, 800,
            //                                1237, 400, 700, 200, 100, 500 }; // village --> for each population element represent
            //                                                                             //  each prone area(village affected).

            //var capacity = new decimal[] { 4050, 330, 580, 250, 530, 150, 40, 85, 70, 490, 1200, 200, 80, 500, 850, 650, 1800, 1800, 350, 900, 3970, 4770, 500, 60,
            //                                20, 145, 240, 500, 300 }; // release center assigned

            //------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

            //var a = dgvDistance.CurrentRow.Cells[0].Value.ToString();

            #endregion

            List<decimal> tempDistance = new List<decimal>();
            List<decimal> tempDistance2 = new List<decimal>();

            var items = new List<KeyValuePair<decimal, decimal>>();
            List<LookUpClass> li = new List<LookUpClass>();
            List<KeyValuePair<decimal, decimal>> listDict = new List<KeyValuePair<decimal, decimal>>();

            List<KeyValuePair<decimal, decimal>> listDict2 = new List<KeyValuePair<decimal, decimal>>();

            var MinimumValueDistance = (dynamic)null;
            var firstPopulation = (dynamic)null;

            decimal balPopulation = 0;
            decimal totalDistance = 0;
            decimal totalDistance2 = 0;
            decimal totalPopulation = 0;

            for (int i = 0; i < PopulationList.Count(); i++)
            {
                totalPopulation += PopulationList[i].pProperty1;
            }

            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            List<PopulationClass> list = new List<PopulationClass>(PopulationList);

            int currentCol = 0;
            var newEntry = (dynamic)null;
            int currentRow = 0;
            dynamic balCap;

            for (int j = 0; j < dgvDistance.Columns.Count; j++) // store data for first row in every column in a list
            {
                tempDistance.Add(Convert.ToDecimal(dgvDistance.Rows[currentRow].Cells[j].Value));
                //tempDistance.Add(Convert.ToDecimal(dgvDistance.CurrentRow.Cells[j].Value));
            }

            for (int i = 0; i < CapacityList.Count(); i++)
            {
                listDict.Add(new KeyValuePair<decimal, decimal>(tempDistance[i], CapacityList[i].cProperty1));
            }

            for (int k = 1; k <= CapacityList.Count(); k++) //Check for next row @ population
            {
                firstPopulation = list.First().pProperty1;

                for (int j = 0; j <= list.Count(); j++) //Check for next column @ capacity
                {
                    if (newEntry is null)
                    {
                        MinimumValueDistance = listDict.FirstOrDefault(x => x.Key == listDict.Min(y => y.Key));
                    }
                    else
                    {
                        MinimumValueDistance = newEntry;
                    }

                    balPopulation = firstPopulation - MinimumValueDistance.Value; //Population minus with Capacity

                    if (balPopulation == 0)
                    {
                        listDict.Remove(new KeyValuePair<decimal, decimal>(MinimumValueDistance.Key, MinimumValueDistance.Value)); //remove distance and capacity because population have been fully used

                        totalDistance = firstPopulation * MinimumValueDistance.Key;

                        txtPopulationExcel.Text += "Final balance for center " + k + ": " + 0 + "\r\n";
                    }

                    else if (balPopulation < 0) //Capacity is more than population. So, capacity still have the balance. Then, proceed with next population. Delete first row
                    {
                        listDict.Clear();
                        tempDistance.Clear();

                        currentRow++;

                        for (int m = 0; m < dgvDistance.Columns.Count; m++) // store data for first row in every column in a list
                        {
                            tempDistance.Add(Convert.ToDecimal(dgvDistance.Rows[currentRow].Cells[m].Value));
                            //tempDistance.Add(Convert.ToDecimal(dgvDistance.CurrentRow.Cells[m].Value));
                        }

                        for (int i = currentCol; i < CapacityList.Count(); i++)
                        {
                            listDict.Add(new KeyValuePair<decimal, decimal>(tempDistance[i], CapacityList[i].cProperty1));
                        }

                        balPopulation = balPopulation * -1; //if result turn out negative, then change it to positive

                        decimal Population = firstPopulation;

                        totalDistance = Population * MinimumValueDistance.Key; //calculate total distance travelled

                        txtCapacityExcel.Text += "Balance for center " + k + ": " + 0 + "\r\n";

                        //newEntry = new KeyValuePair<decimal, decimal>(MinimumValueDistance.Key, balPopulation);
                        newEntry = new KeyValuePair<decimal, decimal>(tempDistance[currentCol], balPopulation); // either remove the previous column that have been fully used or set from which column that it need to start with 

                        break;
                    }
                    else //balPopulation is more than 0, it means population is more than capacity. Then, continue proceed with same population.
                    {
                        currentCol++;

                        listDict.Remove(listDict.FirstOrDefault(x => x.Key == listDict.Min(y => y.Key))); //remove distance and capacity because capacity have been fully used

                        balCap = firstPopulation - balPopulation;

                        firstPopulation = balPopulation;//update current population                   

                        totalDistance = balCap * MinimumValueDistance.Key;

                        totalDistance2 = totalDistance2 + totalDistance;

                        txtPopulationExcel.Text += "Final balance for center " + k + ": " + balPopulation + "\r\n";

                        tempDistance.Clear();

                        for (int a = 0; a < dgvDistance.Columns.Count; a++) // store data for first row in every column in a list
                        {
                            tempDistance.Add(Convert.ToDecimal(dgvDistance.Rows[currentRow].Cells[a].Value));
                        }

                        if (currentCol < tempDistance.Count())
                        {
                            newEntry = new KeyValuePair<decimal, decimal>(tempDistance[currentCol], listDict[0].Value);
                        }

                    }

                    lsTempBigList.Add(new TempDistanceClass { tProperty1 = tempDistance[j] });

                } //end of second loop

                totalDistance2 = totalDistance2 + totalDistance;

                list.RemoveAt(index: 0); // remove first element of population, then replace it with next of element.

                tempDistance.Clear();

                for (int j = 0; j < dgvDistance.Columns.Count; j++) // store data for first row in every column in a list
                {
                    tempDistance.Add(Convert.ToDecimal(dgvDistance.Rows[currentRow].Cells[j].Value));
                }

                listDict.Clear();

                for (int i = currentCol; i < CapacityList.Count(); i++)
                {
                    listDict.Add(new KeyValuePair<decimal, decimal>(tempDistance[i], CapacityList[i].cProperty1));
                }

                bool isEmpty = !list.Any();

                if (isEmpty)
                {
                    MessageBox.Show("List ended");

                    break;
                }
                else
                {

                }

            } //end of first loop

            //MessageBox.Show("Total: " + String.Format("{0:0.00}", totalDistance2));

            txtDistanceExcel.Text = (totalDistance2 / totalPopulation).ToString();
        }

        public class DistanceClass
        {
            public object dProperty1 { get; set; }
            public object dProperty2 { get; set; }
            public object dProperty3 { get; set; }
            public object dProperty4 { get; set; }
            public object dProperty5 { get; set; }
            public object dProperty6 { get; set; }
            public object dProperty7 { get; set; }
            public object dProperty8 { get; set; }
            public object dProperty9 { get; set; }
            public object dProperty10 { get; set; }
            public object dProperty11 { get; set; }
            public object dProperty12 { get; set; }
            public object dProperty13 { get; set; }
            public object dProperty14 { get; set; }
            public object dProperty15 { get; set; }
            public object dProperty16 { get; set; }
            public object dProperty17 { get; set; }
            public object dProperty18 { get; set; }
            public object dProperty19 { get; set; }
            public object dProperty20 { get; set; }
            public object dProperty21 { get; set; }
            public object dProperty22 { get; set; }
            public object dProperty23 { get; set; }
            public object dProperty24 { get; set; }
            public object dProperty25 { get; set; }
            public object dProperty26 { get; set; }
            public object dProperty27 { get; set; }
            public object dProperty28 { get; set; }
            public object dProperty29 { get; set; }
        }

        public class CapacityClass
        {
            public decimal cProperty1 { get; set; }
        }

        public class PopulationClass
        {
            public decimal pProperty1 { get; set; }
        }

        public class TempDistanceClass
        {
            public decimal tProperty1 { get; set; }
        }

        public class TempDistanceClass2
        {
            public object tProperty1 { get; set; }
        }

        public class LookUpClass
        {
            public decimal lProperty1 { get; set; }
            public decimal lProperty2 { get; set; }
        }

        private void ExcelCallDistance(string sFile)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(System.AppDomain.CurrentDomain.BaseDirectory + "combinesortdist.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);  // WORKBOOK TO OPEN THE EXCEL FILE.
            //xlWorkBook = xlApp.Workbooks.Open(sFile);               // WORKBOOK TO OPEN THE EXCEL FILE.
            xlWorkSheet = xlWorkBook.Worksheets["Sheet1"];          // THE SHEET WITH THE DATA.

            dgvDistance.Rows.Clear();
            dgvDistance.Columns.Clear();

            int iRow, iCol;

            // FIRST, CREATE THE DataGridView COLUMN HEADERS.
            for (iCol = 2; iCol <= xlWorkSheet.Columns.Count; iCol++)
            {
                if (xlWorkSheet.Cells[1, iCol].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                    col.HeaderText = xlWorkSheet.Cells[1, iCol].value;
                    int colIndex = dgvDistance.Columns.Add(col);        // ADD A NEW COLUMN.
                }
            }

            // ADD ROWS TO THE GRID USING EXCEL DATA.
            for (iRow = 2; iCol <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    // CREATE A STRING ARRAY USING THE VALUES IN EACH ROW OF THE SHEET.
                    string[] row = new string[] {
                        xlWorkSheet.Cells[iRow, 2].value.ToString(),
                        xlWorkSheet.Cells[iRow, 3].value.ToString(),
                        xlWorkSheet.Cells[iRow, 4].value.ToString(),
                        xlWorkSheet.Cells[iRow, 5].value.ToString(),
                        xlWorkSheet.Cells[iRow, 6].value.ToString(),
                        xlWorkSheet.Cells[iRow, 7].value.ToString(),
                        xlWorkSheet.Cells[iRow, 8].value.ToString(),
                        xlWorkSheet.Cells[iRow, 9].value.ToString(),
                        xlWorkSheet.Cells[iRow, 10].value.ToString(),
                        xlWorkSheet.Cells[iRow, 11].value.ToString(),
                        xlWorkSheet.Cells[iRow, 12].value.ToString(),
                        xlWorkSheet.Cells[iRow, 13].value.ToString(),
                        xlWorkSheet.Cells[iRow, 14].value.ToString(),
                        xlWorkSheet.Cells[iRow, 15].value.ToString(),
                        xlWorkSheet.Cells[iRow, 16].value.ToString(),
                        xlWorkSheet.Cells[iRow, 17].value.ToString(),
                        xlWorkSheet.Cells[iRow, 18].value.ToString(),
                        xlWorkSheet.Cells[iRow, 19].value.ToString(),
                        xlWorkSheet.Cells[iRow, 20].value.ToString(),
                        xlWorkSheet.Cells[iRow, 21].value.ToString(),
                        xlWorkSheet.Cells[iRow, 22].value.ToString(),
                        xlWorkSheet.Cells[iRow, 23].value.ToString(),
                        xlWorkSheet.Cells[iRow, 24].value.ToString(),
                        xlWorkSheet.Cells[iRow, 25].value.ToString(),
                        xlWorkSheet.Cells[iRow, 26].value.ToString(),
                        xlWorkSheet.Cells[iRow, 27].value.ToString(),
                        xlWorkSheet.Cells[iRow, 28].value.ToString(),
                        xlWorkSheet.Cells[iRow, 29].value.ToString(),
                        xlWorkSheet.Cells[iRow, 30].value.ToString() };

                    // ADD A NEW ROW TO THE GRID USING THE ARRAY DATA.
                    dgvDistance.Rows.Add(row);
                }
            }

            xlWorkBook.Close();
            xlApp.Quit();

            // CLEAN UP.
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkSheet);
        }

        private void ExcelCallCapacity(string sFile)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(System.AppDomain.CurrentDomain.BaseDirectory + "capacity.csv", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);  // WORKBOOK TO OPEN THE EXCEL FILE.
            //xlWorkBook = xlApp.Workbooks.Open(sFile);               // WORKBOOK TO OPEN THE EXCEL FILE.
            xlWorkSheet = xlWorkBook.Worksheets["capacity"];          // THE SHEET WITH THE DATA.

            dgvCapacity.Rows.Clear();
            dgvCapacity.Columns.Clear();

            int iRow, iCol;

            // FIRST, CREATE THE DataGridView COLUMN HEADERS.
            for (iCol = 1; iCol <= xlWorkSheet.Columns.Count; iCol++)
            {
                if (xlWorkSheet.Cells[1, iCol].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                    //col.HeaderText = xlWorkSheet.Cells[1, iCol].value.ToString();
                    col.HeaderText = "Capacity";
                    int colIndex = dgvCapacity.Columns.Add(col);        // ADD A NEW COLUMN.                   
                }
            }

            // ADD ROWS TO THE GRID USING EXCEL DATA.
            for (iRow = 1; iCol <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    // CREATE A STRING ARRAY USING THE VALUES IN EACH ROW OF THE SHEET.
                    string[] row = new string[] { xlWorkSheet.Cells[iRow, 1].value.ToString() };

                    // ADD A NEW ROW TO THE GRID USING THE ARRAY DATA.
                    dgvCapacity.Rows.Add(row);
                }
            }

            xlWorkBook.Close();
            xlApp.Quit();

            // CLEAN UP.
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkSheet);
        }

        private void ExcelCallPopulation(string sFile)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(System.AppDomain.CurrentDomain.BaseDirectory + "population.csv", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);  // WORKBOOK TO OPEN THE EXCEL FILE.
            //xlWorkBook = xlApp.Workbooks.Open(sFile);               // WORKBOOK TO OPEN THE EXCEL FILE.
            xlApp = new Excel.Application();

            xlWorkSheet = xlWorkBook.Worksheets["population"];          // THE SHEET WITH THE DATA.

            dgvPopulation.Rows.Clear();
            dgvPopulation.Columns.Clear();

            int iRow, iCol;

            // FIRST, CREATE THE DataGridView COLUMN HEADERS.
            for (iCol = 1; iCol <= xlWorkSheet.Columns.Count; iCol++)
            {
                if (xlWorkSheet.Cells[1, iCol].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                    //col.HeaderText = xlWorkSheet.Cells[1, iCol].value.ToString();
                    col.HeaderText = "Population";
                    int colIndex = dgvPopulation.Columns.Add(col);        // ADD A NEW COLUMN.                   
                }
            }

            // ADD ROWS TO THE GRID USING EXCEL DATA.
            for (iRow = 1; iCol <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    // CREATE A STRING ARRAY USING THE VALUES IN EACH ROW OF THE SHEET.
                    string[] row = new string[] { xlWorkSheet.Cells[iRow, 1].value.ToString() };

                    // ADD A NEW ROW TO THE GRID USING THE ARRAY DATA.
                    dgvPopulation.Rows.Add(row);
                }
            }

            xlWorkBook.Close();
            xlApp.Quit();

            // CLEAN UP.
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkSheet);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnReCalculate.Enabled = false;
            btnCalculate.Enabled = false;
        }

        private void dgvDistance_SelectionChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgvDistance.SelectedRows)
            {
                string value1 = row.Cells[0].Value.ToString();

            }
        }

        public void btnReCalculate_Click(object sender, EventArgs e)
        {
            txtPopulationExcel.Text = string.Empty;
            txtCapacityExcel.Text = string.Empty;
            txtDistanceExcel.Text = string.Empty;
            txtLoop.Text = string.Empty;
            txtTotalUnallocated.Text = string.Empty;
            txtReleaseCenter.Text = string.Empty;
            int posMin = 0;

            #region Declare
            List<object> tempDistance = new List<object>();
            List<decimal> tempDistance2 = new List<decimal>();
            List<KeyValuePair<object, decimal>> listDict = new List<KeyValuePair<object, decimal>>();
            List<PopulationClass> list = new List<PopulationClass>(PopulationList);

            var MinimumValueDistance = (dynamic)null;
            var firstPopulation = (dynamic)null;
            var newEntry = (dynamic)null;
            dynamic balCap;
            decimal balPopulation = 0;
            decimal totalDistance = 0;
            decimal totalDistance2 = 0;
            decimal totalPopulation = 0;
            decimal totalAllocated = 0;
            decimal totalUnallocated = 0;
            int currentCol = 0;
            int currentRow = 0;
            List<decimal> capBalList = new List<decimal>();
            List<decimal> listDict2 = new List<decimal>();
            decimal tempCap = 0;
            #endregion

            string text = txtNoLoop.Text;

            for (int i = 0; i < CapacityList.Count; i++)
            {
                listDict2.Add(CapacityList[i].cProperty1);
            }

            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("Please fill in the numbers of loop that you want.");
            }

            else
            {
                int parsedValue;

                if (!int.TryParse(txtNoLoop.Text, out parsedValue))
                {
                    MessageBox.Show("Please insert only numbers.");
                }
                else
                {
                    for (int ai = 1; ai <= parsedValue; ai++)
                    {
                        List<DistanceClass> tempDistance3 = new List<DistanceClass>();
                        int currentRow2 = 0;

                        List<KeyValuePair<decimal, object>> dict = new List<KeyValuePair<decimal, object>>();

                        for (int i = 0; i < PopulationList.Count; i++)
                        {
                            tempDistance3.Add(new DistanceClass
                            {
                                dProperty1 = dgvDistance.Rows[currentRow2].Cells[0].Value,
                                dProperty2 = dgvDistance.Rows[currentRow2].Cells[1].Value,
                                dProperty3 = dgvDistance.Rows[currentRow2].Cells[2].Value,
                                dProperty4 = dgvDistance.Rows[currentRow2].Cells[3].Value,
                                dProperty5 = dgvDistance.Rows[currentRow2].Cells[4].Value,
                                dProperty6 = dgvDistance.Rows[currentRow2].Cells[5].Value,
                                dProperty7 = dgvDistance.Rows[currentRow2].Cells[6].Value,
                                dProperty8 = dgvDistance.Rows[currentRow2].Cells[7].Value,
                                dProperty9 = dgvDistance.Rows[currentRow2].Cells[8].Value,
                                dProperty10 = dgvDistance.Rows[currentRow2].Cells[9].Value,
                                dProperty11 = dgvDistance.Rows[currentRow2].Cells[10].Value,
                                dProperty12 = dgvDistance.Rows[currentRow2].Cells[11].Value,
                                dProperty13 = dgvDistance.Rows[currentRow2].Cells[12].Value,
                                dProperty14 = dgvDistance.Rows[currentRow2].Cells[13].Value,
                                dProperty15 = dgvDistance.Rows[currentRow2].Cells[14].Value,
                                dProperty16 = dgvDistance.Rows[currentRow2].Cells[15].Value,
                                dProperty17 = dgvDistance.Rows[currentRow2].Cells[16].Value,
                                dProperty18 = dgvDistance.Rows[currentRow2].Cells[17].Value,
                                dProperty19 = dgvDistance.Rows[currentRow2].Cells[18].Value,
                                dProperty20 = dgvDistance.Rows[currentRow2].Cells[19].Value,
                                dProperty21 = dgvDistance.Rows[currentRow2].Cells[20].Value,
                                dProperty22 = dgvDistance.Rows[currentRow2].Cells[21].Value,
                                dProperty23 = dgvDistance.Rows[currentRow2].Cells[22].Value,
                                dProperty24 = dgvDistance.Rows[currentRow2].Cells[23].Value,
                                dProperty25 = dgvDistance.Rows[currentRow2].Cells[24].Value,
                                dProperty26 = dgvDistance.Rows[currentRow2].Cells[25].Value,
                                dProperty27 = dgvDistance.Rows[currentRow2].Cells[26].Value,
                                dProperty28 = dgvDistance.Rows[currentRow2].Cells[27].Value,
                                dProperty29 = dgvDistance.Rows[currentRow2].Cells[28].Value

                            });

                            currentRow2++;
                        }

                        currentRow2 = 0;

                        for (int i = 0; i < PopulationList.Count; i++)
                        {
                            dict.Add(new KeyValuePair<decimal, object>(PopulationList[currentRow2].pProperty1, tempDistance3[i]));

                            currentRow2++;
                        }

                        //var shuffleList = dict.OrderBy(x => Guid.NewGuid()).ToList();
                        var shuffleList = dict;

                        #region                   
                        for (int i = 0; i < PopulationList.Count(); i++)
                        {
                            totalPopulation += PopulationList[i].pProperty1;
                        }

                        tempDistance.Add(shuffleList[0].Value);

                        #region looping distance

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty1));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty2));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty3));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty4));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty5));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty6));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty7));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty8));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty9));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty10));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty11));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty12));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty13));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty14));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty15));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty16));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty17));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty18));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty19));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty20));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty21));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty22));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty23));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty24));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty25));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty26));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty27));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty28));
                        }

                        foreach (DistanceClass a in tempDistance)
                        {
                            tempDistance2.Add(Convert.ToDecimal(a.dProperty29));
                        }

                        #endregion

                        for (int i = 0; i < CapacityList.Count(); i++)
                        {
                            listDict.Add(new KeyValuePair<object, decimal>(tempDistance2[i], CapacityList[i].cProperty1));
                        }

                        for (int k = 1; k <= PopulationList.Count(); k++) //Check for next row @ population
                        {
                            bool isEmpty = !shuffleList.Any();

                            if (isEmpty || currentCol == CapacityList.Count())
                            {
                                // do nothing. The purpose is to skip all the process
                            }
                            else
                            {
                                firstPopulation = shuffleList[0].Key;

                                for (int j = 0; j < CapacityList.Count(); j++) //Check for next column @ capacity
                                {
                                    if (newEntry is null)
                                    {
                                        MinimumValueDistance = listDict.FirstOrDefault(x => x.Key == listDict.Min(y => y.Key));                                       
                                    }
                                    else
                                    {
                                        MinimumValueDistance = newEntry;
                                    }

                                    balPopulation = firstPopulation - MinimumValueDistance.Value; //Population minus with Capacity

                                    posMin = listDict.FindIndex(x => x.Key == listDict.Min(y => y.Key));                                 

                                    //----------------------------------------------------------------------------------------------------------------------//

                                    if (balPopulation == 0)
                                    {
                                        listDict.Remove(new KeyValuePair<object, decimal>(MinimumValueDistance.Key, MinimumValueDistance.Value)); //remove distance and capacity because population have been fully used

                                        totalDistance = firstPopulation * MinimumValueDistance.Key;
                                    }

                                    else if (balPopulation < 0) //Capacity is more than population. So, capacity still have the balance. Then, proceed with next population. Delete first row
                                    {
                                        listDict.Clear();
                                        tempDistance.Clear();
                                        tempDistance2.Clear();

                                        currentRow++;

                                        shuffleList.RemoveAt(index: 0); // remove first element of population, then replace it with next of element.

                                        if (shuffleList.Count != 0)
                                        {
                                            tempDistance.Add(shuffleList[0].Value);
                                        }

                                        #region looping distance

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty1));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty2));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty3));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty4));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty5));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty6));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty7));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty8));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty9));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty10));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty11));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty12));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty13));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty14));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty15));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty16));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty17));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty18));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty19));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty20));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty21));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty22));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty23));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty24));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty25));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty26));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty27));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty28));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty29));
                                        }

                                        #endregion

                                        for (int i = currentCol; i < CapacityList.Count(); i++)
                                        {
                                            listDict.Add(new KeyValuePair<object, decimal>(tempDistance2[i], CapacityList[i].cProperty1));
                                        }

                                        balPopulation = balPopulation * -1; //if result turn out negative, then change it to positive

                                        decimal Population = firstPopulation;

                                        totalDistance = Population * MinimumValueDistance.Key; //calculate total distance travelled                                  

                                        if (currentCol < tempDistance2.Count())
                                        {
                                            newEntry = new KeyValuePair<object, decimal>(tempDistance2[currentCol], balPopulation); // either remove the previous column that have been fully used or set from which column that it need to start with 
                                        }

                                        totalAllocated = totalAllocated + firstPopulation; // Since first row have been completed, we take whole population in a row.

                                        break;
                                    }
                                    else //balPopulation is more than 0, it means population is more than capacity. Then, continue proceed with same population.
                                    {
                                        currentCol++;

                                        tempDistance2.Clear();

                                        tempCap = CapacityList[k - 1].cProperty1 - listDict[0].Value;

                                        int index = listDict.FindIndex(x => x.Key == listDict.Min(y => y.Key));

                                        listDict.Remove(listDict.FirstOrDefault(x => x.Key == listDict.Min(y => y.Key))); //remove distance and capacity because capacity have been fully used

                                        balCap = firstPopulation - balPopulation;

                                        firstPopulation = balPopulation;//update current population                   

                                        totalDistance = balCap * MinimumValueDistance.Key;

                                        totalDistance2 = totalDistance2 + totalDistance;

                                        tempDistance.Clear();

                                        tempDistance.Add(shuffleList[0].Value);

                                        #region looping distance

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty1));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty2));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty3));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty4));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty5));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty6));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty7));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty8));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty9));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty10));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty11));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty12));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty13));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty14));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty15));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty16));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty17));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty18));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty19));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty20));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty21));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty22));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty23));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty24));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty25));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty26));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty27));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty28));
                                        }

                                        foreach (DistanceClass a in tempDistance)
                                        {
                                            tempDistance2.Add(Convert.ToDecimal(a.dProperty29));
                                        }

                                        #endregion

                                        if (currentCol < tempDistance2.Count())
                                        {
                                            newEntry = new KeyValuePair<object, decimal>(tempDistance2[currentCol], listDict[0].Value);
                                        }

                                        totalAllocated = totalAllocated + balCap;

                                        if (currentCol == CapacityList.Count())
                                        {
                                            break;
                                        }

                                        txtReleaseCenter.Text += (posMin+1) + " " + "\r\n" + "\r\n";
                                    }

                                    

                                } //end of second loop  

                                totalDistance2 = totalDistance2 + totalDistance;

                                tempDistance2.Clear();

                                //txtReleaseCenter.Text += posMin + " " + "\r\n" + "\r\n";
                                
                            }                           

                        } //end of first loop                                                            

                        txtLoop.Text += "Loop " + ai + ": " + String.Format("{0:0.0000}", (totalDistance2 / totalPopulation)) + "\r\n" + "\r\n";

                        txtTravelDistance.Text = String.Format("{0:0.0000}", totalDistance2); // (sum travel distance/total population allocated)

                        txtTotalAllocated.Text = totalAllocated.ToString(); // Total Allocated

                        totalUnallocated = totalPopulation - totalAllocated;

                        txtTotalUnallocated.Text = totalUnallocated.ToString(); // Total Unallocated                      

                        balPopulation = 0;
                        currentCol = 0;
                        currentRow = 0;
                        totalDistance = 0;
                        totalDistance2 = 0;
                        totalPopulation = 0;
                        balCap = (dynamic)null;
                        firstPopulation = (dynamic)null;

                        MinimumValueDistance = (dynamic)null;
                        newEntry = (dynamic)null;
                    }
                    #endregion
                }
            }
        }

        private void txtNoLoop_Enter(object sender, EventArgs e)
        {
            txtNoLoop.Text = "";
            txtNoLoop.ForeColor = Color.Black;
        }

        private void txtNoLoop_Leave(object sender, EventArgs e)
        {
            //textBox6.Text = "Please put numbers of loop";

            //MessageBox.Show(textBox6.Text);

            if (txtNoLoop.Text == "")
            {
                txtNoLoop.Text = "Please put numbers of loop";
                txtNoLoop.ForeColor = Color.Gray;
            }
        }

        private void txtNoLoop_DoubleClick(object sender, EventArgs e)
        {
            //txtNoLoop.Clear();
            txtNoLoop.ForeColor = Color.Black;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "ID";
            xlWorkSheet.Cells[1, 2] = "Name";
            xlWorkSheet.Cells[2, 1] = "1";
            xlWorkSheet.Cells[2, 2] = "One";
            xlWorkSheet.Cells[3, 1] = "2";
            xlWorkSheet.Cells[3, 2] = "Two";



            xlWorkBook.SaveAs("C:\\Users\\RIDHWAN\\Desktop\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file in Desktop");
        }

    }
}
