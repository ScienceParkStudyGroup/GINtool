
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SysData = System.Data;



namespace GINtool
{
    public partial class GinRibbon
    {

        /// <summary>
        /// Import the category data from the csv file downloaded from http://subtiwiki.uni-goettingen.de/. 
        /// This routine is specifically made for that specific data format.
        /// </summary>
        /// <returns></returns>
        private bool LoadCategoryData()
        {
            if (Properties.Settings.Default.categoryFile.Length == 0 || Properties.Settings.Default.catSheet.Length == 0)
            {
                btnCatFile.Label = "No file selected";
                return false;
            }

            AddTask(TASKS.LOAD_CATEGORY_DATA);

            SysData.DataTable _tmp = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.catSheet, Properties.Settings.Default.categoryFile, 1, 1);
            if (_tmp != null)
            {
                gCategoryColNames = new string[_tmp.Columns.Count];
                int i = 0;
                foreach (SysData.DataColumn col in _tmp.Columns)
                {
                    gCategoryColNames[i++] = col.ColumnName;
                }
            }



            gCategoriesWB = new SysData.DataTable("Categories")
            {
                CaseSensitive = false
            };



            string IDcol = Properties.Settings.Default.catCatIDColumn;
            string BSUcol = Properties.Settings.Default.catBSUColum;            
            // long list of columns... make cleaner later..

            gCategoriesWB.Columns.Add(IDcol, Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat_short", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add(BSUcol, Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("gene", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat1", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat2", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat3", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat4", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat5", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat1_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("cat2_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("cat3_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("cat4_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("cat5_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("ucat1_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("ucat2_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("ucat3_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("ucat4_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("ucat5_int", Type.GetType("System.Int32"));


            gCategoryColNames = new string[gCategoriesWB.Columns.Count];
            if (gCategoriesWB != null)
            {
                gCategoryColNames = new string[gCategoriesWB.Columns.Count];
                int i = 0;
                foreach (SysData.DataColumn col in gCategoriesWB.Columns)
                {
                    gCategoryColNames[i++] = col.ColumnName;
                }
            }            


            string[] lcols = new string[] { "cat1_int", "cat2_int", "cat3_int", "cat4_int", "cat5_int" };
            string[] ulcols = new string[] { "ucat1_int", "ucat2_int", "ucat3_int", "ucat4_int", "ucat5_int" };

            foreach (SysData.DataRow lRow in _tmp.Rows)
            {
                object[] lItems = lRow.ItemArray;
                SysData.DataRow lNewRow = gCategoriesWB.Rows.Add();
                for (int i = 0; i < lItems.Length; i++)
                {
                    lNewRow[IDcol] = lItems[0];
                    string[] splits = lItems[0].ToString().Split(' ');
                    lNewRow["cat_short"] = splits[splits.Count() - 1];
                    lNewRow[BSUcol] = lItems[1];
                    lNewRow["Gene"] = lItems[2];
                    lNewRow["cat1"] = lItems[3];
                    lNewRow["cat2"] = lItems[4];
                    lNewRow["cat3"] = lItems[5];
                    lNewRow["cat4"] = lItems[6];
                    lNewRow["cat5"] = lItems[7];

                    string[] llItems = lItems[0].ToString().Split(' ')[1].Split('.');



                    for (int j = 0; j < llItems.Length; j++)
                    {
                        lNewRow[lcols[j]] = Int32.Parse(llItems[j]);
                    }

                    int offset = 0;
                    for (int j = 0; j < llItems.Length; j++)
                    {
                        lNewRow[ulcols[j]] = offset + ((Int32)lNewRow[lcols[j]]) * Math.Pow(10, 5 - j);
                        offset = (Int32)lNewRow[ulcols[j]];
                    }

                }
            }
            
            //gCategoriesWB.PrimaryKey = new DataColumn[] { gCategoriesWB.Columns[BSUcol]};

            RemoveTask(TASKS.LOAD_CATEGORY_DATA);

            return gCategoriesWB.Rows.Count > 0;
        }


        /// <summary>
        /// Import the category data from the csv file downloaded from http://subtiwiki.uni-goettingen.de/. 
        /// This routine is specifically made for that specific data format. d.d. 14/2/2022
        /// </summary>
        /// <returns></returns>
        private bool LoadCategoryDataNewFormat()
        {
            if (Properties.Settings.Default.categoryFile.Length == 0 || Properties.Settings.Default.catSheet.Length == 0)
            {
                btnCatFile.Label = "No file selected";
                return false;
            }

            AddTask(TASKS.LOAD_CATEGORY_DATA);

            SysData.DataTable _tmp = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.catSheet, Properties.Settings.Default.categoryFile, 1, 1);


            if (_tmp != null)
            {
                gCategoryColNames = new string[_tmp.Columns.Count];
                int i = 0;
                foreach (SysData.DataColumn col in _tmp.Columns)
                {
                    gCategoryColNames[i++] = col.ColumnName;
                }
            }

            DataView __tmp = _tmp.DefaultView;
            __tmp.Sort = "[category id] asc";
            _tmp = __tmp.ToTable();

            gCategoriesWB = new SysData.DataTable("Categories")
            {
                CaseSensitive = false
            };

            // long list of columns... make cleaner later..

            gCategoriesWB.Columns.Add("catid", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("category", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("gene", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("locus_tag", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat1", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat2", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat3", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat4", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat5", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("cat1_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("cat2_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("cat3_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("cat4_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("cat5_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("ucat1_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("ucat2_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("ucat3_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("ucat4_int", Type.GetType("System.Int32"));
            gCategoriesWB.Columns.Add("ucat5_int", Type.GetType("System.Int32"));


            string[] lcols = new string[] { "cat1_int", "cat2_int", "cat3_int", "cat4_int", "cat5_int" };
            string[] ulcols = new string[] { "ucat1_int", "ucat2_int", "ucat3_int", "ucat4_int", "ucat5_int" };

            foreach (SysData.DataRow lRow in _tmp.Rows)
            {
                object[] lItems = lRow.ItemArray;
                SysData.DataRow lNewRow = gCategoriesWB.Rows.Add();
                for (int i = 0; i < lItems.Length; i++)
                {
                    lNewRow["catid"] = lItems[0];
                    string[] splits = lItems[0].ToString().Split(' ');
                    //lNewRow["catid_short"] = splits[splits.Count() - 1];
                    lNewRow["category"] = lItems[1];
                    lNewRow["locus_tag"] = lItems[2];
                    //lNewRow["cat1"] = lItems[3];
                    //lNewRow["cat2"] = lItems[4];
                    //lNewRow["cat3"] = lItems[5];
                    //lNewRow["cat4"] = lItems[6];
                    //lNewRow["cat5"] = lItems[7];

                    string[] llItems = lItems[0].ToString().Split('.').Skip(1).ToArray();


                    // start at 1.. that contains SW
                    for (int j = 0; j < llItems.Length; j++)
                    {
                        lNewRow[lcols[j]] = Int32.Parse(llItems[j]);
                    }

                    lNewRow[String.Format("cat{0}", llItems.Length)] = lItems[1];

                    int offset = 0;
                    for (int j = 0; j < llItems.Length; j++)
                    {
                        lNewRow[ulcols[j]] = offset + ((Int32)lNewRow[lcols[j]]) * Math.Pow(10, 5 - j);
                        offset = (Int32)lNewRow[ulcols[j]];
                    }

                }
            }

            RemoveTask(TASKS.LOAD_CATEGORY_DATA);

            return gCategoriesWB.Rows.Count > 0;
        }

        /// <summary>
        /// Load the operon data from the specified csv file as downloaded from http://subtiwiki.uni-goettingen.de/.        
        /// </summary>
        /// <returns></returns>

        private bool LoadOperonData()
        {

            if (Properties.Settings.Default.operonFile.Length == 0 || Properties.Settings.Default.operonSheet.Length == 0)
            {
                btnOperonFile.Label = "No file selected";
                return false;
            }

            AddTask(TASKS.LOAD_OPERON_DATA);

            SysData.DataTable _tmp = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.operonSheet, Properties.Settings.Default.operonFile, 1, 1);
            gRefOperonsWB = new SysData.DataTable("OPERONS")
            {
                CaseSensitive = false
            };

            gRefOperonsWB.Columns.Add("operon", Type.GetType("System.String"));
            gRefOperonsWB.Columns.Add("gene", Type.GetType("System.String"));
            gRefOperonsWB.Columns.Add("op_id", Type.GetType("System.Int32"));

            int _op_id = 0;

            foreach (SysData.DataRow lRow in _tmp.Rows)
            {
                string[] lItems = lRow.ItemArray[0].ToString().Split('-');

                if (maxGenesPerOperon < lItems.Length)
                    maxGenesPerOperon = lItems.Length;

                for (int i = 0; i < lItems.Length; i++)
                {
                    SysData.DataRow lNewRow = gRefOperonsWB.Rows.Add();
                    lNewRow["operon"] = lItems[0];
                    lNewRow["gene"] = lItems[i];
                    lNewRow["op_id"] = _op_id;
                }

                _op_id++;
            }

            RemoveTask(TASKS.LOAD_OPERON_DATA);
            return gRefOperonsWB.Rows.Count > 0;
        }


        /// <summary>
        /// Load the main genes information data from a csv file
        /// </summary>
        /// <returns></returns>
        private bool LoadGenesData()
        {

            if (Properties.Settings.Default.genesFileName.Length == 0 || Properties.Settings.Default.genesSheetName.Length == 0)
            {
                btnGenesFileSelected.Label = "No file selected";
                return false;
            }

            AddTask(TASKS.LOAD_GENES_DATA);
            gGenesWB = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.genesSheetName, Properties.Settings.Default.genesFileName, 1, 1);


            if (gGenesWB != null)
            {
                gGenesColNames = new string[gGenesWB.Columns.Count];
                int i = 0;
                foreach (SysData.DataColumn col in gGenesWB.Columns)
                {
                    gGenesColNames[i++] = col.ColumnName;
                }
            }

            gGenesWB.PrimaryKey = new DataColumn[] { gGenesWB.Columns[Properties.Settings.Default.genesBSUColumn] };


            RemoveTask(TASKS.LOAD_GENES_DATA);
            return gGenesWB != null;

        }


        /// <summary>
        /// Load the main Regulon data as downloaded from http://subtiwiki.uni-goettingen.de/. 
        /// The whole add-in is written for data in that specific format!
        /// </summary>
        /// <returns></returns>
        private bool LoadRegulonData()
        {
            if (Properties.Settings.Default.referenceFile.Length == 0 || Properties.Settings.Default.referenceSheetName.Length == 0)
            {
                btnRegulonFileName.Label = "No file selected";
                return false;
            }
            AddTask(TASKS.LOAD_REGULON_DATA);

            gRegulonWB = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.referenceSheetName, Properties.Settings.Default.referenceFile, 1, 1);
            if (gRegulonWB != null)
            {
                gRegulonColNames = new string[gRegulonWB.Columns.Count];
                int i = 0;
                foreach (SysData.DataColumn col in gRegulonWB.Columns)
                {
                    gRegulonColNames[i++] = col.ColumnName;
                }
                // generate database frequency table
                // CreateTableStatistics();
            }

            RemoveTask(TASKS.LOAD_REGULON_DATA);
            return gRegulonWB != null;
        }


        /// <summary>
        /// Specify the regulon data  
        /// </summary>
        /// 
        private void SpecifyRegulonWorksheets()
        {
            Microsoft.Office.Interop.Excel.Application excel = (Microsoft.Office.Interop.Excel.Application)Globals.ThisAddIn.Application;
            excel.DisplayAlerts = false;
            excel.EnableEvents = false;

            Excel.Workbook excelworkBook = excel.Workbooks.Open(Properties.Settings.Default.referenceFile);
            // Set workbook to first worksheet
            Excel.Worksheet ws = (Excel.Worksheet)excelworkBook.Sheets[1];
            Properties.Settings.Default.referenceSheetName = ws.Name;


            excelworkBook.Close();
            gRegulonFileSelected = true;
            excel.EnableEvents = true;
            excel.DisplayAlerts = true;
        }


        /// <summary>
        /// Specify the genes data  
        /// </summary>

        private void SpecifyGenesWorksheets()
        {
            Microsoft.Office.Interop.Excel.Application excel = (Microsoft.Office.Interop.Excel.Application)Globals.ThisAddIn.Application;
            excel.DisplayAlerts = false;
            excel.EnableEvents = false;

            Excel.Workbook excelworkBook = excel.Workbooks.Open(Properties.Settings.Default.genesFileName);
            // Set workbook to first worksheet
            Excel.Worksheet ws = (Excel.Worksheet)excelworkBook.Sheets[1];
            Properties.Settings.Default.genesSheetName = ws.Name;


            excelworkBook.Close();
            gGenesFileSelected = true;
            excel.EnableEvents = true;
            excel.DisplayAlerts = true;
        }


        /// <summary>
        /// Specify the operon data
        /// </summary>
        private void SpecifyOperonSheet()
        {
            Microsoft.Office.Interop.Excel.Application excel = (Microsoft.Office.Interop.Excel.Application)Globals.ThisAddIn.Application;
            excel.DisplayAlerts = false;
            excel.EnableEvents = false;

            Excel.Workbook excelworkBook = excel.Workbooks.Open(Properties.Settings.Default.operonFile);
            // Set workbook to first worksheet
            Excel.Worksheet ws = (Excel.Worksheet)excelworkBook.Sheets[1];
            Properties.Settings.Default.operonSheet = ws.Name;

            excelworkBook.Close();

            excel.EnableEvents = true;
            excel.DisplayAlerts = true;
        }

        /// <summary>
        /// Specify the operon data
        /// </summary>
        private void SpecifyCatFile()
        {
            Microsoft.Office.Interop.Excel.Application excel = (Microsoft.Office.Interop.Excel.Application)Globals.ThisAddIn.Application;
            excel.DisplayAlerts = false;
            excel.EnableEvents = false;

            Excel.Workbook excelworkBook = excel.Workbooks.Open(Properties.Settings.Default.categoryFile);
            // Set workbook to first worksheet
            Excel.Worksheet ws = (Excel.Worksheet)excelworkBook.Sheets[1];
            Properties.Settings.Default.catSheet = ws.Name;


            excelworkBook.Close();
            gCategoryFileSelected = true;
            excel.EnableEvents = true;
            excel.DisplayAlerts = true;
        }


    }
}
