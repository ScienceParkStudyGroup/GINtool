
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
        private bool LoadCategoryDataOld()
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



            string IDcol = "category_id";// Properties.Settings.Default.catCatIDColumn;
            string BSUcol = "locus_tag";// Properties.Settings.Default.catBSUColum;            
            // long list of columns... make cleaner later..

            gCategoriesWB.Columns.Add(IDcol, Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("catid_short", Type.GetType("System.String"));
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
                    lNewRow["catid_short"] = splits[splits.Count() - 1];
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

        private bool LoadCategoryData()
        {
            if (Properties.Settings.Default.categoryFile.Length == 0 || Properties.Settings.Default.catSheet.Length == 0)
            {
                btnCatFile.Label = "No file selected";
                return false;
            }

            AddTask(TASKS.LOAD_CATEGORY_DATA);

            SysData.DataTable _tmp = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.catSheet, Properties.Settings.Default.categoryFile, 1, 1);
            
            DataView __tmp = _tmp.DefaultView;
            __tmp.Sort = "[category id] asc";
            _tmp = __tmp.ToTable();

            gCategoriesWB = new SysData.DataTable("Categories")
            {
                CaseSensitive = false
            };

           
            DataTable _uCat = GetDistinctRecords(__tmp.ToTable(), new string[] { Properties.Settings.Default.catCatIDColumn, Properties.Settings.Default.catCatDescriptionColumn });
            DataView __uCat = _uCat.DefaultView;
            __uCat.Sort = String.Format("[{0}] asc", Properties.Settings.Default.catCatIDColumn);
            _uCat = __uCat.ToTable();
           

            // long list of columns... make cleaner later..

            gCategoriesWB.Columns.Add("catid", Type.GetType("System.String"));
            gCategoriesWB.Columns.Add("catid_short", Type.GetType("System.String"));
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




            //Properties.Settings.Default.catCatIDColumn; // de categorie code
            //Properties.Settings.Default.catBSUColum; // de code voor het gen
            //Properties.Settings.Default.catCatDescriptionColumn; // de beschrijving

            //TreeNode treeNode = new TreeNode("Categories");
            //for (int i = 0; i < _tmp.Rows.Count; i++)
            //{
            //    DataRow row = _tmp.Rows[i];
            //    string sw_code = (string)row[Properties.Settings.Default.catCatIDColumn];
            //    if (sw_code != null)
            //    {
            //        sw_code = sw_code.Replace("SW.", "");
            //        int[] codes = sw_code.Split('.').Select(str => Int32.Parse(str.Trim())).ToArray();

            //        TreeNode[] lInsNode = treeNode.Nodes.Find("sw_code", true);
            //        if (lInsNode != null)
            //        {
            //            //lInsNode.Add()
            //        }

            //    }
            //}
            string[] lcols = new string[] { "cat1_int", "cat2_int", "cat3_int", "cat4_int", "cat5_int" };
            string[] ulcols = new string[] { "ucat1_int", "ucat2_int", "ucat3_int", "ucat4_int", "ucat5_int" };

            foreach (SysData.DataRow lRow in _tmp.Rows)
            {

                string sw_code = "";
                string lc_tag = "";
                string sw_short = "";

                try
                {
                    sw_code = (string)lRow[Properties.Settings.Default.catCatIDColumn];
                    lc_tag = (string)lRow[Properties.Settings.Default.catBSUColum];
                    sw_short = sw_code.Replace("SW.", "");
                }
                catch
                {
                    continue;
                }

                

                int[] codes = sw_short.Split('.').Select(str => Int32.Parse(str.Trim())).ToArray();
                int lvl = codes.Length;

                SysData.DataRow lNewRow = gCategoriesWB.Rows.Add();
                lNewRow["catid"] = sw_code;
                lNewRow["catid_short"] = sw_short;
                lNewRow["locus_tag"] = lc_tag;
                string catCode = String.Format("SW.{0}", codes[0]);
                string _filter = String.Format("[{0}] = '{1}'", Properties.Settings.Default.catCatIDColumn, catCode);
                __uCat.RowFilter = _filter;
                DataTable _desc = __uCat.ToTable();
                string catDesc = "";
                if (_desc.Rows.Count > 0)
                    catDesc = (string)_desc.Rows[0][1];
                else
                    catDesc = String.Format("category_{0}", codes[0]);

                lNewRow["cat1"] = catDesc;
                lNewRow["cat1_int"] = codes[0];               
                lNewRow["ucat1_int"] = Math.Pow(10, 5) * codes[0];
                int offset = (Int32)Math.Pow(10, 5) * codes[0];


                for (int l=1;l<lvl;l++)
                {
                    catCode = String.Format("{0}.{1}", catCode, codes[l]);
                    _filter = String.Format("[{0}] = '{1}'", Properties.Settings.Default.catCatIDColumn, catCode);
                    __uCat.RowFilter = _filter;
                    _desc = __uCat.ToTable();
                    if (_desc.Rows.Count > 0)
                        catDesc = (string)_desc.Rows[0][1];
                    else
                    {
                        string _fmt = String.Join(".", codes.Take(l+1).Select(iv => iv.ToString()).ToArray());
                        catDesc = String.Format("category_{0}", _fmt);
                    }

                    string colName = String.Format("cat{0}", l+1);
                    lNewRow[colName] = catDesc;
                    colName = String.Format("cat{0}_int", l+1);
                    lNewRow[colName] = codes[l];
                    colName = String.Format("ucat{0}_int", l+1);
                    lNewRow[colName] = Math.Pow(10, 5-l) * codes[l];
                    offset = offset + (Int32)Math.Pow(10, 5 - l) * codes[l];


                }
                
                
                
                
                
                
                //object[] lItems = lRow.ItemArray;
                //SysData.DataRow lNewRow = gCategoriesWB.Rows.Add();
                //for (int i = 0; i < lItems.Length; i++)
                //{

                    



                //    lNewRow["catid"] = lItems[0];
                //    string[] splits = lItems[0].ToString().Split(' ');
                //    //lNewRow["catid_short"] = splits[splits.Count() - 1];
                //    lNewRow["category"] = lItems[1];
                //    lNewRow["locus_tag"] = lItems[2];
                //    //lNewRow["cat1"] = lItems[3];
                //    //lNewRow["cat2"] = lItems[4];
                //    //lNewRow["cat3"] = lItems[5];
                //    //lNewRow["cat4"] = lItems[6];
                //    //lNewRow["cat5"] = lItems[7];

                //    string[] llItems = lItems[0].ToString().Split('.').Skip(1).ToArray();


                //    // start at 1.. that contains SW
                //    for (int j = 0; j < llItems.Length; j++)
                //    {
                //        lNewRow[lcols[j]] = Int32.Parse(llItems[j]);
                //    }

                //    lNewRow[String.Format("cat{0}", llItems.Length)] = lItems[1];

                //    int offset = 0;
                //    for (int j = 0; j < llItems.Length; j++)
                //    {
                //        lNewRow[ulcols[j]] = offset + ((Int32)lNewRow[lcols[j]]) * Math.Pow(10, 5 - j);
                //        offset = (Int32)lNewRow[ulcols[j]];
                //    }

                //}
            }

            RemoveTask(TASKS.LOAD_CATEGORY_DATA);

            return gCategoriesWB.Rows.Count > 0;
        }


        /// <summary>
        /// Import the category data from the csv file downloaded from http://subtiwiki.uni-goettingen.de/. 
        /// This routine is specifically made for that specific data format. d.d. 14/2/2022
        /// </summary>
        /// <returns></returns>
        private bool LoadCategoryDataColumns()
        {
            if (Properties.Settings.Default.categoryFile.Length == 0 || Properties.Settings.Default.catSheet.Length == 0)
            {
                btnCatFile.Label = "No file selected";
                return false;
            }            

            gCategoryColNames = ExcelUtils.ReadExcelToDatableHeader(gApplication, Properties.Settings.Default.catSheet, Properties.Settings.Default.categoryFile, 1, 1);

            return gCategoryColNames.Length > 0;
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

        private bool LoadGenesDataColumns()
        {

            if (Properties.Settings.Default.genesFileName.Length == 0 || Properties.Settings.Default.genesSheetName.Length == 0)
            {
                btnGenesFileSelected.Label = "No file selected";
                return false;
            }

            gGenesColNames = ExcelUtils.ReadExcelToDatableHeader(gApplication, Properties.Settings.Default.genesSheetName, Properties.Settings.Default.genesFileName, 1, 1);

            
            return gGenesColNames.Length>0;

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

        private bool LoadRegulonDataColumns()
        {
            if (Properties.Settings.Default.referenceFile.Length == 0 || Properties.Settings.Default.referenceSheetName.Length == 0)
            {
                btnRegulonFileName.Label = "No file selected";
                return false;
            }            

            gRegulonColNames = ExcelUtils.ReadExcelToDatableHeader(gApplication, Properties.Settings.Default.referenceSheetName, Properties.Settings.Default.referenceFile, 1, 1);
            
            return gRegulonColNames.Length > 0;
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
