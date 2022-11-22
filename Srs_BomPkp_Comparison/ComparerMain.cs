using DevExpress.Internal.WinApi;
using DevExpress.Utils;
using DevExpress.Utils.Extensions;
using DevExpress.XtraBars;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraVerticalGrid;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Srs_BomPkp_Comparison
{
    public partial class ComparerMain : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        XlsxFile XlsxFile1 = new XlsxFile();
        XlsxFile XlsxFile2 = new XlsxFile();

        string sheet_name;
        public ComparerMain()
        {
            InitializeComponent();
        }

        private void ParseComponentsForDesignFile(XlsxFile xslxFile, GridView gv)
        {
            for (int i = 0; i < gv.RowCount - 1; i++)
            {
                string referenceLine = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2Refference.SelectedItem.ToString()));
                if (referenceLine.Contains(',') && referenceLine != null && referenceLine != "")
                {
                    string[] refferences = referenceLine.Split(',');
                    foreach (string reffrence in refferences)
                    {
                        if (reffrence.Contains('-'))
                        {
                            string[] referencestartend = reffrence.Split('-');

                            string[] startarr = Regex.Split(referencestartend[0], @"\D+");
                            string[] endarr = Regex.Split(referencestartend[1], @"\D+");

                            string letter = new String(referencestartend[0].Where(Char.IsLetter).ToArray());

                            int start = Convert.ToInt32(startarr[1]);
                            int end = Convert.ToInt32(endarr[1]);
                            for (int inc = start; inc <= end; inc++)
                            {
                                string newrefference = letter + inc.ToString();
                                Components component = new Components();
                                component.Refference = newrefference;
                                if (Cbox_File2CustomerStockCode.SelectedItem != null)
                                    component.CustomerStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2CustomerStockCode.SelectedItem.ToString()));
                                component.SiriusStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2SiriusStockCode.SelectedItem.ToString()));
                                component.StockName = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2StockName.SelectedItem.ToString()));
                                component.PartNo = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2PartNo.SelectedItem.ToString()));
                                xslxFile.component.Add(component);
                            }
                        }
                        else
                        {
                            Components component = new Components();
                            component.Refference = reffrence;
                            if (Cbox_File2CustomerStockCode.SelectedItem != null)
                                component.CustomerStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2CustomerStockCode.SelectedItem.ToString()));
                            component.SiriusStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2SiriusStockCode.SelectedItem.ToString()));
                            component.StockName = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2StockName.SelectedItem.ToString()));
                            component.PartNo = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2PartNo.SelectedItem.ToString()));
                            xslxFile.component.Add(component);
                        }
                    }
                }
                else if (referenceLine != null && referenceLine != "")
                {
                    string[] refferences = referenceLine.Split(',');
                    foreach (string reffrence in refferences)
                    {
                        if (reffrence.Contains('-'))
                        {
                            string[] referencestartend = reffrence.Split('-');

                            string[] startarr = Regex.Split(referencestartend[0], @"\D+");
                            string[] endarr = Regex.Split(referencestartend[1], @"\D+");

                            string letter = new String(referencestartend[0].Where(Char.IsLetter).ToArray());

                            int start = Convert.ToInt32(startarr[1]);
                            int end = Convert.ToInt32(endarr[1]);
                            for (int inc = start; inc <= end; inc++)
                            {
                                string newrefference = letter + inc.ToString();
                                Components component = new Components();
                                component.Refference = newrefference;
                                if (Cbox_File2CustomerStockCode.SelectedItem != null)
                                    component.CustomerStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2CustomerStockCode.SelectedItem.ToString()));
                                component.SiriusStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2SiriusStockCode.SelectedItem.ToString()));
                                component.StockName = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2StockName.SelectedItem.ToString()));
                                component.PartNo = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2PartNo.SelectedItem.ToString()));
                                xslxFile.component.Add(component);
                            }
                        }
                        else
                        {
                            Components component = new Components();
                            component.Refference = reffrence;
                            if (Cbox_File2CustomerStockCode.SelectedItem != null)
                                component.CustomerStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2CustomerStockCode.SelectedItem.ToString()));
                            component.SiriusStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2SiriusStockCode.SelectedItem.ToString()));
                            component.StockName = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2StockName.SelectedItem.ToString()));
                            component.PartNo = Convert.ToString(gv.GetRowCellValue(i, Cbox_File2PartNo.SelectedItem.ToString()));
                            xslxFile.component.Add(component);
                        }
                    }
                }
            }
        }

        private void ParseComponentsForProductionFile(XlsxFile xslxFile, GridView gv)
        {
            for (int i = 0; i < gv.RowCount - 1; i++)
            {
                string referenceLine = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1Refference.SelectedItem.ToString()));
                if (referenceLine.Contains(',') && referenceLine != null && referenceLine != "")
                {
                    string[] refferences = referenceLine.Split(',');
                    foreach (string reffrence in refferences)
                    {
                        if (reffrence.Contains('-'))
                        {
                            string[] referencestartend = reffrence.Split('-');

                            string[] startarr = Regex.Split(referencestartend[0], @"\D+");
                            string[] endarr = Regex.Split(referencestartend[1], @"\D+");

                            string letter = new String(referencestartend[0].Where(Char.IsLetter).ToArray());

                            int start = Convert.ToInt32(startarr[1]);
                            int end = Convert.ToInt32(endarr[1]);
                            for (int inc = start; inc <= end; inc++)
                            {
                                string newrefference = letter + inc.ToString();
                                Components component = new Components();
                                component.Refference = newrefference;
                                if (Cbox_File1CustomerStockCode.SelectedItem != null)
                                    component.CustomerStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1CustomerStockCode.SelectedItem.ToString()));
                                component.SiriusStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1SiriusStockCode.SelectedItem.ToString()));
                                component.StockName = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1StockName.SelectedItem.ToString()));
                                component.PartNo = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1PartNo.SelectedItem.ToString()));
                                xslxFile.component.Add(component);
                            }
                        }
                        else
                        {
                            Components component = new Components();
                            component.Refference = reffrence;
                            if (Cbox_File1CustomerStockCode.SelectedItem != null)
                                component.CustomerStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1CustomerStockCode.SelectedItem.ToString()));
                            component.SiriusStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1SiriusStockCode.SelectedItem.ToString()));
                            component.StockName = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1StockName.SelectedItem.ToString()));
                            component.PartNo = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1PartNo.SelectedItem.ToString()));
                            xslxFile.component.Add(component);
                        }
                    }
                }
                else if (referenceLine != null && referenceLine != "")
                {
                    string[] refferences = referenceLine.Split(',');
                    foreach (string reffrence in refferences)
                    {
                        if (reffrence.Contains('-'))
                        {
                            string[] referencestartend = reffrence.Split('-');

                            string[] startarr = Regex.Split(referencestartend[0], @"\D+");
                            string[] endarr = Regex.Split(referencestartend[1], @"\D+");

                            string letter = new String(referencestartend[0].Where(Char.IsLetter).ToArray());

                            int start = Convert.ToInt32(startarr[1]);
                            int end = Convert.ToInt32(endarr[1]);
                            for (int inc = start; inc <= end; inc++)
                            {
                                string newrefference = letter + inc.ToString();
                                Components component = new Components();
                                component.Refference = newrefference;
                                if (Cbox_File1CustomerStockCode.SelectedItem != null)
                                    component.CustomerStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1CustomerStockCode.SelectedItem.ToString()));
                                component.SiriusStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1SiriusStockCode.SelectedItem.ToString()));
                                component.StockName = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1StockName.SelectedItem.ToString()));
                                component.PartNo = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1PartNo.SelectedItem.ToString()));
                                xslxFile.component.Add(component);
                            }
                        }
                        else
                        {
                            Components component = new Components();
                            component.Refference = reffrence;
                            if (Cbox_File2CustomerStockCode.SelectedItem != null)
                                component.CustomerStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1CustomerStockCode.SelectedItem.ToString()));
                            component.SiriusStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1SiriusStockCode.SelectedItem.ToString()));
                            component.StockName = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1StockName.SelectedItem.ToString()));
                            component.PartNo = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1PartNo.SelectedItem.ToString()));
                            xslxFile.component.Add(component);
                        }
                    }
                }
            }
            //for (int i = 0; i < gv.RowCount - 1; i++)
            //{
            //    string referenceLine = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1Refference.SelectedItem.ToString()));
            //    if (referenceLine.Contains(',') && referenceLine != null && referenceLine != "")
            //    {
            //        string[] refferences = referenceLine.Split(',');
            //        foreach (string reffrence in refferences)
            //        {
            //            Components component = new Components();
            //            component.Refference = reffrence;
            //            component.CustomerStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1CustomerStockCode.SelectedItem.ToString()));
            //            component.SiriusStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1SiriusStockCode.SelectedItem.ToString()));
            //            component.StockName = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1StockName.SelectedItem.ToString()));
            //            component.PartNo = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1PartNo.SelectedItem.ToString()));
            //            xslxFile.component.Add(component);
            //        }
            //    }
            //    else if (referenceLine != null && referenceLine != "")
            //    {
            //        Components component = new Components();
            //        component.Refference = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1Refference.SelectedItem.ToString()));
            //        component.CustomerStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1CustomerStockCode.SelectedItem.ToString()));
            //        component.SiriusStockCode = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1SiriusStockCode.SelectedItem.ToString()));
            //        component.StockName = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1StockName.SelectedItem.ToString()));
            //        component.PartNo = Convert.ToString(gv.GetRowCellValue(i, Cbox_File1PartNo.SelectedItem.ToString()));
            //        xslxFile.component.Add(component);
            //    }
            //}
        }

        private void OpenXslxForProductionFile(string filePath, GridControl gc, GridView gv)
        {
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml; HDR=YES;'");
            conn.Open();
            DataTable dta = new DataTable();
            dta = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            foreach (DataRow item in dta.Rows)
            {
                sheet_name = Convert.ToString(item["TABLE_NAME"]);
                break;
            }

            OleDbDataAdapter data1 = new OleDbDataAdapter("Select * From [" + sheet_name + "]", conn);

            DataTable dt = new DataTable();
            data1.Fill(dt);
            gc.BeginUpdate();

            try
            {
                gv.Columns.Clear();
                gc.DataSource = null;

                gc.DataSource = dt;
            }
            finally
            {
                gc.EndUpdate();
            }
            int i = 0;
            foreach (GridColumn column in gv.Columns)
            {
                Cbox_File1CustomerStockCode.Properties.Items.Add(column.FieldName);
                Cbox_File1CustomerStockCode.SelectedIndex = 1;
                Cbox_File1SiriusStockCode.Properties.Items.Add(column.FieldName);
                Cbox_File1SiriusStockCode.SelectedIndex = 2;
                Cbox_File1StockName.Properties.Items.Add(column.FieldName);
                Cbox_File1StockName.SelectedIndex = 3;
                Cbox_File1PartNo.Properties.Items.Add(column.FieldName);
                Cbox_File1PartNo.SelectedIndex = 5;
                Cbox_File1Refference.Properties.Items.Add(column.FieldName);
                Cbox_File1Refference.SelectedIndex = 6;
                column.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                i++;
            }
            dataGridView1.DataSource = dt;
            gv.BestFitColumns();
        }

        private void OpenXslxForDesignFile(string filePath, GridControl gc, GridView gv)
        {
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0 Xml; HDR=YES;'");
            conn.Open();
            DataTable dta = new DataTable();
            dta = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            foreach (DataRow item in dta.Rows)
            {
                sheet_name = Convert.ToString(item["TABLE_NAME"]);
                break;
            }

            OleDbDataAdapter data1 = new OleDbDataAdapter("Select * From [" + sheet_name + "]", conn);

            DataTable dt = new DataTable();
            data1.Fill(dt);

            gc.BeginUpdate();

            try
            {
                gv.Columns.Clear();
                gc.DataSource = null;

                gc.DataSource = dt;
            }
            finally
            {
                gc.EndUpdate();
            }
            int i = 0;
            foreach (GridColumn column in gv.Columns)
            {
                Cbox_File2CustomerStockCode.Properties.Items.Add(column.FieldName);
                Cbox_File2SiriusStockCode.Properties.Items.Add(column.FieldName);
                Cbox_File2SiriusStockCode.SelectedIndex = 6;
                Cbox_File2StockName.Properties.Items.Add(column.FieldName);
                Cbox_File2StockName.SelectedIndex = 5;
                Cbox_File2PartNo.Properties.Items.Add(column.FieldName);
                Cbox_File2PartNo.SelectedIndex = 4;
                Cbox_File2Refference.Properties.Items.Add(column.FieldName);
                Cbox_File2Refference.SelectedIndex = 2;
                column.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                i++;
            }
            dataGridView2.DataSource = dt;
            dataGridView2.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView2.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView2.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView2.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            gv.BestFitColumns();
        }

        private void Btn_Compare_ItemClick(object sender, ItemClickEventArgs e)
        {
            barListItem1.Strings.Clear();
            barListItem2.Strings.Clear();
            ParseComponentsForProductionFile(XlsxFile1, gridView1);
            ParseComponentsForDesignFile(XlsxFile2, gridView2);
            var difList = XlsxFile2.component.Where(a => !XlsxFile1.component.Any(a1 => a1.Refference == a.Refference));

            var difList1 = XlsxFile1.component.Where(a => !XlsxFile2.component.Any(a1 => a1.Refference == a.Refference));

            var commons = XlsxFile1.component.Select(s1 => s1.Refference).Intersect(XlsxFile2.component.Select(s2 => s2.Refference));

            foreach (var item in commons)
            {
                var element1 = XlsxFile1.component.Find(x => x.Refference.Equals(item));
                var element2 = XlsxFile2.component.Find(x => x.Refference.Equals(item));


                if (element1.SiriusStockCode == element2.SiriusStockCode)
                {
                    for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                    {
                        string val = dataGridView2.Rows[i].Cells[Cbox_File2SiriusStockCode.SelectedIndex].Value.ToString();
                        if (val == element2.SiriusStockCode)
                        {
                            dataGridView2.Rows[i].Cells[Cbox_File2SiriusStockCode.SelectedIndex].Style.BackColor = Color.PaleGreen;
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                    {
                        string val = dataGridView2.Rows[i].Cells[Cbox_File2SiriusStockCode.SelectedIndex].Value.ToString();
                        if (val == element2.SiriusStockCode)
                        {
                            dataGridView2.Rows[i].Cells[Cbox_File2SiriusStockCode.SelectedIndex].Style.BackColor = Color.PaleVioletRed;
                        }
                    }
                }

                for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                {
                    string val = dataGridView2.Rows[i].Cells[Cbox_File2Refference.SelectedIndex].Value.ToString();
                    if (val.Contains(item))
                    {
                            dataGridView2.Rows[i].Cells[Cbox_File2Refference.SelectedIndex].Style.BackColor = Color.PaleGreen;

                    }
                }

                if (element1.SiriusStockCode == element2.SiriusStockCode)
                {
                    for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                    {
                        string val = dataGridView1.Rows[i].Cells[Cbox_File1SiriusStockCode.SelectedIndex].Value.ToString();
                        if (val == element1.SiriusStockCode)
                        {
                            dataGridView1.Rows[i].Cells[Cbox_File1SiriusStockCode.SelectedIndex].Style.BackColor = Color.PaleGreen;
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                    {
                        string val = dataGridView1.Rows[i].Cells[Cbox_File1SiriusStockCode.SelectedIndex].Value.ToString();
                        if (val == element1.SiriusStockCode)
                        {
                            dataGridView1.Rows[i].Cells[Cbox_File1SiriusStockCode.SelectedIndex].Style.BackColor = Color.PaleVioletRed;
                        }
                    }
                }

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    string val = dataGridView1.Rows[i].Cells[Cbox_File1Refference.SelectedIndex].Value.ToString();
                    if (val.Contains(item))
                    {
                            dataGridView1.Rows[i].Cells[Cbox_File1Refference.SelectedIndex].Style.BackColor = Color.PaleGreen;

                    }
                }




                //if (element1.SiriusStockCode == element2.SiriusStockCode)
                //{
                //    for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                //    {
                //        string val = dataGridView1.Rows[i].Cells[Cbox_File1SiriusStockCode.SelectedIndex].Value.ToString();
                //        if (val == element1.SiriusStockCode)
                //        {
                //            dataGridView1.Rows[i].Cells[Cbox_File1SiriusStockCode.SelectedIndex].Style.BackColor = Color.PaleGreen;
                //        }
                //    }
                //}
                //else
                //{
                //    for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                //    {
                //        string val = dataGridView1.Rows[i].Cells[Cbox_File1SiriusStockCode.SelectedIndex].Value.ToString();
                //        if (val == element1.SiriusStockCode)
                //        {
                //            dataGridView1.Rows[i].Cells[Cbox_File1SiriusStockCode.SelectedIndex].Style.BackColor = Color.PaleVioletRed;
                //        }
                //    }
                //}

                //for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                //{
                //    string val = dataGridView1.Rows[i].Cells[Cbox_File1Refference.SelectedIndex].Value.ToString();
                //    if (val.Contains(item))
                //    {
                //        if (dataGridView1.Rows[i].Cells[Cbox_File1Refference.SelectedIndex].Style.BackColor != Color.PaleVioletRed)
                //            dataGridView1.Rows[i].Cells[Cbox_File1Refference.SelectedIndex].Style.BackColor = Color.PaleGreen;

                //    }
                //}
            }

            foreach (var item in difList1)
            {
                barListItem1.Strings.Add(item.Refference);
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    string val = dataGridView1.Rows[i].Cells[Cbox_File1Refference.SelectedIndex].Value.ToString();
                    if (val.Contains(','))
                    {
                        string[] subs = val.Split(',');
                        foreach (string s in subs)
                        {
                            if (s.Contains('-'))
                            {
                                string[] referencestartend = s.Split('-');

                                string[] startarr = Regex.Split(referencestartend[0], @"\D+");
                                string[] endarr = Regex.Split(referencestartend[1], @"\D+");

                                string letter = new String(referencestartend[0].Where(Char.IsLetter).ToArray());

                                int start = Convert.ToInt32(startarr[1]);
                                int end = Convert.ToInt32(endarr[1]);
                                for (int inc = start; inc <= end; inc++)
                                {
                                    string newrefference = letter + inc.ToString();
                                    if (newrefference == item.Refference)
                                    {
                                        dataGridView1.Rows[i].Cells[Cbox_File1Refference.SelectedIndex].Style.BackColor = Color.PaleVioletRed;
                                    }
                                }
                            }
                            else
                            {
                                if (s == item.Refference)
                                {
                                    dataGridView1.Rows[i].Cells[Cbox_File1Refference.SelectedIndex].Style.BackColor = Color.PaleVioletRed;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (val == item.Refference)
                        {
                            dataGridView1.Rows[i].Cells[Cbox_File1Refference.SelectedIndex].Style.BackColor = Color.PaleVioletRed;
                        }
                    }
                }



                //for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                //{
                //    string val = dataGridView1.Rows[i].Cells[Cbox_File1Refference.SelectedIndex].Value.ToString(); 
                //    if (val.Contains(item.Refference))
                //    {
                //        dataGridView1.Rows[i].Cells[Cbox_File1Refference.SelectedIndex].Style.BackColor = Color.PaleVioletRed;
                //    }
                //}
            }

            foreach (var item in difList)
            {
                barListItem2.Strings.Add(item.Refference);
                for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                {
                    string val = dataGridView2.Rows[i].Cells[Cbox_File2Refference.SelectedIndex].Value.ToString();
                    if (val.Contains(','))
                    {
                        string[] subs = val.Split(',');
                        foreach (string s in subs)
                        {
                            if (s.Contains('-'))
                            {


                                string[] referencestartend = s.Split('-');

                                string[] startarr = Regex.Split(referencestartend[0], @"\D+");
                                string[] endarr = Regex.Split(referencestartend[1], @"\D+");

                                string letter = new String(referencestartend[0].Where(Char.IsLetter).ToArray());

                                int start = Convert.ToInt32(startarr[1]);
                                int end = Convert.ToInt32(endarr[1]);
                                for (int inc = start; inc <= end; inc++)
                                {
                                    string newrefference = letter + inc.ToString();
                                    if (newrefference == item.Refference)
                                    {
                                        dataGridView2.Rows[i].Cells[Cbox_File2Refference.SelectedIndex].Style.BackColor = Color.PaleVioletRed;
                                    }
                                }
                            }
                            else
                            {
                                if (s == item.Refference)
                                {
                                    dataGridView2.Rows[i].Cells[Cbox_File2Refference.SelectedIndex].Style.BackColor = Color.PaleVioletRed;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (val == item.Refference)
                        {
                            dataGridView2.Rows[i].Cells[Cbox_File2Refference.SelectedIndex].Style.BackColor = Color.PaleVioletRed;
                        }
                    }
                }
            }
        }

        private void button1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void button1_DragDrop(object sender, DragEventArgs e)
        {
            XlsxFile1.component.Clear();
            Cbox_File1CustomerStockCode.Properties.Items.Clear();
            Cbox_File1SiriusStockCode.Properties.Items.Clear();
            Cbox_File1StockName.Properties.Items.Clear();
            Cbox_File1PartNo.Properties.Items.Clear();
            Cbox_File1Refference.Properties.Items.Clear();
            var data = e.Data.GetData(DataFormats.FileDrop);
            if (data != null)
            {
                var filenames = data as string[];
                if (filenames.Length > 0)
                {
                    OpenXslxForProductionFile(filenames[0], gridControl1, gridView1);
                }
            }
        }

        private void button2_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void button2_DragDrop(object sender, DragEventArgs e)
        {
            XlsxFile2.component.Clear();
            Cbox_File2CustomerStockCode.Properties.Items.Clear();
            Cbox_File2SiriusStockCode.Properties.Items.Clear();
            Cbox_File2StockName.Properties.Items.Clear();
            Cbox_File2PartNo.Properties.Items.Clear();
            Cbox_File2Refference.Properties.Items.Clear();
            var data = e.Data.GetData(DataFormats.FileDrop);
            if (data != null)
            {
                var filenames = data as string[];
                if (filenames.Length > 0)
                {
                    OpenXslxForDesignFile(filenames[0], gridControl2, gridView2);
                    
                }
            }
        }

        private void checkEdit1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit1.Checked)
            {
                ribbonPageGroup1.Visible = true;
            }
            else
            {
                ribbonPageGroup1.Visible = false;
            }
        }

        private void checkEdit2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkEdit2.Checked)
            {
                ribbonPageGroup3.Visible = true;
            }
            else
            {
                ribbonPageGroup3.Visible = false;
            }
        }
    }
}