﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Deployment.Application;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ribbon_xnk_sunzex
{
    public partial class xnk_ribbon 
    {
        string sTemplatePath = "";
        string sOutputPath = "";
        private readonly string KEY_NAME_SUNZEX_templatePath = "SUNZEX_XNK_TEMPLATE_PATH";
        private readonly string KEY_NAME_SUNZEX_outputPath = "SUNZEX_XNK_OUTPUT_PATH";
        private readonly string KEY_NAME_SUNZEX_openAfter = "SUNZEX_XNK_OPEN_AFTER";
        private readonly string KEY_NAME_SUNZEX_outSeparate = "SUNZEX_XNK_OUT_SEPARATE";
        private readonly string KEY_NAME_SUNZEX_PKL = "SUNZEX_XNK_PKL";
        private readonly string KEY_NAME_SUNZEX_COMPANY = "SUNZEX_XNK_COMPANY";

        string mCompanyName = "";
        CommonOpenFileDialog mDialog;
        string[] tisuSheetType = { "70", "80", "100", "150", "200" };
        String[] tisu_sheet_name = { "sheets", "trang:", "trang" };
        String[] tisu_sheet_alt = { " sheets", "trang: ", " trang" };

        string[] tisuPOPrototype = { "P.O", "PO #", "PO:", "PO#" };

        string test = "";
        readonly bool DEBUG_PKL = false;
        readonly bool DEBUG_ALL = false;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            mDialog = new CommonOpenFileDialog
            {
                EnsurePathExists = true,
                EnsureFileExists = false,
                AllowNonFileSystemItems = false,
                DefaultFileName = "Chọn thư mục",
                Title = "Chọn thư mục"
            };
            mDialog.Filters.Add(new CommonFileDialogFilter("Excel Worksheets ", "xlsx,xls"));
            mDialog.ShowHiddenItems = true;

            sTemplatePath = ReadFromRegistry(KEY_NAME_SUNZEX_templatePath, @"C:\xnk_tisu");
            editBox_hoadonmau.Text = sTemplatePath;

            sOutputPath = ReadFromRegistry(KEY_NAME_SUNZEX_outputPath, sTemplatePath);
            editBox_thumucxuat.Text = sOutputPath;

            checkBox_PKL.Checked = "1".Equals(ReadFromRegistry(KEY_NAME_SUNZEX_PKL, "1"));
            checkBox_openAfter.Checked = "1".Equals(ReadFromRegistry(KEY_NAME_SUNZEX_openAfter, "1"));
            checkBox_xuatRieng.Checked = "1".Equals(ReadFromRegistry(KEY_NAME_SUNZEX_outSeparate, "0"));
            combobox1.Text = "Tisu".Equals(ReadFromRegistry(KEY_NAME_SUNZEX_COMPANY, "Tisu")) ? "Tisu" : "Sunzex";
            mCompanyName = combobox1.Text;
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            mDialog.InitialDirectory = sTemplatePath;
            if (mDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                editBox_hoadonmau.Text = Directory.Exists(mDialog.FileName) ? mDialog.FileName : Path.GetDirectoryName(mDialog.FileName);
                StoreInRegistry(KEY_NAME_SUNZEX_templatePath, editBox_hoadonmau.Text);
                sTemplatePath = editBox_hoadonmau.Text;
            }
        }

        public void StoreInRegistry(string keyName, string value)
        {
            RegistryKey rootKey = Registry.CurrentUser;
            string registryPath = @"Software\Sunzex\XNK_ribbon";
            using (RegistryKey rk = rootKey.CreateSubKey(registryPath))
            {
                rk.SetValue(keyName, value, RegistryValueKind.String);
            }
        }
        public string ReadFromRegistry(string keyName, string defaultValue)
        {
            RegistryKey rootKey = Registry.CurrentUser;
            string registryPath = @"Software\Sunzex\XNK_ribbon";
            using (RegistryKey rk = rootKey.OpenSubKey(registryPath, false))
            {
                if (rk == null)
                {
                    return defaultValue;
                }
                var res = rk.GetValue(keyName, defaultValue);
                if (res == null)
                {
                    return defaultValue;
                }
                return res.ToString();
            }
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            mDialog.InitialDirectory = sOutputPath;
            if (mDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                editBox_thumucxuat.Text = Directory.Exists(mDialog.FileName) ? mDialog.FileName : Path.GetDirectoryName(mDialog.FileName);
                StoreInRegistry(KEY_NAME_SUNZEX_outputPath, editBox_thumucxuat.Text);
                sOutputPath = editBox_thumucxuat.Text;
            }
        }

        private void button_shipping_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = (Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = wb.ActiveSheet;

            //Check
            if (!"Số tờ khai".Equals(worksheet.Range["C4"].Value2) || !worksheet.Range["E2"].Value2.Contains("Tờ khai"))
            {
                MessageBox.Show("Hãy mở tờ khai để tạo shipping");
                return;
            }
            ShippingModel ship = new ShippingModel();
            string a;
            a = worksheet.Range["F8"].Value2;
            ship.date = a.Substring(0, 10);
            a = worksheet.Range["I46"].Value2;
            ship.shipment = a.Substring(3, 3) + a.Substring(0, 3) + a.Substring(6, 4);
            a = worksheet.Range["L6"].Value2;
            ship.transport = "BY " + (("" + a[a.Length - 1]).Equals("3") || ("" + a[a.Length - 1]).Equals("2") ? "SEA" : ("" + a[a.Length - 1]).Equals("4") ? "TRUCK" : ("" + a[a.Length - 1]).Equals("1") ? "AIR" : "OTHER");
            ship.destination = worksheet.Range["M43"].Value2;
            ship.portload = worksheet.Range["M44"].Value2;
            ship.voyage = worksheet.Range["M45"].Value2;
            ship.comodity = "AS PER INVOICE NO: " + worksheet.Range["R49"].Value2;
            ship.amount = worksheet.Range["U53"].Value2;

            ship.total = worksheet.Range["H40"].Value2;
            if (!"CT".Equals(worksheet.Range["M40"].Value2))
            {
                a = worksheet.Range["H47"].Value2;
                if (a.LastIndexOf("=") < a.LastIndexOf(")") && a.LastIndexOf("C") < a.LastIndexOf("T") && a.LastIndexOf("=") < a.LastIndexOf("C") && a.LastIndexOf("T") < a.LastIndexOf(")"))
                {
                    ship.total = a.Substring(a.LastIndexOf("=") + 1, a.LastIndexOf("C") - a.LastIndexOf("=") - 1);
                }
            }

            ship.gross = worksheet.Range["H41"].Value2;

            //a = ship.date.Substring(3, 2) + "." + ship.date.Substring(0, 2);
            int d1 = int.Parse(ship.date.Substring(3, 2));
            int d2 = int.Parse(ship.date.Substring(0, 2));
            ship.sheetname = d1 + "." + d2;

            if (checkBox_xuatRieng.Checked)
            {
                ship.filename = mCompanyName.ToUpper() + @"_SHIPING." + worksheet.Range["R49"].Value2.Replace("/", "");
                //ship.sheetname = worksheet.Range["R49"].Value2.Replace("/", "");
            }
            else
            {
                a = ship.comodity.Substring(ship.comodity.IndexOf(@":") + 6, 2) + "." + ship.comodity.Substring(ship.comodity.IndexOf(@":") + 8, 3);
                //ship.sheetname = a.Replace("/", "");

                ship.filename = mCompanyName.ToUpper() + @"_SHIPING." + ship.date.Substring(6, 4) + ".T" + ship.date.Substring(3, 2);
            }

            Copy_shipping_work_sheet(ship);
        }
        private void Copy_shipping_work_sheet(ShippingModel ship)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                string templatePath = sTemplatePath + @"\"+ mCompanyName + "_SHIP.xlsx";
                string outputPath = sOutputPath + @"\" + ship.filename + ".xlsx";

                app.Visible = false;
                app.Workbooks.Add();
                if (!System.IO.File.Exists(templatePath))
                {
                    MessageBox.Show("Không tìm thấy file mẫu: "+ mCompanyName + " : "+ templatePath);
                    app.Workbooks.Close();
                    app.Quit();
                    return;
                }
                app.Workbooks.Add(templatePath);
                try
                {
                    app.Workbooks.Add(!System.IO.File.Exists(outputPath) ? "" : outputPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                _Worksheet ws = (_Worksheet)app.Workbooks[2].Worksheets[1];
                _Worksheet sheet = (_Worksheet)app.Workbooks[3].Worksheets[1];

                //check exist
                bool found = false;
                
                if (IsOpenedWB_ByName(ship.filename + ".xlsx"))
                {
                    throw new IOException("File " + ship.filename + " đang được mở nên không thể ghi đè.  Hãy đóng file và thử tạo lại shipping! ");
                }

                //Begin copy
                ws.Copy(sheet);
                app.Workbooks[3].Sheets[1].Activate();
                sheet = app.ActiveSheet;
                sheet.Range["H7"].Value2 = ship.date;
                sheet.Range["E14"].Value2 = ship.shipment;
                sheet.Range["E15"].Value2 = ship.transport;
                sheet.Range["E16"].Value2 = ship.destination;
                sheet.Range["E17"].Value2 = ship.portload;
                sheet.Range["E18"].Value2 = ship.voyage;
                sheet.Range["E19"].Value2 = ship.comodity;
                sheet.Range["E20"].Value2 = ship.amount.Replace(".", "").Replace(",", ".");
                sheet.Range["E22"].Value2 = ship.total.Replace(".", "");
                sheet.Range["E23"].Value2 = ship.gross.Replace(".", "").Replace(",", ".");

                while (isExistingWorkbook(app.Workbooks[3].Sheets, ship.sheetname))
                {
                    ship.sheetname = makeSheetName(ship.sheetname);
                }
                sheet.Name = ship.sheetname;

                //TODO: Handle saves option "No"/"Cancel"
                app.ActiveWorkbook.SaveAs(outputPath);

                object misValue = System.Reflection.Missing.Value;
                app.Workbooks[3].Close(false, misValue, misValue);
                app.Workbooks[2].Close(false, misValue, misValue);
                app.Workbooks[1].Close(false, misValue, misValue);
                app.Workbooks.Close();
                app.Quit();
                if (checkBox_openAfter.Checked)
                {
                    OpenFile(outputPath);
                }
                else MessageBox.Show("Shipping " + ship.sheetname + " tạo thành công!");
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi đọc/ghi/mở file: Do " + ex.ToString());
                app.Workbooks.Close();
                app.Quit();
            }
            finally
            {
                app.Workbooks.Close();
                app.Quit();
            }
        }
        private void OpenFile(string path)
        {
            Process fileopener = new Process();
            fileopener.StartInfo.FileName = "explorer";
            //fileopener.StartInfo.Arguments = @"/r ";
            fileopener.StartInfo.Arguments += ("\"" + path + "\"");
            fileopener.Start();
        }
        public class InvoiceModel
        {
            public int type, pkl_num;
            public string number, date, consignee, portload, destination, sailing, test, sheetname, filename;
            public string[] detail_name, detail_order, detail_PO, detail_quantity, detail_price, detail_size, detail_sheets_num;
            public InvoiceModel(int type)
            {
                this.type = type;
                detail_name = new string[type + 1];
                detail_order = new string[type + 1];
                detail_PO = new string[type + 1];
                detail_quantity = new string[type + 1];
                detail_price = new string[type + 1];
                detail_size = new string[type + 1];
                detail_sheets_num = new string[type + 1];
            }
        }
        public class ShippingModel
        {
            public string date, shipment, transport, destination, portload, voyage, comodity, amount, total, gross, filename, sheetname;
            public ShippingModel() { }
        }

        //---------------------  WBHelper   ---------------------
        //-------------------------------------------------------
        public static bool IsOpenedWB_ByName(string wbName)
        {
            return (GetOpenedWB_ByName(wbName) != null);
        }

        public static bool IsOpenedWB_ByPath(string wbPath)
        {
            return (GetOpenedWB_ByPath(wbPath) != null);
        }

        public static Workbook GetOpenedWB_ByName(string wbName)
        {
            return (Workbook)GetRunningObjects().FirstOrDefault(x => (System.IO.Path.GetFileName(x.Path) == wbName) && (x.Obj is Workbook)).Obj;
        }

        public static Workbook GetOpenedWB_ByPath(string wbPath)
        {
            return (Workbook)GetRunningObjects().FirstOrDefault(x => (x.Path == wbPath) && (x.Obj is Workbook)).Obj;
        }

        public static List<RunningObject> GetRunningObjects()
        {
            // Get the table.
            List<RunningObject> roList = new List<RunningObject>();
            IBindCtx bc;
            CreateBindCtx(0, out bc);
            IRunningObjectTable runningObjectTable;
            bc.GetRunningObjectTable(out runningObjectTable);
            IEnumMoniker monikerEnumerator;
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            // Enumerate and fill list
            IMoniker[] monikers = new IMoniker[1];
            IntPtr numFetched = IntPtr.Zero;
            List<object> names = new List<object>();
            List<object> books = new List<object>();
            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                RunningObject running;
                monikers[0].GetDisplayName(bc, null, out running.Path);
                runningObjectTable.GetObject(monikers[0], out running.Obj);
                roList.Add(running);
            }
            return roList;
        }

        public struct RunningObject
        {
            public string Path;
            public object Obj;
        }

        [System.Runtime.InteropServices.DllImport("ole32.dll")]
        static extern void CreateBindCtx(int a, out IBindCtx b);
        //-------------------------------------------------------
        //-------------------End of WBHelper---------------------

        private void CheckBox3_Click(object sender, RibbonControlEventArgs e)
        {
            StoreInRegistry(KEY_NAME_SUNZEX_openAfter, checkBox_openAfter.Checked ? "1" : "0");
        }

        private void checkBox_xuatRieng_Click(object sender, RibbonControlEventArgs e)
        {
            StoreInRegistry(KEY_NAME_SUNZEX_outSeparate, checkBox_xuatRieng.Checked ? "1" : "0");
        }
        private void checkBox_PKL_Click(object sender, RibbonControlEventArgs e)
        {
            StoreInRegistry(KEY_NAME_SUNZEX_PKL, checkBox_PKL.Checked ? "1" : "0");
        }

        private void button_invoice_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = (Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = wb.ActiveSheet;

            //Check
            if (!"Số tờ khai".Equals(worksheet.Range["C4"].Value2) || !worksheet.Range["E2"].Value2.Contains("Tờ khai"))
            {
                MessageBox.Show("Hãy mở tờ khai để tạo invoice!");
                return;
            }

            int type = 1;
            Range[] items = new Range[10];
            Range invoiceItemFindRange = worksheet.Cells;
            for (int i = 1; i < 10; i++)
            {
                items[i] = invoiceItemFindRange.Find("<" + i.ToString("00") + ">");
                if (items[i] is null)
                {
                    break;
                }
                type = i;
            }

            InvoiceModel invoice = new InvoiceModel(type);

            string a;
            invoice.number = "No: " + worksheet.Range["R49"].Value2;
            a = worksheet.Range["F8"].Value2;
            invoice.date = a.Substring(0, 10);
            a = worksheet.Range["H47"].Value2;
            invoice.consignee = a.Substring(0, a.IndexOf("TONZEX")).Replace("GIAO HANG CHO ", "").Replace(" THEO CHI DINH CUA", "").Replace("G/H CHO ", "");
            invoice.portload = worksheet.Range["M44"].Value2;
            invoice.destination = worksheet.Range["M43"].Value2;
            a = worksheet.Range["I46"].Value2;
            invoice.sailing = a.Substring(3, 3) + a.Substring(0, 3) + a.Substring(6, 4);

            a = worksheet.Range["R49"].Value2;
            a = (a.IndexOf("/") == -1 ? a : a.Substring(0, a.IndexOf("/")) + "-" + a.Substring(a.LastIndexOf("/") + 1, a.Length - a.LastIndexOf("/") - 1));

            int d1 = int.Parse(invoice.date.Substring(3, 2));
            int d2 = int.Parse(invoice.date.Substring(0, 2));
            invoice.sheetname = d1 + "." + d2;
            //invoice.sheetname = invoice.date.Substring(3, 2) + "." + a.Substring(6, a.Length - 6);

            if (checkBox_xuatRieng.Checked)
            {
                invoice.filename = mCompanyName.ToUpper() + @"_" + a;
            }
            else
            {
                invoice.filename = mCompanyName.ToUpper() + @"_INVOICE." + invoice.date.Substring(6, 4) + ".T" + invoice.date.Substring(3, 2);
            }

            try
            {
                for (int i = 1; i <= type; i++)
                {
                    a = worksheet.Range["F" + (items[i].Row + 3)].Value2;
                    invoice.detail_name[i] = isTisuCom() ? (a.ToUpper().Contains("NOTE") ? "NOTEBOOK" : "COMPOSITION BOOK") : (a.Contains("giấy") ? "PAPER FOLDER" : "PP FOLDER");
                    invoice.detail_order[i] = a.Substring(1, a.IndexOf("#&")-1);
                    if (isTisuCom()) {
                        foreach (string num in tisuSheetType)
                            if (a.Contains(num) && a.Contains(@" tờ")) invoice.detail_sheets_num[i] = num;
                    }
                    a = worksheet.Range["F" + (items[i].Row + 3)].Value2;
                    a = isTisuCom() ? a.Substring(a.LastIndexOf("(") + 1, a.LastIndexOf(")") - a.LastIndexOf("(") - 1) : a.Substring(a.IndexOf("(") + 1, a.IndexOf(")") - a.IndexOf("(") - 1);
                    invoice.detail_size[i] = a.ToLower().Replace(" ", "").Replace("x", "-").Replace("*", "-").Replace("c", "").Replace("m", "").Replace(",", ".");
                    invoice.detail_quantity[i] = worksheet.Range["Q" + (items[i].Row + 6)].Value2.Replace(".", "").Replace(",", ".");
                    invoice.detail_price[i] = worksheet.Range["R" + (items[i].Row + 8)].Value2.Replace(".", "").Replace(",", ".");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (checkBox_PKL.Checked) {
                Microsoft.Office.Interop.Excel.Application pklapp = Globals.ThisAddIn.Application;
                try
                {
                    CommonOpenFileDialog aDialog = new CommonOpenFileDialog
                    {
                        EnsurePathExists = true,
                        EnsureFileExists = true,
                        AllowNonFileSystemItems = false,
                        DefaultFileName = "Mở packing list",
                        Title = "Tự đánh số PO cho invoice từ packing list"
                    };
                    aDialog.DefaultDirectory = sOutputPath;
                    aDialog.Filters.Add(new CommonFileDialogFilter("Excel Worksheets", "xlsx,xls"));
                    aDialog.ShowHiddenItems = true;

                    if (aDialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        pklapp = new Microsoft.Office.Interop.Excel.Application();
                        pklapp.Visible = false;
                        pklapp.Workbooks.Add(aDialog.FileName);
                        _Worksheet sheet = (_Worksheet)pklapp.Workbooks[1].Sheets[1];
                        sheet.Activate();
                        _Worksheet sheet1 = (_Worksheet)pklapp.Workbooks[1].ActiveSheet;

                        if (isTisuCom())
                        {
                            Range workSheetCells = sheet.Cells;
                            Range aPOItem = null, aSizeItem, aTypeItem = null;
                            Range firstResultPO = null, firstResultSize, firstResultType = null;
                            SortedList<int, string> existPOProtos = new SortedList< int, string> ();
                            for (int i = 0; i < tisuPOPrototype.Length; i++)
                            {
                                aPOItem = workSheetCells.Find(tisuPOPrototype[i]);
                                if (aPOItem != null && aPOItem.Value2.Contains(tisuPOPrototype[i])) {
                                    existPOProtos.Add(aPOItem.Row, aPOItem.Value2);
                                    firstResultPO = aPOItem;
                                    while (aPOItem != null) {
                                        aPOItem = findAfterString(tisuPOPrototype[i], workSheetCells, aPOItem);
                                        if (aPOItem != null && aPOItem.Address != firstResultPO.Address) {
                                            existPOProtos.Add(aPOItem.Row, aPOItem.Value2);
                                        } else {
                                            break;
                                        }
                                    }
                                }
                            }
                            test = "PO count="+existPOProtos.Count+"\r\n";
                            foreach (var kvp in existPOProtos) {
                                test += " [" +kvp.Key+":"+kvp.Value+ "]";
                            }
                            if (DEBUG_PKL) MessageBox.Show(test);

                            //make Size stored value table
                            string size = ""; string sheetType = "";
                            SortedList<int, string> existSizeProtos = new SortedList<int, string>();
                            aSizeItem = workSheetCells.Find("cm");
                            firstResultSize = aSizeItem;
                            if (aSizeItem != null && aSizeItem.Value2.Contains("cm"))
                            {
                                firstResultSize = aSizeItem;
                                while (aSizeItem != null)
                                {
                                    //if (DEBUG_PKL) MessageBox.Show(aSizeItem.Address + ":" + aSizeItem.Value2);
                                    size = getTisuBookSize(aSizeItem.Value2);
                                    if (!size.Equals(WRONG_MEASURE)) {
                                        existSizeProtos.Add(aSizeItem.Row, size);
                                    }
                                    aSizeItem = findAfterString("cm", workSheetCells, aSizeItem);
                                    if (aSizeItem == null || aSizeItem.Address == firstResultSize.Address) {
                                        break;
                                    }
                                }
                            }
                            test = "Size count=" + existSizeProtos.Count + "\r\n";
                            foreach (var kvp in existSizeProtos)
                            {
                                test += " [" + kvp.Key + ":" + kvp.Value + "]";
                            }
                            if (DEBUG_PKL) MessageBox.Show(test);

                            //make sheet number stored value table
                            SortedList<int, string> existSheetNumProtos = new SortedList<int, string>();
                            for (int j = 0; j < tisu_sheet_name.Length; j++)
                            {
                                aTypeItem = workSheetCells.Find(tisu_sheet_name[j]);
                                if (aTypeItem != null && aTypeItem.Value2.Contains(tisu_sheet_name[j]))
                                {
                                    firstResultType = aTypeItem;
                                    while (aTypeItem != null) {
                                        sheetType = getTisuBookType(aTypeItem.Value2, j);
                                        if (!sheetType.Equals(WRONG_MEASURE) && aTypeItem.Value2.Contains(tisu_sheet_name[j]))
                                        {
                                            if (!existSheetNumProtos.ContainsKey(aTypeItem.Row))
                                            existSheetNumProtos.Add(aTypeItem.Row, sheetType);
                                        }
                                        aTypeItem = findAfterString(tisu_sheet_name[j], workSheetCells, aTypeItem);
                                        if (aTypeItem == null || aTypeItem.Address == firstResultType.Address)
                                        {
                                            break;
                                        }
                                    }
                                }
                            }
                            test = "Sheet Type count=" + existSheetNumProtos.Count + "\r\n";
                            foreach (var kvp in existSheetNumProtos) {
                                test += " [" + kvp.Key + ":" + kvp.Value + "]";
                            }
                            if (DEBUG_PKL) MessageBox.Show(test);
                            int max = 0;
                            while (max < 30 && max < existPOProtos.Count && max < existSizeProtos.Count && max < existSheetNumProtos.Count)
                            {
                                if (DEBUG_PKL) MessageBox.Show("LOOP while: " + max);
                                size = existSizeProtos.Values[max];
                                sheetType = existSheetNumProtos.Values[max];
                                string b;
                                string size1 = size!= null ? size.Split('-')[1]+"-"+size.Split('-')[0] : "";
                                for (int i = 1; i <= type; i++) {
                                    if ((size1.Equals(invoice.detail_size[i]) || size.Equals(invoice.detail_size[i])) && sheetType.Equals(invoice.detail_sheets_num[i]))
                                    {
                                        if (DEBUG_PKL) MessageBox.Show("COMPARE: size=" + size + " sheetType=" + sheetType + "\r\nCOMAPRE WITH: i=" + i + " size=" + invoice.detail_size[i] + @" sheetType=" + invoice.detail_sheets_num[i] + "\r\n COMPARE: size1=" + size1 + " sheetType=" + sheetType+"\r\n" + "\r\n\r\n aPOItem Address: " + existPOProtos.Keys[max] + "\r\n aSizeItem Address ROW: " + existSizeProtos.Keys[max] + "\r\n aTypeItem Address: " + aTypeItem.Address);
                                        b = getTisuPO(existPOProtos.Values[max]);
                                        invoice.detail_PO[i] = getShortPOString(invoice.detail_PO[i], b);
                                        break;
                                    }
                                }

                                if (DEBUG_PKL) MessageBox.Show(" === max=" + max + " existPOProtos=" + existPOProtos.Values[max]+ " size:"+existSizeProtos.Values[max] + (aTypeItem == null ? "aTypeItem=null" : ""));
                                max++;
                            }
                            if (DEBUG_ALL || DEBUG_PKL) 
                                MessageBox.Show(" === OUT and max=" + max + " existPOProtos=" + existPOProtos.Count + " size:" + existSizeProtos.Count + "existSheetNumProtos=" + existSheetNumProtos.Count);
                        }
                        else
                        {
                            Range POitems = sheet.Cells;
                            Range aPOitems;
                            aPOitems = POitems.Find("Folder size");
                            Range firstResult = aPOitems;
                            int max = 0;
                            while (!(aPOitems is null) && max++ < 20)
                            {
                                a = sheet.Range["A" + aPOitems.Row].Value2;
                                a = a.Substring(a.IndexOf("(") + 1, a.IndexOf(")") - a.IndexOf("(") - 1);
                                a = a.ToLower().Replace(" ", "");
                                a = a.Replace("x", "-").Replace("cm", "");
                                a = a.Replace("*", "-").Replace(",", ".");
                                string b;
                                for (int i = 1; i <= type; i++)
                                {
                                    if (a.Equals(invoice.detail_size[i]))
                                    {
                                        b = sheet.Range["E" + (aPOitems.Row + 3)].Value2.Replace("PO", "").Replace(" ", "");
                                        b = b.Replace(":", "").Replace("#", "");
                                        invoice.detail_PO[i] = getShortPOString(invoice.detail_PO[i], b);
                                    }
                                }

                                aPOitems = POitems.FindNext(aPOitems);
                                if (aPOitems.Address == firstResult.Address) aPOitems = null;
                            }
                        }
                        object misValue = System.Reflection.Missing.Value;
                        pklapp.Workbooks[1].Close(false, misValue, misValue);
                        pklapp.Quit();
                    }
                }
                catch (System.Runtime.InteropServices.COMException ex) {
                    MessageBox.Show(ex.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(@"Lỗi khi đọc packing list: " + ex.ToString());
                    object misValue = System.Reflection.Missing.Value;
                    pklapp.Workbooks[1].Close(false, misValue, misValue);
                    pklapp.Workbooks.Close();
                    pklapp.Quit();
                }
            }

            Copy_invoice_work_sheet(invoice);
        }
        private void Copy_invoice_work_sheet(InvoiceModel invoice)
        {
            object misValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                string templatePath = sTemplatePath + @"\" + mCompanyName + @"_INV.xlsx";
                string outputPath = sOutputPath + @"\" + invoice.filename + ".xlsx";

                app.Visible = false;
                app.Workbooks.Add();
                if (!System.IO.File.Exists(templatePath))
                {
                    MessageBox.Show("Không tìm thấy file mẫu: " + mCompanyName + " : " + templatePath);
                    app.Workbooks.Close();
                    app.Quit();
                    return;
                }
                app.Workbooks.Add(templatePath);
                try
                {
                    app.Workbooks.Add(!System.IO.File.Exists(outputPath) ? "" : outputPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                _Worksheet ws = (_Worksheet)app.Workbooks[2].Worksheets[invoice.type];
                _Worksheet sheet = (_Worksheet)app.Workbooks[3].Worksheets[1];

                //check exist
                if (IsOpenedWB_ByName(invoice.filename + ".xlsx"))
                {
                    app.Workbooks[3].Close(false, misValue, misValue);
                    app.Workbooks[2].Close(false, misValue, misValue);
                    throw new IOException("File " + invoice.filename + " đang được mở nên không thể ghi đè.  Hãy đóng file và thử tạo lại invoice! ");
                }

                //Begin copy
                ws.Copy(sheet);
                app.Workbooks[3].Sheets[1].Activate();
                sheet = app.ActiveSheet;

                sheet.Range["E7"].Value2 = invoice.number;
                if (isTisuCom()) {
                    sheet.Range["H7"].Value2 = "Date: " + invoice.date;
                } else { 
                    sheet.Range["I7"].Value2 = invoice.date;
                }
                sheet.Range["A10"].Value2 = invoice.consignee;
                sheet.Range["A17"].Value2 = invoice.portload;
                sheet.Range["C17"].Value2 = invoice.destination;
                sheet.Range["C19"].Value2 = invoice.sailing;

                while (isExistingWorkbook(app.Workbooks[3].Sheets, invoice.sheetname)) {
                    invoice.sheetname = makeSheetName(invoice.sheetname);
                }
                sheet.Name = invoice.sheetname;

                if (invoice.type <= 3)
                {
                    int row1 = 19;
                    for (int i = 1; i <= invoice.type; i++)
                    {
                        sheet.Range["A" + (row1 + i * 3)].Value2 = invoice.detail_name[i];
                        sheet.Range["A" + (row1 + i * 3 + 1)].Value2 = @"ORDER: HSS" + (Int32.Parse(invoice.detail_order[i]) + 9000);
                        sheet.Range["A" + (row1 + i * 3 + 2)].Value2 = @"PO#" + invoice.detail_PO[i];
                        sheet.Range["E" + (row1 + i * 3)].Value2 = invoice.detail_quantity[i];
                        sheet.Range["G" + (row1 + i * 3)].Value2 = isTisuCom() ? "0.16":"0.03";
                    }
                    int row2 = row1 + 4 + 3 * invoice.type;
                    for (int i = 1; i <= invoice.type; i++)
                    {
                        sheet.Range["A" + (row2 + i * 3)].Value2 = invoice.detail_name[i];
                        sheet.Range["A" + (row2 + i * 3 + 1)].Value2 = @"ORDER: HSS" + (Int32.Parse(invoice.detail_order[i]) + 9000);
                        sheet.Range["A" + (row2 + i * 3 + 2)].Value2 = @"PO#" + invoice.detail_PO[i];
                        sheet.Range["E" + (row2 + i * 3)].Value2 = invoice.detail_quantity[i];
                        sheet.Range["G" + (row2 + i * 3)].Value2 = invoice.detail_price[i];
                    }
                }
                else if (invoice.type <= 5)
                {
                    int row1 = 20;
                    for (int i = 1; i <= invoice.type; i++)
                    {
                        sheet.Range["A" + (row1 + i * 2)].Value2 = invoice.detail_name[i] + " - " + @"ORDER: HSS" + (Int32.Parse(invoice.detail_order[i]) + 9000);
                        sheet.Range["A" + (row1 + i * 2 + 1)].Value2 = @"PO#" + invoice.detail_PO[i];
                        sheet.Range["E" + (row1 + i * 2)].Value2 = invoice.detail_quantity[i];
                        sheet.Range["G" + (row1 + i * 2)].Value2 = isTisuCom() ? "0.16" : "0.03";
                    }
                    int row2 = row1 + 4 + 2 * invoice.type;
                    for (int i = 1; i <= invoice.type; i++)
                    {
                        sheet.Range["A" + (row2 + i * 2)].Value2 = invoice.detail_name[i] + " - " + @"ORDER: HSS" + (Int32.Parse(invoice.detail_order[i]) + 9000);
                        sheet.Range["A" + (row2 + i * 2 + 1)].Value2 = @"PO#" + invoice.detail_PO[i];
                        sheet.Range["E" + (row2 + i * 2)].Value2 = invoice.detail_quantity[i];
                        sheet.Range["G" + (row2 + i * 2)].Value2 = invoice.detail_price[i];
                    }
                }

                app.ActiveWorkbook.SaveAs(outputPath);
                app.Workbooks[3].Close(false, misValue, misValue);
                app.Workbooks[2].Close(false, misValue, misValue);
                app.Workbooks[1].Close(false, misValue, misValue);
                app.Workbooks.Close();
                app.Quit();
                if (checkBox_openAfter.Checked)
                {
                    OpenFile(outputPath);
                }
                else MessageBox.Show("Invoice " + invoice.sheetname + " tạo thành công!");
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.ToString());
                app.Workbooks[1].Close(false, misValue, misValue);
                app.Workbooks.Close();
                app.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                app.Workbooks[1].Close(false, misValue, misValue);
                app.Workbooks.Close();
                app.Quit();
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Bộ công cụ XNK cty Sunzex. Phiên bản " + Version.Major.ToString() + "." + Version.Minor.ToString() + "." + Version.Build.ToString() + "." + Version.Revision.ToString() + " Liên hệ & tác giả: Bùi Khánh Duy Anh - email: shipping.anh@sunzex.com");
        }

        public Version Version
        {
            get
            {
                return ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
        }

        private void combobox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            StoreInRegistry(KEY_NAME_SUNZEX_COMPANY, combobox1.Text);
            mCompanyName = combobox1.Text;
        }

        private bool isTisuCom() {
            return mCompanyName.Equals("Tisu");
        }

        string getTisuPO(string tisuPOCellValue) {
            string rawData = tisuPOCellValue;
            while (rawData.Contains("  ")) rawData = rawData.Replace(@"  ", @" ");
            rawData = rawData.Replace(@"# ", @"#").Replace(@" #", @"#");
            rawData = rawData.Replace(@": ", @":").Replace(@" :", @":");
            rawData = rawData.Replace(@"P.O", @"PO").Replace(@"P/O", @"PO").Replace(@"P O", @"PO");
            if (rawData.Contains("PO"))
            {
                String[] spearator = { "PO" };
                String[] strlist = rawData.Split(spearator, StringSplitOptions.None);
                rawData = strlist[1].IndexOf(" ") != -1 ? strlist[1].Substring(0, strlist[1].IndexOf(" ")) : strlist[1];
            }
            rawData = rawData.Replace(@".", @"").Replace(@":", @"").Replace(@"#", @"");
            if (DEBUG_PKL) MessageBox.Show("getTisuPO:" + rawData + "\r\n" + tisuPOCellValue);
            return rawData;
        }

        readonly string WRONG_MEASURE = "WRONG_INPUT";
        string getTisuBookSize(string tisuSizeCellValue)
        {
            string rawData = tisuSizeCellValue.ToLower();
            String[] sepearator = { "\r\n", ",", ":" };
            String[] stringseperate;
            stringseperate = rawData.Split(sepearator, StringSplitOptions.RemoveEmptyEntries);
            foreach (string ak in stringseperate) if (ak.Contains("cm")) rawData = ak;
            rawData = rawData.Replace(" ", "").Replace("cm", "");
            rawData = rawData.Replace("x", "-").Replace("*", "-");

            string[] a = rawData.Split('-');
            if (a.Length == 2) {
                if (a[0].IndexOf(".") == -1) a[0] += ".0";
                if (a[1].IndexOf(".") == -1) a[1] += ".0";
                rawData = a[0] + "-" + a[1];
                //if (DEBUG_PKL) MessageBox.Show("get Book Size:"+rawData); 
                return rawData; 
            }

            //if (DEBUG_PKL) MessageBox.Show("get Book Size:" + WRONG_MEASURE+"\r\n"+ tisuSizeCellValue);
            return WRONG_MEASURE;
        }
        string getTisuBookType(string tisuTypeCellValue, int type_id)
        {
            string rawData = tisuTypeCellValue.ToLower().Replace("\r\n", " ");
            while (rawData.Contains("  ")) rawData = rawData.Replace("  ", " ");
            rawData.Replace(tisu_sheet_alt[type_id], tisu_sheet_name[type_id]);
            foreach (string bk in tisuSheetType) if (rawData.Contains(bk)) { 
                    //if (DEBUG_PKL) MessageBox.Show("getTisuBookType:" + bk + "\r\n" + tisuTypeCellValue + "\r\n" + type_id);  
                    return bk; 
                }

            //if (DEBUG_PKL) MessageBox.Show("getTisuBookType:" + WRONG_MEASURE + "\r\n" + tisuTypeCellValue + "\r\n" + tisu_sheet_name[type_id]);
            return WRONG_MEASURE;
        }

        string getShortPOString(string originStr, string newCome) {
            string strPO = originStr;
            if (strPO != null && strPO.Length > 0) {
                if (!strPO.Contains(newCome)) {
                    string c = newCome.Substring(0, newCome.Length - 5);
                    if (strPO.Contains(c)) {
                        strPO = strPO.Insert(strPO.IndexOf(c) + c.Length, (newCome.Substring(c.Length, 5) + "/"));
                    } else {
                        strPO += ("," + newCome);
                    }
                }
            } else {
                strPO = newCome;
            }
            //if (DEBUG_PKL) MessageBox.Show("getShortPOString:" + strPO);
            return strPO;
        }

        string getLastCellAddress(Range range) { 
            string lastColumn = ""+ ('A' + range.Columns.Count - 1);
            int lastRow = range.Rows.Count;
            return lastColumn+lastRow;
        }

        Range findAfterString(string val, Range rangeToFind, Object after)
        {
            return rangeToFind.Find(val, after, XlFindLookIn.xlValues, Type.Missing, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, false, Type.Missing);
        }
        Range findAfterCell(Range cell, Range rangeToFind, Object after) {
            return rangeToFind.Find(""+cell.Value2, after, XlFindLookIn.xlValues, Type.Missing, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, false, Type.Missing);
        }

        string makeSheetName(string nameToUp) {
            string a = nameToUp;
            if (!nameToUp.Contains("("))
            {
                a = nameToUp + @" (1)";
            }
            else {
                int i = int.Parse(nameToUp.Substring(nameToUp.IndexOf("(")+1, nameToUp.IndexOf(")")- nameToUp.IndexOf("(")-1));
                a = nameToUp.Substring(0, nameToUp.IndexOf("(") + 1) + (int)(i + 1) + (")");
            }
            return a;
        }

        bool isExistingWorkbook(Sheets Sheets, string name) {
            foreach (_Worksheet aSheet in Sheets)
            {
                if (name.Equals(aSheet.Name))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
