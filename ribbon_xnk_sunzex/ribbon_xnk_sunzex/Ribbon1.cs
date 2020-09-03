using System;
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
    public partial class Ribbon1
    {
        string sTemplatePath = "";
        string sOutputPath = "";
        private readonly string KEY_NAME_SUNZEX_templatePath = "SUNZEX_XNK_TEMPLATE_PATH";
        private readonly string KEY_NAME_SUNZEX_outputPath = "SUNZEX_XNK_OUTPUT_PATH";
        private readonly string KEY_NAME_SUNZEX_openAfter = "SUNZEX_XNK_OPEN_AFTER";
        private readonly string KEY_NAME_SUNZEX_outSeparate = "SUNZEX_XNK_OUT_SEPARATE";
        private readonly string KEY_NAME_SUNZEX_PKL = "SUNZEX_XNK_PKL";

        CommonOpenFileDialog mDialog;

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

            sTemplatePath = ReadFromRegistry(KEY_NAME_SUNZEX_templatePath, "");
            editBox_hoadonmau.Text = sTemplatePath;

            sOutputPath = ReadFromRegistry(KEY_NAME_SUNZEX_outputPath, sTemplatePath);
            editBox_thumucxuat.Text = sOutputPath;

            checkBox_PKL.Checked = "1".Equals(ReadFromRegistry(KEY_NAME_SUNZEX_PKL, "1"));

            checkBox_openAfter.Checked = "1".Equals(ReadFromRegistry(KEY_NAME_SUNZEX_openAfter, "1"));

            checkBox_xuatRieng.Checked = "1".Equals(ReadFromRegistry(KEY_NAME_SUNZEX_outSeparate, "0"));
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            mDialog.InitialDirectory = editBox_hoadonmau.Text;
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
            mDialog.InitialDirectory = editBox_thumucxuat.Text;
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

            if (checkBox_xuatRieng.Checked)
            {
                ship.filename = @"SUNZEX_SHIPING." + worksheet.Range["R49"].Value2.Replace("/", "");
                ship.sheetname = worksheet.Range["R49"].Value2.Replace("/", "");
            }
            else
            {
                a = ship.comodity.Substring(ship.comodity.IndexOf(@":") + 6, 2) + "." + ship.comodity.Substring(ship.comodity.IndexOf(@":") + 8, 3);
                ship.sheetname = a.Replace("/", "");

                ship.filename = @"SUNZEX_SHIPING." + ship.date.Substring(6, 4) + ".T" + ship.date.Substring(3, 2);
            }

            Copy_shipping_work_sheet(ship);
        }
        private void Copy_shipping_work_sheet(ShippingModel ship)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                string templatePath = sTemplatePath + @"\Sunzex_SHIP.xlsx";
                string outputPath = sOutputPath + @"\" + ship.filename + ".xlsx";

                app.Visible = false;
                app.Workbooks.Add();
                if (!System.IO.File.Exists(templatePath))
                {
                    MessageBox.Show("Không tìm thấy file mẫu");
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
                    MessageBox.Show(ex.Message);
                }

                _Worksheet ws = (_Worksheet)app.Workbooks[2].Worksheets[1];
                _Worksheet sheet = (_Worksheet)app.Workbooks[3].Worksheets[1];

                //check exist
                bool found = false;
                foreach (_Worksheet aSheet in app.Workbooks[3].Sheets)
                {
                    if (ship.sheetname.Equals(aSheet.Name))
                    {
                        found = true;
                        break;
                    }
                }
                if (found)
                {
                    throw new IOException("Shipping " + ship.sheetname + " đã tồn tại, không thể tạo mới!");
                }
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
                MessageBox.Show(ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi đọc/ghi/mở file: Do " + ex.Message);
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
            fileopener.StartInfo.Arguments = "\"" + path + "\"";
            fileopener.Start();
            fileopener.Close();
        }
        public class InvoiceModel
        {
            public int type, pkl_num;
            public string number, date, consignee, portload, destination, sailing, test, sheetname, filename;
            public string[] detail_name, detail_order, detail_PO, detail_quantity, detail_price, detail_size;
            public InvoiceModel(int type)
            {
                this.type = type;
                detail_name = new string[type + 1];
                detail_order = new string[type + 1];
                detail_PO = new string[type + 1];
                detail_quantity = new string[type + 1];
                detail_price = new string[type + 1];
                detail_size = new string[type + 1];
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

            //TODO
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
            invoice.sheetname = invoice.date.Substring(3, 2) + "." + a.Substring(6, a.Length - 6);

            if (checkBox_xuatRieng.Checked)
            {
                invoice.filename = @"SUNZEX_" + a;
            }
            else
            {
                invoice.filename = @"SUNZEX_INVOICE." + invoice.date.Substring(6, 4) + ".T" + invoice.date.Substring(3, 2);
            }

            try
            {
                for (int i = 1; i <= type; i++)
                {
                    invoice.detail_name[i] = (worksheet.Range["F" + (items[i].Row + 3)].Value2.Contains("giấy") ? "PAPER FOLDER" : "PP FOLDER");
                    invoice.detail_order[i] = worksheet.Range["F" + (items[i].Row + 3)].Value2.Substring(0, 3);
                    a = worksheet.Range["F" + (items[i].Row + 3)].Value2;
                    a = a.Substring(a.IndexOf("(") + 1, a.IndexOf(")") - a.IndexOf("(") - 1);
                    invoice.detail_size[i] = a.Replace(" ", "").Replace("X", "-").Replace("x", "-").Replace("*", "-").Replace("C", "").Replace("c", "").Replace("M", "").Replace("m", "").Replace(",", ".");
                    invoice.detail_quantity[i] = worksheet.Range["Q" + (items[i].Row + 6)].Value2.Replace(".", "").Replace(",", ".");
                    invoice.detail_price[i] = worksheet.Range["R" + (items[i].Row + 8)].Value2.Replace(".", "").Replace(",", ".");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                        pklapp.Visible = true;
                        pklapp.Workbooks.Add(aDialog.FileName);
                        //pklapp.Workbooks[1].Sheets[1].Activate();
                        _Worksheet sheet = pklapp.ActiveSheet;

                        Range POitems = sheet.Cells;
                        Range aPOitems;
                        aPOitems = POitems.Find("Folder size");
                        Range firstResult = aPOitems;
                        int max = 0;
                        while (!(aPOitems is null) && max++ < 20)
                        {
                            a = sheet.Range["A" + aPOitems.Row].Value2;
                            //throw new Exception("aPOitems.Row = "+ aPOitems.Row+" test valuee: " + a);
                            a = a.Substring(a.IndexOf("(") + 1, a.IndexOf(")") - a.IndexOf("(") - 1);
                            a = a.Replace(" ", "").Replace("X", "-").Replace("x", "-");
                            a = a.Replace("*", "-").Replace("C", "").Replace("c", "");
                            a = a.Replace("M", "").Replace("m", "").Replace(",", ".");
                            for (int i = 1; i <= type; i++)
                            {
                                string b;
                                if (a.Equals(invoice.detail_size[i]))
                                {
                                    b = sheet.Range["E" + (aPOitems.Row + 3)].Value2.Replace("PO#", "").Replace(" ", "");
                                    if (invoice.detail_PO[i] != null && invoice.detail_PO[i].Length > 0)
                                    {
                                        if (!invoice.detail_PO[i].Contains(b))
                                        {
                                            if (!invoice.detail_PO[i].Contains(b.Substring(0,4)))
                                            {
                                                int k;
                                                for (k = 4; i < b.Length; k++) {
                                                    if (!invoice.detail_PO[i].Contains(b.Substring(0, k))) break;
                                                }
                                                string c = b.Substring(0, k - 1);
                                                b = b.Substring(k, b.Length - k);

                                                invoice.detail_PO[i].Insert(invoice.detail_PO[i].IndexOf(c)+c.Length, (@"/" + b));
                                            }
                                            else {
                                                invoice.detail_PO[i] += ("," + b);
                                            }
                                        }
                                        break;
                                    }
                                    else
                                    {
                                        invoice.detail_PO[i] = b;
                                    }
                                    break;
                                }
                            }

                            aPOitems = POitems.FindNext(aPOitems);
                            if (aPOitems.Address == firstResult.Address) aPOitems = null;
                        }
                        object misValue = System.Reflection.Missing.Value;
                        pklapp.Workbooks[1].Close(false, misValue, misValue);
                        pklapp.Quit();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show((checkBox_PKL.Checked ? "checkBox_PKL.Checked: " : "") + ex.ToString());
                    object misValue = System.Reflection.Missing.Value;
                    pklapp.Workbooks[1].Close(false, misValue, misValue);
                    pklapp.Quit();
                }
            }

            Copy_invoice_work_sheet(invoice);
        }
        private void Copy_invoice_work_sheet(InvoiceModel invoice)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                string templatePath = sTemplatePath + @"\Sunzex_INV.xlsx";
                string outputPath = sOutputPath + @"\" + invoice.filename + ".xlsx";

                app.Visible = false;
                app.Workbooks.Add();
                if (!System.IO.File.Exists(templatePath))
                {
                    MessageBox.Show("Không tìm thấy file mẫu");
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
                    MessageBox.Show(ex.Message);
                }

                _Worksheet ws = (_Worksheet)app.Workbooks[2].Worksheets[invoice.type];
                _Worksheet sheet = (_Worksheet)app.Workbooks[3].Worksheets[1];

                //check exist
                bool found = false;
                foreach (_Worksheet aSheet in app.Workbooks[3].Sheets)
                {
                    if (invoice.sheetname.Equals(aSheet.Name))
                    {
                        found = true;
                        break;
                    }
                }
                if (found)
                {
                    throw new IOException("Invoice " + invoice.sheetname + " đã tồn tại, không thể tạo mới!");
                }
                if (IsOpenedWB_ByName(invoice.filename + ".xlsx"))
                {
                    throw new IOException("File " + invoice.filename + " đang được mở nên không thể ghi đè.  Hãy đóng file và thử tạo lại invoice! ");
                }

                //TODO
                //Begin copy
                ws.Copy(sheet);
                app.Workbooks[3].Sheets[1].Activate();
                sheet = app.ActiveSheet;

                sheet.Range["E7"].Value2 = invoice.number;
                sheet.Range["I7"].Value2 = invoice.date;
                sheet.Range["A10"].Value2 = invoice.consignee;
                sheet.Range["A17"].Value2 = invoice.portload;
                sheet.Range["C17"].Value2 = invoice.destination;
                sheet.Range["C19"].Value2 = invoice.sailing;
                sheet.Name = invoice.sheetname;

                if (invoice.type <= 3)
                {
                    int row1 = 19;
                    for (int i = 1; i <= invoice.type; i++)
                    {
                        sheet.Range["A" + (row1 + i * 3)].Value2 = invoice.detail_name[i];
                        sheet.Range["A" + (row1 + i * 3 + 1)].Value2 = @"ORDER: HSS90" + invoice.detail_order[i];
                        sheet.Range["A" + (row1 + i * 3 + 2)].Value2 = @"PO#" + invoice.detail_PO[i];
                        sheet.Range["E" + (row1 + i * 3)].Value2 = invoice.detail_quantity[i];
                        sheet.Range["G" + (row1 + i * 3)].Value2 = "0.03";
                    }
                    int row2 = row1 + 4 + 3 * invoice.type;
                    for (int i = 1; i <= invoice.type; i++)
                    {
                        sheet.Range["A" + (row2 + i * 3)].Value2 = invoice.detail_name[i];
                        sheet.Range["A" + (row2 + i * 3 + 1)].Value2 = @"ORDER: HSS90" + invoice.detail_order[i];
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
                        sheet.Range["A" + (row1 + i * 2)].Value2 = invoice.detail_name[i] + " - " + @"ORDER: HSS90" + invoice.detail_order[i];
                        sheet.Range["A" + (row1 + i * 2 + 1)].Value2 = @"PO#" + invoice.detail_PO[i];
                        sheet.Range["E" + (row1 + i * 2)].Value2 = invoice.detail_quantity[i];
                        sheet.Range["G" + (row1 + i * 2)].Value2 = "0.03";
                    }
                    int row2 = row1 + 4 + 2 * invoice.type;
                    for (int i = 1; i <= invoice.type; i++)
                    {
                        sheet.Range["A" + (row2 + i * 2)].Value2 = invoice.detail_name[i] + " - " + @"ORDER: HSS90" + invoice.detail_order[i];
                        sheet.Range["A" + (row2 + i * 2 + 1)].Value2 = @"PO#" + invoice.detail_PO[i];
                        sheet.Range["E" + (row2 + i * 2)].Value2 = invoice.detail_quantity[i];
                        sheet.Range["G" + (row2 + i * 2)].Value2 = invoice.detail_price[i];
                    }

                }
                

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
                else MessageBox.Show("Invoice " + invoice.sheetname + " tạo thành công!");
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                object misValue = System.Reflection.Missing.Value;
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

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
