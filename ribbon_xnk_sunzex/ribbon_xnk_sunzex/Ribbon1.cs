using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;

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
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            sTemplatePath = ReadFromRegistry(KEY_NAME_SUNZEX_templatePath, "");
            editBox_hoadonmau.Text = sTemplatePath;

            sOutputPath = ReadFromRegistry(KEY_NAME_SUNZEX_outputPath, sTemplatePath);
            editBox_thumucxuat.Text = sOutputPath;

            checkBox_openAfter.Checked = "1".Equals(ReadFromRegistry(KEY_NAME_SUNZEX_openAfter, "1"));

            checkBox_xuatRieng.Checked = "1".Equals(ReadFromRegistry(KEY_NAME_SUNZEX_outSeparate, "0"));
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            folderBrowserDialog1.SelectedPath = editBox_hoadonmau.Text;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                editBox_hoadonmau.Text = folderBrowserDialog1.SelectedPath;
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
            folderBrowserDialog1.SelectedPath = editBox_thumucxuat.Text;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                editBox_thumucxuat.Text = folderBrowserDialog1.SelectedPath;
                StoreInRegistry(KEY_NAME_SUNZEX_outputPath, editBox_thumucxuat.Text);
                sOutputPath = editBox_thumucxuat.Text;
            }
        }

        private void button_shipping_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = (Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = wb.ActiveSheet;

            //Check
            if (!"Số tờ khai".Equals(worksheet.Range["C4"].Value2) || !worksheet.Range["E2"].Value2.Contains("Tờ khai")) {
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

            if (checkBox_xuatRieng.Checked) {
                ship.filename = @"SUNZEX_SHIPING." + worksheet.Range["R49"].Value2.Replace("/", "");
                ship.sheetname = worksheet.Range["R49"].Value2.Replace("/", "");
            } else {
                a = ship.comodity.Substring(ship.comodity.IndexOf(@":") + 6, 2) + "." + ship.comodity.Substring(ship.comodity.IndexOf(@":") + 8, 3);
                ship.sheetname = a.Replace("/", "");

                ship.filename = @"SUNZEX_SHIPING.T" + ship.date.Substring(3, 2) + "." + ship.date.Substring(6, 4);
            }
            
            Copy_shipping_work_sheet(ship);
        }
        private void Copy_invoice_work_sheet(InvoiceModel invoice)
        { 
        
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
                foreach (Microsoft.Office.Interop.Excel._Worksheet aSheet in app.Workbooks[3].Sheets) {
                    if (ship.sheetname.Equals(aSheet.Name)) {
                        found = true;
                        break;
                    }
                }
                if (found) {
                    throw new IOException();
                }

                ws.Copy(sheet);
                app.Workbooks[3].Sheets[1].Activate();
                sheet = app.ActiveSheet;
                //Modify
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

                app.ActiveWorkbook.SaveAs(outputPath);
                app.Workbooks[3].Close();
                app.Workbooks[2].Close();
                app.Workbooks[1].Close();
                app.Workbooks.Close();
                app.Quit();
                if (checkBox_openAfter.Checked) {
                    OpenFile(outputPath);
                } else MessageBox.Show("Shipping " + ship.sheetname + " tạo thành công!");
            }
            catch (IOException ex)
            {
                MessageBox.Show("Shipping "+ship.sheetname+" đã tồn tại, không thể tạo mới! "+ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally {
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
        }
        public class InvoiceModel
        {
            public int type;
            public string number, date, consignee, portload, destination, sailing;
            public string[] detail_name, detail_order, detail_PO, detail_quantity, detail_price;
            public InvoiceModel(int type)
            {
                this.type = type;
                detail_name = new string[3];
                detail_order = new string[3];
                detail_PO = new string[3];
                detail_quantity = new string[3];
                detail_price = new string[3];
            }
        }
        public class ShippingModel
        {
            public string date, shipment, transport, destination, portload, voyage, comodity, amount, total, gross, filename, sheetname;
            public ShippingModel() { }
        }

        private void CheckBox3_Click(object sender, RibbonControlEventArgs e)
        {
            StoreInRegistry(KEY_NAME_SUNZEX_openAfter, checkBox_openAfter.Checked ? "1" : "0");
        }

        private void checkBox_xuatRieng_Click(object sender, RibbonControlEventArgs e)
        {
            StoreInRegistry(KEY_NAME_SUNZEX_outSeparate, checkBox_xuatRieng.Checked ? "1" : "0");
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
            //find type number
            int type = 1;
            for (int i = 1; i < 10; i++) {
                if (findInvoiceItems(("" + i).PadLeft(2), worksheet) != -1) {
                    type = i;
                } else break;
            }

            InvoiceModel invoice = new InvoiceModel(type);
            //TO DO
            string a;
            invoice.number = "No: "+ worksheet.Range["R49"].Value2;
            a = worksheet.Range["F8"].Value2;
            invoice.date = a.Substring(0, 10);
            a = worksheet.Range["H47"].Value2;
            invoice.consignee = a.Substring(0, a.Length - a.IndexOf("TONZEX")).Replace("GIAO HANG CHO", "").Replace(" THEO CHI DINH CUA ", "");
            invoice.portload = worksheet.Range[""].Value2;
            invoice.destination = worksheet.Range[""].Value2;
            invoice.sailing = worksheet.Range[""].Value2;

            Copy_invoice_work_sheet(invoice);
        }
        private int findInvoiceItems(string n, Worksheet ws)
        {
            Range currentFind = null;
            Range firstFind = null;

            Range Items = ws.get_Range("C1", "C400");
            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = Items.Find(@"<" + n + ">", Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(XlReferenceStyle.xlA1)
                      == firstFind.get_Address(XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                currentFind.Font.Bold = true;

                currentFind = Items.FindNext(currentFind);
            }
            return -1;
        }
    }
}
