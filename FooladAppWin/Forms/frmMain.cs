using ClosedXML.Excel;
using FooladAppWin.Classes;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FooladAppWin.Forms
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {

            string mainPath = Application.StartupPath+"\\";

            string jsonPath =mainPath+ "data.json";
            string templatePath =mainPath+ "template.xlsx";
            string outputPath = mainPath + "output.xlsx";


          


            if (!File.Exists(jsonPath) || !File.Exists(templatePath))
            {
                MessageBox.Show("فایل json یا فایل قالب پیدا نشد.");
                return;
            }

            var json = File.ReadAllText(jsonPath);

            var allRecords = JsonConvert.DeserializeObject<List<PersonnelRecord>>(json);

            var grouped = allRecords
                .GroupBy(r => new { r.PersonnelNumber, r.FullName })
                .OrderBy(g => g.Key.PersonnelNumber);

             var templateStream = File.OpenRead(templatePath);

             var workbook = new XLWorkbook(templateStream);

            var templateSheet = workbook.Worksheet("Sheet1");

            int RecordCount = 1;

            foreach (var group in grouped)
            {
                string sheetName = $"{RecordCount.ToString()}_{group.Key.FullName}";

                if (sheetName.Length > 31)

                    sheetName = sheetName.Substring(0, 31);

                var personSheet = templateSheet.CopyTo(sheetName);

                // 👇 اضافه کن تا نام در سلول A1 قرار بگیره
                personSheet.Cell("A1").Value = group.Key.FullName;
                personSheet.Cell("E1").Value = txtTitle.Text;

                for (int row = 2; row <= 48; row++)
                {
                    for (int col = 1; col <= 6; col++)
                    {
                       // personSheet.Cell(row, col).Clear(XLClearOptions.Contents);
                    }
                }



                //int dataRow = 3;
                //foreach (var record in group)
                //{
                //    if (dataRow > 48) break;


                 

                //    personSheet.Cell(dataRow, 1).Value = record.PersonnelNumber;
                //    personSheet.Cell(dataRow, 2).Value = record.FullName;
                //    personSheet.Cell(dataRow, 3).Value = record.Date.ToString();
                //    personSheet.Cell(dataRow, 4).Value = record.Day;
                //    personSheet.Cell(dataRow, 5).Value = record.Time;
                //    personSheet.Cell(dataRow, 6).Value = record.Status;


                //    dataRow++;
                //}



                // گروه‌بندی بر اساس تاریخ
                var dailyGroups = group
                    .GroupBy(r => r.Date)
                    .OrderBy(g => g.Key);

                int dataRow = 3;
                foreach (var dayGroup in dailyGroups)
                {
                    var entry = dayGroup.FirstOrDefault(r => r.Status == "ورود");
                    var exit = dayGroup.FirstOrDefault(r => r.Status == "خروج");

                    // ردیف ورود
                    personSheet.Cell(dataRow, 1).Value = group.Key.PersonnelNumber;
                    personSheet.Cell(dataRow, 2).Value = group.Key.FullName;
                    personSheet.Cell(dataRow, 3).Value = dayGroup.Key;
                    personSheet.Cell(dataRow, 4).Value = entry?.Day ?? "";
                    personSheet.Cell(dataRow, 5).Value = entry?.Time ?? "0";
                    personSheet.Cell(dataRow, 6).Value = "ورود";
                    dataRow++;

                    // ردیف خروج
                    personSheet.Cell(dataRow, 1).Value = group.Key.PersonnelNumber;
                    personSheet.Cell(dataRow, 2).Value = group.Key.FullName;
                    personSheet.Cell(dataRow, 3).Value = dayGroup.Key;
                    personSheet.Cell(dataRow, 4).Value = exit?.Day ?? "";
                    personSheet.Cell(dataRow, 5).Value = exit?.Time ?? "0";
                    personSheet.Cell(dataRow, 6).Value = "خروج";
                    dataRow++;

                    if (dataRow > 48) break;
                }






                personSheet.RightToLeft = true;
                RecordCount++;
            }

            // شیت الگو را تغییر نام داده و به اول منتقل کن
            templateSheet.Name = "الگو";
            templateSheet.Position = 1;
            workbook.SaveAs(outputPath);
            MessageBox.Show("خروجی اکسل با موفقیت ساخته شد.", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public  void ImportExcel()
        {
             var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "انتخاب فایل اکسل"
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;

            var filePath = openFileDialog.FileName;
            var records = new List<PersonnelRecord>();

             var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1); // اولین شیت

            int row = 2; // فرض بر این که ردیف اول هدر است
            while (!worksheet.Cell(row, 1).IsEmpty())
            {
                try
                {

       

                    var record = new PersonnelRecord
                    {
                        PersonnelNumber = int.Parse(worksheet.Cell(row, 1).GetValue<string>()),
                        FullName = worksheet.Cell(row, 2).GetValue<string>(),
                        Date =(worksheet.Cell(row, 3).GetValue<string>()),
                        Day = worksheet.Cell(row, 4).GetValue<string>(),
                        Time = worksheet.Cell(row, 5).GetValue<string>(),
                        Status = worksheet.Cell(row, 6).GetValue<string>()
                    };


      

                    records.Add(record);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"خطا در ردیف {row}: {ex.Message}");
                }

                row++;
            }

            // تبدیل به JSON و ذخیره در فایل
            var json = JsonConvert.SerializeObject(records, Formatting.Indented);
            File.WriteAllText("data.json", json);

            MessageBox.Show("اطلاعات با موفقیت به فایل JSON ذخیره شد.", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public  void  OpenSource()
        {
            // دریافت مسیر پوشه اجرایی برنامه
            string exePath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            // باز کردن پنجره Explorer در مسیر مورد نظر
            System.Diagnostics.Process.Start("explorer.exe", exePath);
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            ImportExcel();
        }

        private void btnOpenExplorer_Click(object sender, EventArgs e)
        {
            OpenSource();
        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            string url = "https://www.sadegh-dehghani.ir";
            System.Diagnostics.Process.Start(url);
        }
    }
}
