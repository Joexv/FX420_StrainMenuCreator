using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using IniParser;
using Microsoft.Office.Interop.Excel;
using AppForm = System.Windows.Forms.Application;
using System.Drawing.Imaging;
using Excel = Microsoft.Office.Interop.Excel;
//using Windows.UI.Xaml.Media.Imaging;
using System.Windows.Media.Imaging;
using System.Drawing.Drawing2D;
using Microsoft.VisualBasic;
using ios = System.Runtime.InteropServices;
using System.Text;
using System.Linq;
using DataTable = System.Data.DataTable;

namespace StrainMenuCreator
{
    public partial class Form1 : Form
    {
        public string StartupPath = AppForm.StartupPath;
        public string TemplateFile = Path.Combine(AppForm.StartupPath, "Template_36.xlsx");
        public string PremadeList = Path.Combine(AppForm.StartupPath, "Premade.ini");

        public string Indica_Color;
        public string Sativa_Color;
        public string Hybrid_Color;
        public string CBD_Color;

        public string Bar1_Color;
        public string Bar2_Color;
        public string Background_Color;

        public Form1()
        {
            InitializeComponent();
        }

        private static void Extract(string nameSpace, string outDirectory, string internalFilePath, string resourceName)
        {
            var assembly = Assembly.GetCallingAssembly();

            using (var s =
                assembly.GetManifestResourceStream(nameSpace + "." +
                                                   (internalFilePath == "" ? "" : internalFilePath + ".") +
                                                   resourceName))
            using (var r = new BinaryReader(s))
            using (var fs = new FileStream(outDirectory + "\\" + resourceName, FileMode.OpenOrCreate))
            using (var w = new BinaryWriter(fs))
            {
                w.Write(r.ReadBytes((int)s.Length));
            }
        }

        private static void ExtractFile(string FileName)
        {
            Extract("StrainMenuCreator", AppForm.StartupPath, "Files", FileName);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!File.Exists("INIFileParser.dll"))
                ExtractFile("INIFileParser.dll");
            if (!File.Exists("INIFileParser.xml"))
                ExtractFile("INIFileParser.xml");
            if (!File.Exists("Premade.ini"))
                ExtractFile("Premade.ini");
            if (!File.Exists("Template_36.xlsx"))
                ExtractFile("Template_36.xlsx");
            if (!File.Exists("Template_42.xlsx"))
                ExtractFile("Template_42.xlsx");
            DialogResult dialogResult = MessageBox.Show("There's a premade flower list, load that for a base?", "???", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                var yesterday = DateTime.Today.AddDays(-1);
                Console.WriteLine(yesterday.ToString("MM-dd-yyyy"));
                if (File.Exists("Premade" + DateAndTime.Today.ToString("MM-dd-yyyy") + ".ini"))
                    dataGridView1.DataSource = DataTable("Premade" + DateAndTime.Today.ToString("MM-dd-yyyy") + ".ini");
                //LoadPremade("Premade" + DateAndTime.Today.ToString("MM-dd-yyyy") + ".ini");
                else if (File.Exists("Premade" + yesterday.ToString("MM-dd-yyyy") + ".ini"))
                    dataGridView1.DataSource = DataTable("Premade" + yesterday.ToString("MM-dd-yyyy") + ".ini"); //LoadPremade("Premade" + yesterday.ToString("MM-dd-yyyy") + ".ini");
                else
                    dataGridView1.DataSource = DataTable("Premade.ini");
            }

            if (File.Exists(TemplateFile))
                Template_Label.Text = "Template File Loaded!";
            if (!File.Exists(PremadeList))
                Premade_Butt.Enabled = false;

            this.FormClosed += form_FormClosed;
        }
        private void form_FormClosed(object sender, FormClosedEventArgs e)
        {
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
                process.WaitForExit();
            }
        }


        public void AdjustColors()
        {
            try
            {
                Bar1_Box.BackColor = GetColor(Bar1_Box.Text);
                Bar2_Box.BackColor = GetColor(Bar2_Box.Text);
                Background_Box.BackColor = GetColor(Background_Box.Text);

                Indica_Box.BackColor = GetColor(Indica_Box.Text);
                Hybrid_Box.BackColor = GetColor(Hybrid_Box.Text);
                Sativa_Box.BackColor = GetColor(Sativa_Box.Text);
                Heavy_Box.BackColor = GetColor(Heavy_Box.Text);
            }
            catch { }
        }

        public Color GetColor(string Hex)
        {
            return ColorTranslator.FromHtml("#" + Hex);
        }

        private void Indica_Box_TextChanged(object sender, EventArgs e)
        {
            AdjustColors();
        }

        private void Sativa_Box_TextChanged(object sender, EventArgs e)
        {
            AdjustColors();
        }

        private void Hybrid_Box_TextChanged(object sender, EventArgs e)
        {
            AdjustColors();
        }

        private void Heavy_Box_TextChanged(object sender, EventArgs e)
        {
            AdjustColors();
        }

        private void Bar1_Box_TextChanged(object sender, EventArgs e)
        {
            AdjustColors();
        }

        private void Bar2_Box_TextChanged(object sender, EventArgs e)
        {
            AdjustColors();
        }

        private void Background_Box_TextChanged(object sender, EventArgs e)
        {
            AdjustColors();
            if (Background_Box.Text.ToUpper() != "F0F0F0")
            {
                Logo_Box.Enabled = false;
                Logo_Box.Checked = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Bar1_Box.Text = Bar1_Color;
            Bar2_Box.Text = Bar2_Color;
            Background_Box.Text = Background_Color;

            Indica_Box.Text = "1f4e78";
            Sativa_Box.Text = "FF0000";
            Hybrid_Box.Text = "375623";
            Heavy_Box.Text = "ed7d31";

            Logo_Box.Enabled = true;
            Logo_Box.Checked = true;

            Range1.Text = "A1";
            Range2.Text = "L33";

            Image_Width.Value = 2955;
            Image_Height.Value = 2164;

            Tax.Value = 20;
            FontSize.Value = 28;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button4.Text = "Processing...";
            CreateExcel();
            //EditExcel();
            EditExcel_Table();
            //CreateImage();
            CreateImage_Alt();
            ResizeImage("menu_Small.png");
            MessageBox.Show("Menu has been created and should be located on your desktop!", "Done!");
            button4.Text = "Create Menu";
            var Premade = "Premade" + DateTime.Today.ToString("MM-dd-yyyy-hh-mm") + "_AUTO_BACKUP_.ini";
            GenPremade(Premade);
        }

        public void EditExcel_Table()
        {
            try
            {
                Console.WriteLine("Starting Excel edit...");
                Excel.Application excel = new Excel.Application();
                Excel.Workbook wkb = excel.Workbooks.Open(Path.Combine(StartupPath, "Generated.xlsx"));
                //Excel.Workbook wkb = excel.Workbooks.Open("Generated.xlsx");
                Excel.Worksheet sheet = wkb.Worksheets[1] as Excel.Worksheet;
                Console.WriteLine("Opened Excel file, starting edit...");
                //B-E4 - B-E21 First Row
                //H-K4 - H-K21 Second Row

                //.Range["A1:L33"]
                int Letter = 2;
                int Num = 4;
                Range row = sheet.Rows.Cells[2, 1];

                #region Create lists from DataGridView
                List<string> Values = new List<string>();

                DataTable sam = (DataTable)dataGridView1.DataSource;
                DataView dt1 = new DataView(sam);
                dt1.Sort = "Cost ASC";

                DataTable dt = dt1.ToTable();            
                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[0].ToString());
                string[] Names = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[1].ToString());
                string[] Types = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[2].ToString());
                string[] Costs = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[3].ToString());
                string[] THCs = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[4].ToString());
                string[] CBDs = Values.ToArray();
                Values.Clear();
                #endregion

                int i = 0;
                foreach (var name in Names)
                {
                    if (name != "" && name != "/n" && name != "/r")
                    {
                        if (Num >= ((TemplateMax / 2) + 3) && Letter != 8)
                        {
                            Letter = 8;
                            Num = 4;
                        }
                        if (Num >= (TemplateMax + 4))
                        {
                            DialogResult dialogResult = MessageBox.Show("You have more flowers than what this template supports continue?", "???", MessageBoxButtons.YesNo);
                            if (dialogResult == DialogResult.No)
                            {
                                break;
                            }
                        }

                        Console.WriteLine(Letter);
                        Console.WriteLine(GetLetter(Letter));
                        Console.WriteLine(GetLetter(Letter) + Num);

                        //Flower Name
                        row = sheet.Rows.Cells[Num, Letter];
                        row.Value = name;
                        row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        row.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        row.Font.Size = FontSize.Value;
                        row.Font.Bold = true;
                        row.Font.Color = GetColor(DetermineColor(Types[i]));

                        //Flower Cost
                        row = sheet.Rows.Cells[Num, Letter + 1];

                        #region Tax Calculation
                        decimal Cost_ = 0;
                        //MessageBox.Show(Costs[i]);
                        if (Costs[i].Substring(0, 1) == "0")
                        {
                            //MessageBox.Show(Costs[i].Substring(1));
                            Cost_ = Int32.Parse(Costs[i].Substring(1));
                            //MessageBox.Show(Cost_.ToString());
                        }
                        else
                            Cost_ = Int32.Parse(Costs[i]);
                        decimal percent = (Tax.Value / 100);
                        decimal test = Cost_ * percent;
                        Cost_ = Cost_ + test;
                        //MessageBox.Show(test.ToString() + percent.ToString() + Decimal.ToInt32(test).ToString());
                        #endregion

                        row.Value = Cost_;
                        row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        row.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        row.Font.Size = FontSize.Value;
                        row.Font.Bold = true;
                        row.Font.Color = GetColor(DetermineColor(Types[i]));

                        //Flower THC%
                        row = sheet.Rows.Cells[Num, Letter + 2];
                        row.Value = THCs[i] + "%";
                        row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        row.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        row.Font.Size = FontSize.Value;
                        row.Font.Bold = true;
                        row.Font.Color = GetColor(DetermineColor(Types[i]));

                        //Flower CBD%
                        row = sheet.Rows.Cells[Num, Letter + 3];
                        row.Value = CBDs[i] + "%";
                        row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        row.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        row.Font.Size = FontSize.Value;
                        row.Font.Bold = true;
                        row.Font.Color = GetColor(DetermineColor(Types[i]));


                        //Range r = sheet.Range[Merge(Letter, Num) + ":" + Merge(Letter, Num + 3)];
                        //r.Font.Size = 26;
                        //r.Font.Color = GetColor(DetermineColor(Types[i]));
                    }
                    i++;
                    Num++;
                }
                Console.WriteLine("Done editing saving...");
                excel.Application.ActiveWorkbook.Save();
                object objFalse = false;
                wkb.Close(true, Type.Missing, Type.Missing);
                excel.Quit();
                Console.WriteLine("Done");
            }
            catch(Exception e) { MessageBox.Show(e.ToString()); }
        }

        public string DetermineColor(string Type)
        {
            Type = Type.ToUpper();
            Console.WriteLine(Type);
            //MessageBox.Show(Type);
            if(Type.Contains("INDICA"))
                return Indica_Box.Text;
            if (Type.Contains("SATIVA"))
                return Sativa_Box.Text;
            if (Type.Contains("CBD"))
                return Heavy_Box.Text;
            if (Type.Contains("HYBRID"))
                return Hybrid_Box.Text;
            else
            {
                MessageBox.Show("No Color found for strain: " + Type);
                return Hybrid_Box.Text;
            }
        }

        const int ColumnBase = 26;
        const int DigitMax = 7; // ceil(log26(Int32.Max))
        const string Digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public static string GetLetter(int index)
        {
            if (index <= 0)
                throw new IndexOutOfRangeException("index must be a positive number");

            if (index <= ColumnBase)
                return Digits[index - 1].ToString();

            var sb = new StringBuilder().Append(' ', DigitMax);
            var current = index;
            var offset = DigitMax;
            while (current > 0)
            {
                sb[--offset] = Digits[--current % ColumnBase];
                current /= ColumnBase;
            }
            return sb.ToString(offset, DigitMax - offset);
        }

        public string Merge(int Num_Letter, int Num)
        {
            return (GetLetter(Num_Letter) + Num.ToString());
        }


        private void AllBorders(Borders _borders)
        {
            _borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            _borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            _borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            _borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            _borders.Color = Color.Black;
        }

        //xlWorkSheet.Columns[5].ColumnWidth = 18

        private static Bitmap ResetResolution(Metafile mf, float resolution)
        {
            int width = (int)(mf.Width * resolution / mf.HorizontalResolution);
            int height = (int)(mf.Height * resolution / mf.VerticalResolution);
            Bitmap bmp = new Bitmap(width, height);
            bmp.SetResolution(resolution, resolution);
            Graphics g = Graphics.FromImage(bmp);
            g.DrawImage(mf, 0, 0);
            g.Dispose();
            return bmp;
        }

        private void CreateExcel()
        {
            try
            {
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    process.Kill();
                    process.WaitForExit();
                }
                if (File.Exists("Generated.xlsx"))
                    File.Delete("Generated.xlsx");
            }
            catch(Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            File.Copy(TemplateFile, "Generated.xlsx");
        }

        private void CreateImage()
        {
            if (File.Exists(@"C:\Users\Public\Public/ Desktop\Menu.png"))
                File.Delete(@"C:\Users\Public\Public/ Desktop\Menu.png");
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wkb = excel.Workbooks.Open(Path.Combine(StartupPath, "Generated.xlsx"));
            Excel.Worksheet sheet = wkb.Worksheets[1] as Excel.Worksheet;
            Excel.Range range = sheet.Cells[33, 12] as Excel.Range;
            range.Formula = "";

            // copy as seen when printed
            //range.CopyPicture(Excel.XlPictureAppearance.xlPrinter, Excel.XlCopyPictureFormat.xlPicture);

            // uncomment to copy as seen on screen
            range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

            Console.WriteLine("Please enter a full file name to save the image from the Clipboard:");
            Console.WriteLine("Menu_Small.jpeg");
            string fileName = "Menu_Small.jpeg";
            using (FileStream fileStream = new FileStream(fileName, FileMode.Create))
            {
                if (Clipboard.ContainsData(DataFormats.EnhancedMetafile))
                {
                    Metafile metafile = Clipboard.GetData(DataFormats.EnhancedMetafile) as Metafile;
                    metafile.Save(fileName);
                }
                else if (Clipboard.ContainsData(DataFormats.Bitmap))
                {
                    BitmapSource bitmapSource = Clipboard.GetData(DataFormats.Bitmap) as BitmapSource;

                    JpegBitmapEncoder encoder = new JpegBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                    encoder.QualityLevel = 100;
                    encoder.Save(fileStream);
                }
            }
            object objFalse = false;
            wkb.Close(objFalse, Type.Missing, Type.Missing);
            excel.Quit();
        }

        public void ExportRangeAsJpg()
        {
            Excel.Application xl;

            xl = (Excel.Application)ios.Marshal.GetActiveObject("Excel.Application");

            if (xl == null)
            {
                MessageBox.Show("No Excel !!");
                return;
            }

            Excel.Workbook wb = xl.ActiveWorkbook;
            Excel.Range r = wb.ActiveSheet.Range["A1:L33"];
            r.CopyPicture(Excel.XlPictureAppearance.xlScreen,
                           Excel.XlCopyPictureFormat.xlBitmap);

            if (Clipboard.GetDataObject() != null)
            {
                IDataObject data = Clipboard.GetDataObject();

                if (data.GetDataPresent(DataFormats.Bitmap))
                {
                    Image image = (Image)data.GetData(DataFormats.Bitmap, true);
                    image.Save("sample.jpg",
                        System.Drawing.Imaging.ImageFormat.Jpeg);
                }
                else
                {
                    MessageBox.Show("No image in Clipboard !!");
                }
            }
            else
            {
                MessageBox.Show("Clipboard Empty !!");
            }
        }

        public void CreateImage_Alt()
        {
            Console.WriteLine("Creating initial image...");
            Excel.Application excel = new Excel.Application();
            Excel.Workbook w = excel.Workbooks.Open(Path.Combine(StartupPath, "Generated.xlsx")); ;
            Worksheet ws = w.Sheets[1];
            ws.Protect(Contents: false);
            Console.WriteLine("Range is " + Range1 + ":" + Range2);
            Console.WriteLine("Size is " + Image_Width.Value.ToString() + " x " + Image_Height.Value.ToString());
            string ImageRange = Range1.Text + ":" + Range2.Text;
            Range r = ws.Range[ImageRange];
            r.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);

            Bitmap image = new Bitmap(Clipboard.GetImage());
            image.Save("Menu_Small.png");
            Console.WriteLine("Small image saved...");
            w.Close(false, Type.Missing, Type.Missing);
            excel.Quit();
        }

        public void ResizeImage(string fileName)
        {
            Console.WriteLine("Resizing image...");
            FileInfo info = new FileInfo(fileName);
            using (Image image = Image.FromFile(fileName))
            {
                using (Bitmap resizedImage = ImageHelper.ResizeImage(image, 2m))
                {
                    resizedImage.Save(Environment.GetFolderPath
                   (Environment.SpecialFolder.DesktopDirectory) + "\\" + "Menu_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png", ImageFormat.Png);
                }
            }
            Console.WriteLine("Done");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var Premade = "Premade" + DateTime.Today.ToString("MM-dd-yyyy") + ".ini";
            GenPremade(Premade);
        }

        public void GenPremade(string Premade)
        {
            Console.WriteLine("Generating Premade...");
            using (var sw = File.CreateText(Premade)) { }

            var parser = new FileIniDataParser();
            var data = parser.ReadFile(Premade);

            #region Create lists from DataGridView
            List<string> Values = new List<string>();
            DataTable dt = (DataTable)dataGridView1.DataSource;
            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[0].ToString());
            string[] Names = Values.ToArray();
            Values.Clear();

            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[1].ToString());
            string[] Types = Values.ToArray();
            Values.Clear();

            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[2].ToString());
            string[] Costs = Values.ToArray();
            Values.Clear();

            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[3].ToString());
            string[] THCs = Values.ToArray();
            Values.Clear();

            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[4].ToString());
            string[] CBDs = Values.ToArray();
            Values.Clear();
            #endregion

            int i = 1;
            /*
            string[] Names = Name_Box.Text.Split('\n');
            string[] Costs = Cost_Box.Text.Split('\n');
            string[] Types = Type_Box.Text.Split('\n');
            string[] THCs = THC_Box.Text.Split('\n');
            string[] CBDs = CBD_Box.Text.Split('\n');*/

            int Total = Names.Length;
            foreach (var name in Names)
            {
                if (name != "/r" && name != "/n" && name != "")
                {
                    data["Flower"][i.ToString()] = name;
                    data["Cost"][i.ToString()] = Costs[i - 1];
                    data["Types"][i.ToString()] = Types[i - 1];
                    data["THC"][i.ToString()] = THCs[i - 1];
                    data["CBD"][i.ToString()] = CBDs[i - 1];
                    i++;
                }
                else
                {
                    Total--;
                }
            }
            data["Settings"]["Total"] = Total.ToString();

            data["Settings"]["Indica"] = Indica_Box.Text;
            data["Settings"]["Sativa"] = Sativa_Box.Text;
            data["Settings"]["Hybrid"] = Hybrid_Box.Text;
            data["Settings"]["CBD"] = Heavy_Box.Text;

            data["Settings"]["Range1"] = Range1.Text;
            data["Settings"]["Range2"] = Range2.Text;

            data["Settings"]["Width"] = Image_Width.Value.ToString();
            data["Settings"]["Height"] = Image_Height.Value.ToString();

            data["Settings"]["Tax"] = Tax.Value.ToString();
            data["Settings"]["Font"] = FontSize.Value.ToString();

            parser.WriteFile(Premade, data);
        }

        public string OpenPremade;

        private void Premade_Butt_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Premade ini File (*.ini)|*.ini";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string filePath = ofd.FileName;
                ofd.Dispose();

                OpenPremade = Path.GetFileNameWithoutExtension(filePath) + ".ini";
                dataGridView1.DataSource = DataTable(OpenPremade);
            }
        }

        public DataTable DataTable(string filePath)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Flowers");
            dt.Columns.Add("Types");
            dt.Columns.Add("Cost");
            dt.Columns.Add("THC");
            dt.Columns.Add("CBD");
            Console.WriteLine("Attempting to read Premade");
            var parser = new FileIniDataParser();
            var data = parser.ReadFile(filePath);
            try
            {

                Console.WriteLine("Premade loaded... Reading Settings>Total");
                int Total = Int32.Parse(data["Settings"]["Total"]);
                Console.WriteLine("Total = " + Total.ToString());
                for (int i = 0; i < Total; i++)
                {
                    DataRow dr = dt.NewRow();
                    Console.WriteLine("Reading " + i + "/" + (i + 1));
                    dr["Flowers"] = data["Flower"][(i + 1).ToString()];
                    dr["Types"] = data["Types"][(i + 1).ToString()];

                    if(int.Parse(data["Cost"][(i + 1).ToString()]) < 10)
                        dr["Cost"] = "0" + data["Cost"][(i + 1).ToString()];
                    else
                        dr["Cost"] = data["Cost"][(i + 1).ToString()];

                    dr["THC"] = data["THC"][(i + 1).ToString()];
                    dr["CBD"] = data["CBD"][(i + 1).ToString()];
                    
                    dt.Rows.Add(dr);
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }

            Console.WriteLine("Loading Premade settings...");
            try
            {
                Indica_Box.Text = data["Settings"]["Indica"];
                Sativa_Box.Text = data["Settings"]["Sativa"];
                Hybrid_Box.Text = data["Settings"]["Hybrid"];
                Heavy_Box.Text = data["Settings"]["CBD"];

                Range1.Text = data["Settings"]["Range1"];
                Range2.Text = data["Settings"]["Range2"];

                Image_Width.Value = decimal.Parse(data["Settings"]["Width"]);
                Image_Height.Value = decimal.Parse(data["Settings"]["Height"]);

                Tax.Value = decimal.Parse(data["Settings"]["Tax"]);
                FontSize.Value = decimal.Parse(data["Settings"]["Font"]);
            }
            catch { }
            return dt;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Indica_Color = Indica_Box.Text;
            Hybrid_Color = Hybrid_Box.Text;
            CBD_Color = Heavy_Box.Text;
            Sativa_Color = Sativa_Box.Text;

            Bar1_Color = Bar1_Box.Text;
            Bar2_Color = Bar2_Box.Text;
            Background_Color = Background_Box.Text;

            AdjustColors();

            List<string> Values = new List<string>();
            DataTable dt = (DataTable)dataGridView1.DataSource;
            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[0].ToString());
            string[] Names = Values.ToArray();

            Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                #region Create lists from DataGridView
                List<string> Values = new List<string>();
                DataTable dt = (DataTable)dataGridView1.DataSource;
                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[0].ToString());
                string[] Names = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[1].ToString());
                string[] Types = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[2].ToString());
                string[] Costs = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[3].ToString());
                string[] THCs = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[4].ToString());
                string[] CBDs = Values.ToArray();
                Values.Clear();
                #endregion

                int i = Array.IndexOf(Names, Delete_Box.Text);

                var list = new List<string>(Names);
                list.RemoveAt(i);
                Names = list.ToArray();

                list = new List<string>(Costs);
                list.RemoveAt(i);
                Costs = list.ToArray();

                list = new List<string>(Types);
                list.RemoveAt(i);
                Types = list.ToArray();

                list = new List<string>(THCs);
                list.RemoveAt(i);
                THCs = list.ToArray();

                list = new List<string>(CBDs);
                list.RemoveAt(i);
                CBDs = list.ToArray();

                //Name_Box.Text = String.Join("\n", Names);
                //Cost_Box.Text = String.Join("\n", Costs);
                //Type_Box.Text = String.Join("\n", Types);
                //THC_Box.Text = String.Join("\n", THCs);
                //CBD_Box.Text = String.Join("\n", CBDs);

                dt.Rows.RemoveAt(i);
                dataGridView1.DataSource = dt;
                Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
            }
            catch { }
        }

        public int TemplateMax = 36;

        private void button1_Click(object sender, EventArgs e)
        {
            OpenTemplate();
        }

        public void OpenTemplate()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Template File (*.xlsx)|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string filePath = ofd.FileName;
                ofd.Dispose();

                TemplateFile = Path.GetFileNameWithoutExtension(filePath);
                Template_Label.Text = "Loaded, " + TemplateFile;
                TemplateMax = Int32.Parse(TemplateFile.Replace("Template_", ""));
                TemplateFile = filePath;
            }
        }

        private void Name_Box_TextChanged(object sender, EventArgs e)
        {
            CountCheck();
        }

        private void Delete_Box_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            List<string> Values = new List<string>();
            DataTable dt = (DataTable)dataGridView1.DataSource;
            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[0].ToString());
            string[] Names = Values.ToArray();
            Values.Clear();

            Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
        }

        public void CountCheck()
        {
            List<string> Values = new List<string>();
            DataTable dt = (DataTable)dataGridView1.DataSource;
            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[0].ToString());
            string[] Names = Values.ToArray();

            FlowerCount.Text = "Flower Count: " + Names.Count();
            if (Names.Count() > TemplateMax)
            {
                DialogResult dialogResult = MessageBox.Show("You have more flowers than what this template supports. Would you like to swap to a larger template?", "???", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    OpenTemplate();
                    if (TemplateMax < Names.Count())
                     CountCheck();
                }
            }
        }

        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            CountCheck();
            List<string> Values = new List<string>();
            DataTable dt = (DataTable)dataGridView1.DataSource;
            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[0].ToString());
            string[] Names = Values.ToArray();
            Values.Clear();

            Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Process.Start(@"https://www.w3schools.com/colors/colors_picker.asp");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //this.dataGridView1.DataSource.Rows.Add("Flowers", "Types", "Cost", "THC", "CBD");
            //this.dataGridView1.Rows.Insert(0, "", "" , "", "", "");
            DataTable dt = (DataTable)dataGridView1.DataSource;
            dt.NewRow();
            dataGridView1.DataSource = dt;

            List<string> Values = new List<string>();
            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[0].ToString());
            string[] Names = Values.ToArray();
            Values.Clear();

            Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
        }

        public Decimal Final_Width
        {
            get { return Image_Width.Value; }
            set { Image_Width.Value = value; }
        }

        public Decimal Final_Height
        {
            get { return Image_Height.Value; }
            set { Image_Height.Value = value; }
        }
    }

    public static class ImageHelper
    {
        /// <summary>
        /// Resize the image to the specified width and height.
        /// </summary>
        /// <param name="image">The image to resize.</param>
        /// <param name="width">The width to resize to.</param>
        /// <param name="height">The height to resize to.</param>
        /// <returns>The resized image.</returns>
        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            Console.WriteLine(width);
            Console.WriteLine(height);
            var destRect = new System.Drawing.Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }
        public static Form1 f1;

        public static Bitmap ResizeImage(Image image, decimal percentage)
        {
            f1 = new Form1();
            int width = Decimal.ToInt32(f1.Final_Width); //(int)Math.Round(image.Width * percentage, MidpointRounding.AwayFromZero);
            int height = Decimal.ToInt32(f1.Final_Height); //(int)Math.Round(image.Height * percentage, MidpointRounding.AwayFromZero);
            return ResizeImage(image, width, height);
        }
    }
}