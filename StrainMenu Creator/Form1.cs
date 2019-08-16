using IniParser;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AppForm = System.Windows.Forms.Application;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Data;
using Renci.SshNet;

namespace StrainMenuCreator
{
    public partial class Form1 : Form
    {
        public string StartupPath = AppForm.StartupPath;
        public string TemplateFile = Path.Combine(AppForm.StartupPath, "Template_40.xlsx");
        public string PremadeList = Path.Combine(AppForm.StartupPath, "Premade.ini");

        public string Indica_Color;
        public string Sativa_Color;
        public string Hybrid_Color;
        public string CBD_Color;

        public string Bar1_Color;
        public string Bar2_Color;
        public string Background_Color;

        private static string oauth => File.Exists(@"Z:\Slack Bot\SlackBot_Auth.txt") ? File.ReadAllText(@"Z:\Slack Bot\SlackBot_Auth.txt") : "https://hooks.slack.com/services/T00000000/B00000000/XXXXXXXXXXXXXXXXXXXXXXXX";
        SlackClient Bot = new SlackClient(oauth);

        public bool shouldCreate = false;
        public bool ounceEdit = false;
        //public string botPath = "";

        public Form1(string[] Args)
        {
            InitializeComponent();
            if (Args.Length > 0)
            {
                if (Args[0].ToLower() == "upload")
                    shouldCreate = true;
                else if (Args[0].ToLower() == "ounce")
                    ounceEdit = true;
            }
        }

        private static void Extract(string nameSpace, string outDirectory, string internalFilePath, string resourceName)
        {
            var assembly = Assembly.GetCallingAssembly();
            using (var s = assembly.GetManifestResourceStream(nameSpace + "." + (internalFilePath == "" ? "" : internalFilePath + ".") + resourceName))
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
            /*
            if (!File.Exists("INIFileParser.dll"))
                ExtractFile("INIFileParser.dll");

            if (!File.Exists("INIFileParser.xml"))
                ExtractFile("INIFileParser.xml");

            if (!File.Exists("Premade.ini"))
                ExtractFile("Premade.ini");

            if (!File.Exists("Template_36.xlsx"))
                ExtractFile("Template_36.xlsx");

            if (!File.Exists("Template_40.xlsx"))
                ExtractFile("Template_40.xlsx");

             * Originally this was used to just use todays or yesterdays list, but I changed it so that they will backup the lists on said days,
             * but it overwrites one universal one and will automatically use that.
             * That way theres less chance of using a very old menu.
             *
            string PremadeString = "";
            string yesterday = DateTime.Now.AddDays(-1).ToString("MM-dd-yyyy");
            Console.WriteLine(yesterday);
            Console.WriteLine(DateAndTime.Now.ToString("MM-dd-yyyy"));
            if (File.Exists("Premade" + DateAndTime.Now.ToString("MM-dd-yyyy") + ".ini"))
                PremadeString = "From Today.";
            else if (File.Exists("Premade" + yesterday + ".ini"))
                PremadeString = "From Yesterday.";
            else
                PremadeString = "No recent list, defaulting to an old one.";

            DialogResult dialogResult = MessageBox.Show("There's a premade flower list, load that for a base? " + PremadeString, "??", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                Console.WriteLine(yesterday);
                if (File.Exists("Premade" + DateAndTime.Now.ToString("MM-dd-yyyy") + ".ini"))
                    dataGridView1.DataSource = DataTable("Premade" + DateAndTime.Now.ToString("MM-dd-yyyy") + ".ini");
                else if (File.Exists("Premade" + yesterday + ".ini"))
                    dataGridView1.DataSource = DataTable("Premade" + yesterday + ".ini"); //LoadPremade("Premade" + yesterday.ToString("MM-dd-yyyy") + ".ini");
                else
                   dataGridView1.DataSource = DataTable("Premade.ini");
            }
            */
            dataGridView1.DataSource = DataTable("Premade.ini");

            if (File.Exists(TemplateFile))
                Template_Label.Text = "Template File Loaded!";

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
            Color color;
            try
            {
                color = ColorTranslator.FromHtml("#" + Hex);
            }
            catch
            {
                color = Color.White;
            }
            return color;
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
            Background_Box.Text = Background_Box.Text.ToUpper();
            AdjustColors();
        }

        //Restore some basic default settings. I think these settings are old and outdated anyways.
        private void button3_Click(object sender, EventArgs e)
        {
            Bar1_Box.Text = Bar1_Color;
            Bar2_Box.Text = Bar2_Color;
            Background_Box.Text = Background_Color;

            Indica_Box.Text = "1f4e78";
            Sativa_Box.Text = "FF0000";
            Hybrid_Box.Text = "375623";
            Heavy_Box.Text = "ed7d31";

            Range1.Text = "A1";
            Range2.Text = "L30";

            Image_Width.Value = 4893;
            Image_Height.Value = 2706;

            Tax.Value = 0;
            FontSize.Value = 28;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            doProcess(); 
        }

        private void doProcess()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                button4.Text = "Processing...";
                CreateExcel();
                EditExcel_Table();
                CreateImage();
                ResizeImage("menu_Small.png");
                button4.Text = "Create Menu";
                var Premade = "Premade" + DateTime.Today.ToString("MM-dd-yyyy") + ".ini";
                GenPremade(Premade);
                if (File.Exists("Premade.ini"))
                    File.Delete("Premade.ini");

                File.Copy(Premade, "Premade.ini");
                //MessageBox.Show("Menu has been created and should be located on your desktop!", "Done!");
                if (!checkBox1.Checked)
                    UploadScreenly();
                Cursor.Current = Cursors.Default;
                if (shouldCreate)
                    Bot.PostMessage("Flower menu was created and uploaded!", channel: "#menu_updates", username: "menubot");
                else
                    MessageBox.Show("Done :)");
            }
            catch (Exception ex)
            {
                if (!shouldCreate)
                    MessageBox.Show(ex.ToString() + "\n\n\nPlease try again. Please close all other programs or restart the computer before trying again.");
                else
                    Bot.PostMessage("Flower menu creation failed! Please try again, or try manually.", channel: "#menu_updates", username: "menubot");
                button4.Text = "Create Menu";
            }

            if (shouldCreate)
                this.Close();
        }

        public string webURL = "http://192.168.1.210/manage/shares/Server/Menus/MenuCreator/Uploads/";
        public void UploadScreenly()
        {
            try
            {
                webURL = @"/home/pi/screenly_assets/AUTOMATED_" + DateTime.Now.ToString("MM-dd-yyyy_hhmm");
                string Output =
                    Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) +
                    "\\MenuImages\\Menu_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";

                Console.WriteLine(Output);
                Console.WriteLine(webURL);
                Console.WriteLine("Deleting all old assets");
                DeleteOldAssetsAsync("192.168.1.112");
                Console.WriteLine("Uploading to Screenly Menus");

                SFTPUpload(Output, webURL, "192.168.1.112");
                Upload(webURL, "192.168.1.112");
            }
            catch (Exception ex)
            {
                if(!shouldCreate)
                 MessageBox.Show(ex.ToString());
            }
        }

        private Asset AssetToUpdate { get; set; }
        private async void Upload(string fileLocation, string IP)
        {
            Device newDevice = new Device
            {
                Name = "Flower",
                Location = "Floor",
                IpAddress = IP,
                Port = "80",
                ApiVersion = "v1.1/"
            };

            Asset a = new Asset();
            a.AssetId = "AUTOMATED_" + DateTime.Now.ToString("MM-dd-yyyy_hhmm");
            a.Name = "Menu_AUTO_" + DateTime.Now.ToString("MM-dd-yyyy hh-mm tt");
            a.Uri = fileLocation;
            a.StartDate = DateTime.Today.AddDays(-1).ToUniversalTime();
            a.EndDate = DateTime.Today.AddDays(20).ToUniversalTime();
            a.Duration = "10";
            a.IsEnabled = 1;
            a.NoCache = 0;
            a.Mimetype = "image";
            a.SkipAssetCheck = 1;
            a.IsProcessing = 0;
            await newDevice.CreateAsset(a);
        }

        public void SFTPUpload(string fileToUpload, string fileLocation, string host = "192.168.1.112", string user = "pi", string password = "raspberry", int Port = 22)
        {
            try
            {
                var client = new SftpClient(host, Port, user, password);
                client.Connect();
                if (client.IsConnected)
                    using (var fileStream = new FileStream(fileToUpload, FileMode.Open))
                    {
                        client.UploadFile(fileStream, fileLocation);
                        client.Disconnect();
                        client.Dispose();
                    }
                else
                    Console.WriteLine("Couldn't connect to host");
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        }

        public async void DeleteOldAssetsAsync(string IP)
        {
            Device newDevice = new Device
            {
                Name = "Specials",
                Location = "Floor",
                IpAddress = IP,
                Port = "80",
                ApiVersion = "v1.1/"
            };

            await newDevice.GetAssetsAsync();
            foreach (Asset asset in newDevice.ActiveAssets)
            {
                Console.WriteLine(asset.AssetId);
                await newDevice.RemoveAssetAsync(asset.AssetId);
            }
        }

        public void EditExcel_Table()
        {
            try
            {
                Console.WriteLine("Starting Excel edit...");
                int menuOffset = (Int32)menuStart.Value;
                Excel.Application excel = new Excel.Application();
                Excel.Workbook wkb = excel.Workbooks.Open(Path.Combine(StartupPath, "Generated.xlsx"));
                //Excel.Workbook wkb = excel.Workbooks.Open("Generated.xlsx");
                Excel.Worksheet sheet = wkb.Worksheets[1] as Excel.Worksheet;
                Console.WriteLine("Opened Excel file, starting edit...");
                //B-E4 - B-E21 First Row
                //H-K4 - H-K21 Second Row

                //.Range["A1:L33"]
                int Letter = 3;
                int Num = menuOffset;
                Range row = sheet.Rows.Cells[2, 1];

                #region Create lists from DataGridView

                List<string> Values = new List<string>();

                DataTable sam = (DataTable)dataGridView1.DataSource;
                DataView dt1 = new DataView(sam);
                //dt1.Sort = "Numbers ASC";
                dt1.Sort = "Cost ASC";

                DataTable dt = dt1.ToTable();
                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[0].ToString());
                }

                string[] Names = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[1].ToString());
                }

                string[] Types = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[2].ToString());
                }

                string[] Costs = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[3].ToString());
                }

                string[] THCs = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[4].ToString());
                }

                string[] CBDs = Values.ToArray();
                Values.Clear();
                /*
                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[5].ToString());
                string[] Numbers = Values.ToArray();
                Values.Clear();
                */

                #endregion Create lists from DataGridView

                Console.WriteLine(GetNum(Range2.Text) + Range2.Text.Substring(1));
                Range rng = sheet.UsedRange;
                //rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(GetColor(Background_Box.Text));
                int i = 0;
                bool ContinueAnyways = false;
                foreach (var name in Names)
                {
                    if (name != "" && name != "/n" && name != "/r")
                    {
                        if (Num >= ((TemplateMax / 2) + menuOffset) && Letter != 9)
                        {
                            Letter = 9;
                            Num = menuOffset;
                        }
                        if (Num >= (TemplateMax + menuOffset) || ContinueAnyways)
                        {
                            DialogResult dialogResult = MessageBox.Show("You have more flowers than what this template supports continue?", "???", MessageBoxButtons.YesNo);
                            if (dialogResult == DialogResult.No)
                            {
                                break;
                            }

                            ContinueAnyways = true;
                        }

                        Console.WriteLine(Letter);
                        Console.WriteLine(GetLetter(Letter));
                        Console.WriteLine(GetLetter(Letter) + Num);
                        string Color = DetermineColor(Types[i]);

                        //Flower Name
                        row = sheet.Rows.Cells[Num, Letter];
                        //row.Value = (i + 1) + ": " + name;
                        row.Value = name;
                        row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        row.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        row.Font.Size = FontSize.Value;
                        row.Font.Bold = true;
                        row.Font.Color = GetColor(Color);
                        if (i % 2 == 0)
                        {
                            row.Interior.Color = System.Drawing.ColorTranslator.ToOle(GetColor(Bar1_Box.Text));
                        }
                        else
                        {
                            row.Interior.Color = System.Drawing.ColorTranslator.ToOle(GetColor(Bar2_Box.Text));
                        }

                        //Flower Cost
                        row = sheet.Rows.Cells[Num, Letter + 1];

                        #region Tax Calculation

                        decimal Cost_ = 0;
                        //MessageBox.Show(Costs[i]);
                        if (Costs[i].Substring(0, 1) == "0")
                        {
                            Cost_ = Int32.Parse(Costs[i].Substring(1));
                        }
                        else
                        {
                            Cost_ = Int32.Parse(Costs[i]);
                        }

                        decimal percent = (Tax.Value / 100);
                        decimal test = Cost_ * percent;
                        Cost_ = Cost_ + test;

                        #endregion Tax Calculation

                        row.Value = Cost_;
                        row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        row.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        row.Font.Size = FontSize.Value;
                        row.Font.Bold = true;
                        row.Font.Color = GetColor(Color);
                        if (i % 2 == 0)
                        {
                            row.Interior.Color = System.Drawing.ColorTranslator.ToOle(GetColor(Bar1_Box.Text));
                        }
                        else
                        {
                            row.Interior.Color = System.Drawing.ColorTranslator.ToOle(GetColor(Bar2_Box.Text));
                        }

                        //Flower THC%
                        row = sheet.Rows.Cells[Num, Letter + 2];
                        row.Value = THCs[i] + "%";
                        row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        row.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        row.Font.Size = FontSize.Value;
                        row.Font.Bold = true;
                        row.Font.Color = GetColor(Color);
                        if (i % 2 == 0)
                        {
                            row.Interior.Color = System.Drawing.ColorTranslator.ToOle(GetColor(Bar1_Box.Text));
                        }
                        else
                        {
                            row.Interior.Color = System.Drawing.ColorTranslator.ToOle(GetColor(Bar2_Box.Text));
                        }

                        //Flower CBD%
                        row = sheet.Rows.Cells[Num, Letter + 3];
                        row.Value = CBDs[i] + "%";
                        row.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        row.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        row.Font.Size = FontSize.Value;
                        row.Font.Bold = true;
                        row.Font.Color = GetColor(Color);
                        if (i % 2 == 0)
                        {
                            row.Interior.Color = System.Drawing.ColorTranslator.ToOle(GetColor(Bar1_Box.Text));
                        }
                        else
                        {
                            row.Interior.Color = System.Drawing.ColorTranslator.ToOle(GetColor(Bar2_Box.Text));
                        }

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
            catch (Exception e) { MessageBox.Show(e.ToString()); }
        }

        public string DetermineColor(string Type)
        {
            Type = Type.ToUpper();
            Console.WriteLine(Type);
            if (Type.Contains("INDICA"))
            {
                return Indica_Box.Text;
            }

            if (Type.Contains("SATIVA"))
            {
                return Sativa_Box.Text;
            }

            if (Type.Contains("CBD"))
            {
                return Heavy_Box.Text;
            }

            if (Type.Contains("HYBRID"))
            {
                return Hybrid_Box.Text;
            }
            else
            {
                try
                {
                    var parser = new FileIniDataParser();
                    var data = parser.ReadFile(AppForm.StartupPath + @"\Settings.ini");
                    return data["StrainColors"][Type];
                }
                catch
                {
                    MessageBox.Show("No Color found for strain: " + Type);
                    return Hybrid_Box.Text;
                }
            }
        }

        #region Weird Excel Letter to Number functions that really arent needed but im lazy

        private const int ColumnBase = 26;
        private const int DigitMax = 7; // ceil(log26(Int32.Max))
        private const string Digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static string GetLetter(int index)
        {
            if (index <= 0)
            {
                throw new IndexOutOfRangeException("index must be a positive number");
            }

            if (index <= ColumnBase)
            {
                return Digits[index - 1].ToString();
            }

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

        #endregion Weird Excel Letter to Number functions that really arent needed but im lazy

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
            catch (Exception e){ MessageBox.Show(e.ToString()); Console.WriteLine(e.ToString()); }
            File.Copy(TemplateFile, "Generated.xlsx");
        }

        public void CreateImage()
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
            ImageHelper.sizeCheck = sizeCheck.Checked;
            using (Image image = Image.FromFile(fileName))
            {
                using (Bitmap resizedImage = ImageHelper.ResizeImage(image, 2m))
                {
                    resizedImage.Save(Environment.GetFolderPath
                   (Environment.SpecialFolder.DesktopDirectory) + "\\MenuImages\\" + "Menu_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png", ImageFormat.Png);
                }
            }
            Console.WriteLine("Done");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var Premade = "Premade" + DateTime.Today.ToString("MM-dd-yyyy") + ".ini";
            GenPremade(Premade);
        }

        //Generates ini file that the program uses for keeping track and editing the Strains
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
            {
                Values.Add(DataRow[0].ToString());
            }

            string[] Names = Values.ToArray();
            Values.Clear();

            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[1].ToString());
            }

            string[] Types = Values.ToArray();
            Values.Clear();

            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[2].ToString());
            }

            string[] Costs = Values.ToArray();
            Values.Clear();

            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[3].ToString());
            }

            string[] THCs = Values.ToArray();
            Values.Clear();

            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[4].ToString());
            }

            string[] CBDs = Values.ToArray();
            Values.Clear();
            /*
            foreach (DataRow DataRow in dt.Rows)
                Values.Add(DataRow[5].ToString());
            string[] Numbers = Values.ToArray();
            Values.Clear();
            */

            #endregion Create lists from DataGridView

            int i = 1;
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
                    //data["Numbers"][i.ToString()] = Numbers[i - 1];
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

            data["Settings"]["menuStart"] = menuStart.Value.ToString();

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
            this.BringToFront();
        }

        //Fills DataGrid from ini files and will also update settings as needed.
        public DataTable DataTable(string filePath)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Flowers");
            dt.Columns.Add("Types");
            dt.Columns.Add("Cost");
            dt.Columns.Add("THC");
            dt.Columns.Add("CBD");
            //dt.Columns.Add("Numbers");
            Console.WriteLine("Attempting to read Premade");
            var parser = new FileIniDataParser();
            var data = parser.ReadFile(filePath);
            try
            {
                int Total = Int32.Parse(data["Settings"]["Total"]);
                for (int i = 0; i < Total; i++)
                {
                    DataRow dr = dt.NewRow();
                    //dr["Numbers"] = data["Numbers"][(i + 1).ToString()];
                    dr["Flowers"] = data["Flower"][(i + 1).ToString()];
                    dr["Types"] = data["Types"][(i + 1).ToString()];

                    //Big Yikes. Supossed to help filter out unwated zeroes preceeding the Costs
                    if (int.Parse(data["Cost"][(i + 1).ToString()]) < 10 &&
                        data["Cost"][(i + 1).ToString()].Substring(0, 1) != "0" &&
                        !data["Cost"][(i + 1).ToString()].ToString().Contains("00"))
                    {
                        dr["Cost"] = "0" + data["Cost"][(i + 1).ToString()];
                    }
                    else
                    {
                        dr["Cost"] = data["Cost"][(i + 1).ToString()];
                    }

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
            Indica_Box.Text = data["Settings"]["Indica"];
            Sativa_Box.Text = data["Settings"]["Sativa"];
            Hybrid_Box.Text = data["Settings"]["Hybrid"];
            Heavy_Box.Text = data["Settings"]["CBD"];
            Range1.Text = data["Settings"]["Range1"] ?? "";
            Range2.Text = data["Settings"]["Range2"] ?? "";
            Image_Width.Value = decimal.Parse(data["Settings"]["Width"] ?? "0");
            Image_Height.Value = decimal.Parse(data["Settings"]["Height"] ?? "0");
            Tax.Value = decimal.Parse(data["Settings"]["Tax"] ?? "0");
            FontSize.Value = decimal.Parse(data["Settings"]["Font"] ?? "0");
            menuStart.Value = decimal.Parse(data["Settings"]["menuStart"] ?? "0");
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
            {
                Values.Add(DataRow[0].ToString());
            }

            string[] Names = Values.ToArray();

            Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
            this.BringToFront();
            if (shouldCreate)
                doProcess();
            if (ounceEdit)
            {
                Ounces.OunceForm frm = new Ounces.OunceForm();
                frm.ShouldCreate = true;
                frm.FormClosed += EndProgram;
                frm.Show();
            }
        }

        private void EndProgram(object sender, EventArgs e)
        {
            this.Close();
        }

        //Delete row from DataGrid based on the comboBox above the button
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                #region Create lists from DataGridView

                List<string> Values = new List<string>();
                DataTable dt = (DataTable)dataGridView1.DataSource;
                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[0].ToString());
                }

                string[] Names = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[1].ToString());
                }

                string[] Types = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[2].ToString());
                }

                string[] Costs = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[3].ToString());
                }

                string[] THCs = Values.ToArray();
                Values.Clear();

                foreach (DataRow DataRow in dt.Rows)
                {
                    Values.Add(DataRow[4].ToString());
                }

                string[] CBDs = Values.ToArray();
                Values.Clear();

                /*
                foreach (DataRow DataRow in dt.Rows)
                    Values.Add(DataRow[5].ToString());
                string[] Numbers = Values.ToArray();
                Values.Clear();
                */

                #endregion Create lists from DataGridView

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

                /*
                list = new List<string>(Numbers);
                list.RemoveAt(i);
                Numbers = list.ToArray();
                */

                dt.Rows.RemoveAt(i);
                dataGridView1.DataSource = dt;
                Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error removing the selected strain!" + System.Environment.NewLine + ex.ToString());
            }
            UpdateCount();
        }

        public void UpdateCount()
        {
            List<string> Values = new List<string>();
            DataTable dt = (DataTable)dataGridView1.DataSource;
            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[0].ToString());
            }

            string[] Names = Values.ToArray();
            FlowerCount.Text = "Flower Count: " + Names.Count();
            Values.Clear();

            Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
        }

        public int TemplateMax = 40;

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
                string s = TemplateFile.Substring(TemplateFile.LastIndexOf("Template_"));
                Template_Label.Text = "Loaded, " + TemplateFile;
                TemplateMax = Int32.Parse(s.Replace("Template_", ""));
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
            UpdateCount();
            CountCheck();
        }

        public void CountCheck()
        {
            List<string> Values = new List<string>();
            DataTable dt = (DataTable)dataGridView1.DataSource;
            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[0].ToString());
            }

            string[] Names = Values.ToArray();

            FlowerCount.Text = "Flower Count: " + Names.Count();
            if (Names.Count() > TemplateMax)
            {
                DialogResult dialogResult = MessageBox.Show("You have more flowers than what this template supports. Would you like to swap to a larger template?", "???", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    OpenTemplate();
                    if (TemplateMax < Names.Count())
                    {
                        CountCheck();
                    }
                }
            }
        }

        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            CountCheck();
            List<string> Values = new List<string>();

            DataTable sam = (DataTable)dataGridView1.DataSource;
            DataView dt1 = new DataView(sam);
            dt1.Sort = "Cost ASC";
            DataTable dt = dt1.ToTable();

            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[0].ToString());
            }

            string[] Names = Values.ToArray();
            Values.Clear();

            Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
            this.BringToFront();
            dataGridView1.DataSourceChanged -= dataGridView1_DataSourceChanged;
            dataGridView1.DataSource = dt;
            dataGridView1.DataSourceChanged += dataGridView1_DataSourceChanged;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Process.Start(@"https://www.w3schools.com/colors/colors_picker.asp");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dataGridView1.DataSource;
            dt.NewRow();
            dataGridView1.DataSource = dt;

            List<string> Values = new List<string>();
            foreach (DataRow DataRow in dt.Rows)
            {
                Values.Add(DataRow[0].ToString());
            }

            string[] Names = Values.ToArray();
            Values.Clear();

            Delete_Box.DataSource = Names.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
        }

        //Values used for the resize class
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

        public bool Export_1080p;

        //Takes the 'Names' column and turns it into a list and prints it as an excel file.
        private List<string> parts = new List<string>();

        private void button8_Click(object sender, EventArgs e)
        {
            parts.Clear();
            try
            {
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    process.Kill();
                    process.WaitForExit();
                }
                Console.WriteLine("Turning menu into nice and easy list");
                DataTable sam = (DataTable)dataGridView1.DataSource;
                DataView dt1 = new DataView(sam);
                dt1.Sort = "Cost ASC";
                DataTable dt = dt1.ToTable();

                foreach (DataRow DataRow in dt.Rows)
                {
                    parts.Add(DataRow[0].ToString() ?? "");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            PrintableExcel();
        }

        private void SendToPrinter(string File)
        {
            try
            {
                var info = new ProcessStartInfo();
                info.Verb = "print";
                info.FileName = File;
                info.CreateNoWindow = true;
                info.WindowStyle = ProcessWindowStyle.Hidden;

                var p = new Process();
                p.StartInfo = info;
                p.Start();

                p.WaitForInputIdle();
            }
            catch { }
        }

        private void PrintableExcel()
        {
            Console.WriteLine("Menu scanned creating printable document");
            if (File.Exists("PrintMenu.xlsx"))
            {
                File.Delete("PrintMenu.xlsx");
            }

            File.Copy("Blank.xlsx", "PrintMenu.xlsx");
            string excelFile = AppForm.StartupPath + "\\PrintMenu.xlsx";
            Excel.Application excel = new Excel.Application();
            Workbook w = excel.Workbooks.Open(excelFile);
            Worksheet ws = w.Sheets[1];
            ws.Protect(Contents: false);
            int i = 0;
            int k = 1;
            foreach (var value in parts)
            {
                i++;
                if (i >= 45)
                {
                    i = 1;
                    k++;
                }
                ws.Cells[i, k].Value = value;
                ws.Cells[i, k].Font.Size = 9;
                ws.Cells[i, k].Font.Bold = true;
            }
            w.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
            Console.WriteLine("Printing to default printer");
            SendToPrinter(excelFile);
        }

        private bool IsDigitsOnly(object str)
        {
            if (str != null)
            {
                foreach (char c in str.ToString())
                {
                    if (c < '0' || c > '9')
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        public int GetNum(string Range)
        {
            char c = char.Parse(Range.Substring(0, 1).ToLower());
            return char.ToUpper(c) - 63;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (File.Exists("MenuCreator.exe"))
            {
                Process.Start("MenuCreator.exe");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Process.Start(TemplateFile);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Ounces.OunceForm frm = new Ounces.OunceForm();
            frm.TemplateFile = TemplateFile;
            frm.ShowDialog();
        }

        private void Image_Width_ValueChanged(object sender, EventArgs e)
        {

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
        public static bool Export_1080p;
        public static bool sizeCheck;

        public static Bitmap ResizeImage(Image image, decimal percentage)
        {
            int width = 1920;
            int height = 1080;
            string[] args = { };
            f1 = new Form1(args);
            if (!f1.shouldCreate)
            {
                if (sizeCheck) { Export_1080p = false; }
                else
                {
                    DialogResult dialogResult = MessageBox.Show("Do you want to export with your custom size(typically 4k)? Only say no if you plan on making a GIF or video.", "Export options", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        Export_1080p = false;
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        Export_1080p = true;
                    }
                }
            }
            else
                Export_1080p = false;

            if (!Export_1080p)
            {
                width = Decimal.ToInt32(f1
                    .Final_Width); //(int)Math.Round(image.Width * percentage, MidpointRounding.AwayFromZero);
                height =
                    Decimal.ToInt32(f1
                        .Final_Height); //(int)Math.Round(image.Height * percentage, MidpointRounding.AwayFromZero);
            }

            return ResizeImage(image, width, height);
        }
    }


    public class Asset
    {
        [Newtonsoft.Json.JsonProperty(PropertyName = "asset_id")]
        public string AssetId { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "mimetype")]
        public string Mimetype { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "name")]
        public string Name { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "end_date")]
        public DateTime EndDate { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "is_enabled")]
        public Int32 IsEnabled { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "is_processing")]
        public Int32? IsProcessing { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "skip_asset_check")]
        public Int32 SkipAssetCheck { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public bool IsEnabledSwitch
        {
            get
            {
                return IsEnabled.Equals(1) ? true : false;
            }
        }

        [Newtonsoft.Json.JsonProperty(PropertyName = "nocache")]
        public Int32 NoCache { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "is_active")]
        public Int32 IsActive { get; set; }

        private string _Uri;

        [Newtonsoft.Json.JsonProperty(PropertyName = "uri")]
        public string Uri
        {
            get { return _Uri; }
            set { _Uri = System.Net.WebUtility.UrlEncode(value); }
        }

        [Newtonsoft.Json.JsonIgnore]
        public string ReadableUri
        {
            get
            {
                return System.Net.WebUtility.UrlDecode(this.Uri);
            }
        }

        [Newtonsoft.Json.JsonProperty(PropertyName = "duration")]
        public string Duration { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "play_order")]
        public Int32 PlayOrder { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "start_date")]
        public DateTime StartDate { get; set; }
    }
    public class Device
    {
        [Newtonsoft.Json.JsonIgnore]
        private List<Asset> Assets;

        [Newtonsoft.Json.JsonIgnore]
        public bool IsUp { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public ObservableCollection<Asset> ActiveAssets
        {
            get
            {
                return new ObservableCollection<Asset>(this.Assets.FindAll(x => x.IsActive.Equals(1)));
            }
        }

        [Newtonsoft.Json.JsonIgnore]
        public ObservableCollection<Asset> InactiveAssets
        {
            get
            {
                return new ObservableCollection<Asset>(this.Assets.FindAll(x => x.IsActive.Equals(0)));
            }
        }

        [Newtonsoft.Json.JsonProperty(PropertyName = "name")]
        public string Name { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "ip_address")]
        public string IpAddress { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "port")]
        public string Port { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "location")]
        public string Location { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "api_version")]
        public string ApiVersion { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public string HttpLink
        {
            get
            {
                return $"http://{IpAddress}:{Port}";
            }
        }

        public Device()
        {
            this.Assets = new List<Asset>();
            this.IsUp = false;
        }

        public async Task<bool> IsReachable()
        {
            try
            {
                HttpClient client = new HttpClient();
                client.Timeout = new TimeSpan(0, 0, 1);

                HttpResponseMessage response = await client.GetAsync(this.HttpLink);
                if (response == null || !response.IsSuccessStatusCode)
                {
                    this.IsUp = false;
                    return false;
                }
                else
                {
                    this.IsUp = true;
                    return true;
                }
            }
            catch
            {
                this.IsUp = false;
                return false;
            }
        }


        #region Screenly's API methods

        /// <summary>
        /// Get assets trought Screenly API
        /// </summary>
        /// <returns></returns>
        public async Task GetAssetsAsync()
        {
            List<Asset> returnedAssets = new List<Asset>();
            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets";

            try
            {
                HttpClient request = new HttpClient();
                using (HttpResponseMessage response = await request.GetAsync(this.HttpLink + parameters))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }

                if (!resultJson.Equals(string.Empty))
                    this.Assets = JsonConvert.DeserializeObject<List<Asset>>(resultJson);
            }
            catch (Exception ex)
            {
                throw new Exception("Error while getting assets.", ex);
            }
        }

        /// <summary>
        /// Remove specific asset for selected device
        /// </summary>
        /// <param name="assetId">Asset ID</param>
        /// <returns>Boolean for result of execution</returns>
        public async Task<bool> RemoveAssetAsync(string assetId)
        {
            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets/{assetId}";

            try
            {
                HttpClient request = new HttpClient();
                using (HttpResponseMessage response = await request.DeleteAsync(this.HttpLink + parameters))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error when asset deleting.", ex);
            }

            return true;
        }

        /// <summary>
        /// Update specific asset
        /// </summary>
        /// <param name="a">Asset to update</param>
        /// <returns>Asset updated</returns>
        public async Task<Asset> UpdateAssetAsync(Asset a)
        {
            Asset returnedAsset = new Asset();
            JsonSerializerSettings settings = new JsonSerializerSettings();
            IsoDateTimeConverter dateConverter = new IsoDateTimeConverter
            {
                DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss.fff'Z'"
            };
            settings.Converters.Add(dateConverter);

            string json = JsonConvert.SerializeObject(a, settings);
            var postData = $"model={json}";
            var data = System.Text.Encoding.UTF8.GetBytes(postData);

            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets/{a.AssetId}";

            try
            {
                HttpClient client = new HttpClient();
                HttpContent content = new ByteArrayContent(data, 0, data.Length);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");
                using (HttpResponseMessage response = await client.PutAsync(this.HttpLink + parameters, content))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }

                if (!resultJson.Equals(string.Empty))
                {
                    returnedAsset = JsonConvert.DeserializeObject<Asset>(resultJson, settings);
                }
            }
            catch (WebException ex)
            {
                using (var stream = ex.Response.GetResponseStream())
                using (var reader = new StreamReader(stream))
                {
                    throw new Exception(reader.ReadToEnd(), ex);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error while updating asset.", ex);
            }

            return returnedAsset;
        }

        /// <summary>
        /// Update order of active assets throught API
        /// </summary>
        /// <param name="newOrder"></param>
        /// <returns></returns>
        public async Task UpdateOrderAssetsAsync(string newOrder)
        {
            var postData = $"ids={newOrder}";
            var data = System.Text.Encoding.UTF8.GetBytes(postData);

            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets/order";

            try
            {
                HttpClient client = new HttpClient();
                HttpContent content = new ByteArrayContent(data, 0, data.Length);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");
                using (HttpResponseMessage response = await client.PostAsync(this.HttpLink + parameters, content))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }
            }
            catch (WebException ex)
            {
                using (var stream = ex.Response.GetResponseStream())
                using (var reader = new StreamReader(stream))
                {
                    throw new Exception(reader.ReadToEnd(), ex);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error while updating assets order.", ex);
            }
        }

        /// <summary>
        /// Create new asset on Raspberry using API
        /// </summary>
        /// <param name="a">New asset to create on Raspberry</param>
        /// <returns></returns>
        public async Task CreateAsset(Asset a)
        {
            Asset returnedAsset = new Asset();
            JsonSerializerSettings settings = new JsonSerializerSettings();
            IsoDateTimeConverter dateConverter = new IsoDateTimeConverter
            {
                DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss.fff'Z'"
            };
            settings.Converters.Add(dateConverter);

            string json = JsonConvert.SerializeObject(a, settings);
            var postData = $"model={json}";
            var data = System.Text.Encoding.UTF8.GetBytes(postData);

            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets";

            try
            {
                HttpClient client = new HttpClient();
                HttpContent content = new ByteArrayContent(data, 0, data.Length);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");

                using (HttpResponseMessage response = await client.PostAsync(this.HttpLink + parameters, content))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }

                if (!resultJson.Equals(string.Empty))
                    returnedAsset = JsonConvert.DeserializeObject<Asset>(resultJson, settings);
            }
            catch (WebException ex)
            {
                using (var stream = ex.Response.GetResponseStream())
                using (var reader = new StreamReader(stream))
                {
                    throw new Exception(reader.ReadToEnd(), ex);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error while creating asset.", ex);
            }
        }

        /// <summary>
        /// Return asset identified by asset ID in param API
        /// </summary>
        /// <param name="assetId">Asset ID to find on device</param>
        /// <returns></returns>
        public async Task<Asset> GetAssetAsync(string assetId)
        {
            Asset returnedAsset = new Asset();
            JsonSerializerSettings settings = new JsonSerializerSettings();
            IsoDateTimeConverter dateConverter = new IsoDateTimeConverter
            {
                DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss.fff'Z'"
            };
            settings.Converters.Add(dateConverter);

            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets/{assetId}";

            try
            {
                HttpClient request = new HttpClient();
                using (HttpResponseMessage response = await request.GetAsync(this.HttpLink + parameters))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }

                if (!resultJson.Equals(string.Empty))
                    return JsonConvert.DeserializeObject<Asset>(resultJson);
            }
            catch (Exception ex)
            {
                throw new Exception("Error while getting assets.", ex);
            }
            return null;
        }

        #endregion
    }
}