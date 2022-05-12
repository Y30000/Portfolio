using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Timers;
using System.Linq;

namespace Cal
{
    public partial class IDIROM : Form
    {
        int workDayCount = 0;
        int totalWorkDayCount = 0;
        int workThreadCount = 0;
        private static System.Timers.Timer aTimer;
        
        void ProgresseBar(Object source, ElapsedEventArgs e)
        {
            UpdateTextBox(workDayCount + " / " + totalWorkDayCount , Convert.ToInt32((double)workDayCount / totalWorkDayCount * progressBar1.Maximum));

            if (workThreadCount == 0)
            {
                aTimer.Stop();
                UpdateTextBox("Done" , progressBar1.Maximum);
            }
        }

        private void UpdateTextBox(string data, int progressValue)
        {
            // 호출한 쓰레드가 작업쓰레드인가?
            if (textBox_Progress.InvokeRequired)
            {
                // 작업쓰레드인 경우
                textBox_Progress.BeginInvoke(new Action(() => textBox_Progress.Text = data));
                progressBar1.BeginInvoke(new Action(() => progressBar1.Value = progressValue));
            }
            else
            {
                // UI 쓰레드인 경우
                textBox_Progress.Text = data;
                progressBar1.Value = progressValue;
            }
        }

        public IDIROM()
        {
            InitializeComponent();
            //folderPathTextBox.Text = System.IO.Directory.GetCurrentDirectory();
            comboBox_Formula.SelectedIndex = comboBox_Region.SelectedIndex = 0;
            textBox_Progress.Text = "Stay";

            aTimer = new System.Timers.Timer(100);
            aTimer.Elapsed += ProgresseBar;
            aTimer.Stop();
        }
        /*
        Excel.Application excelAppData = null;
        Excel.Workbook wbData = null;
        Excel.Worksheet wsData = null;
        Excel.Application excelAppResult = null;
        Excel.Workbook wbResult = null;
        Excel.Worksheet wsResult1 = null;
        Excel.Worksheet wsResult2 = null;
        */
        private void IWR_Calculation(object sender, EventArgs e)
        {
            workDayCount = 0;
            workThreadCount++;

            if (string.Empty == openFileDialog1.FileName)
            {
                Button_OpenFile_Click(sender, e);
                return;
            }
            aTimer.Start();
            foreach (var fileName in openFileDialog1.FileNames)
            {
                _ = ThreadPool.QueueUserWorkItem(DoIWR, fileName);
            }
            workThreadCount--;
        }

        private void DoIWR(object fileName)
        {
            workThreadCount++;
            IWR cIWR = new IWR();

            Excel.Application excelAppData = new Excel.Application() { DisplayAlerts = false };
            Excel.Workbook wbData = excelAppData.Workbooks.Open(fileName.ToString());
            Excel.Worksheet wsData = wbData.Worksheets.get_Item(1);


            Excel.Worksheet wsResultDaily;

            try
            {
                wsResultDaily = wbData.Worksheets["Daily IWR"];
            }
            catch
            {
                wsResultDaily = wbData.Worksheets.Add(After: wbData.Worksheets[wbData.Worksheets.Count]);
                wsResultDaily.Name = "Daily IWR";
            }

            Excel.Worksheet wsResultYearly;

            try
            {
                wsResultYearly = wbData.Worksheets["Yearly IWR"];
            }
            catch
            {
                wsResultYearly = wbData.Worksheets.Add(After: wbData.Worksheets[wbData.Worksheets.Count]);
                wsResultYearly.Name = "Yearly IWR";
            }

            OpenExcelForIWR(wsData, wsResultDaily, wsResultYearly, cIWR);
            
            wbData.Save();
            wbData.Close(true);
            excelAppData.Quit();
            ReleaseExcelObject(wsResultYearly);
            ReleaseExcelObject(wsResultDaily);
            ReleaseExcelObject(wsData);
            ReleaseExcelObject(wbData);
            ReleaseExcelObject(excelAppData);
            workThreadCount--;
        }

        private void OpenExcelForIWR(Excel.Worksheet wsData, Excel.Worksheet wsDaily, Excel.Worksheet wsYearly, IWR cIWR)
        {

            // Excel 첫번째 워크시트 가져오기                


            /*
            wb = excelApp.Workbooks.Add();
            ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;
            */
            //(System.IO.Directory.GetCurrentDirectory() + @"\IWR\ttest.xlsx");


            object[,] data = wsData.UsedRange.Value;
            int length = 2;
            for (int i = 2 ; i < data.GetLength(0) && null != data[i, 1]; ++i)
            {
                ++totalWorkDayCount;
                ++length;
            }

            object[,] newDaily = new object[length, 10];
            object[,] newYearly = new object[length / 354 + 2, 7];

            cIWR.Calculate_Init(Convert.ToString(data[2, 1]));

            cIWR.SetPaddyFieldData
                (
                textBox_Infiltration,
                textBox_MaxPondingDepth,
                textBox_MinPondingDepth,
                textBox_TransplantingWater,
                textBox_RiceNurseryWater,
                textBox_NurseryAreaRate
                );

            cIWR.SetCroppingPeriods
                (
                maskedTextBox_NurseryPreparationPeriod,
                maskedTextBox_NurseryPeriod,
                maskedTextBox_NurseryAndTransplantingPeriod,
                maskedTextBox_TransplantingAndGrowingPeriod,
                maskedTextBox_GrowingPeriod
                );

            cIWR.SetClimateData(
                new TextBox[]{
                    IWRTextBox42,
                    IWRTextBox43,
                    IWRTextBox51,
                    IWRTextBox52,
                    IWRTextBox53,
                    IWRTextBox61,
                    IWRTextBox62,
                    IWRTextBox63,
                    IWRTextBox71,
                    IWRTextBox72,
                    IWRTextBox73,
                    IWRTextBox81,
                    IWRTextBox82,
                    IWRTextBox83,
                    IWRTextBox91,
                    IWRTextBox92,
                    IWRTextBox93
                });

            cIWR.Calculate_SetCultivationArea(Convert.ToDouble(data[2, 7]));
            /*
            if(data == null || data[1, 1].GetType() != typeof(Double) || data[1,2] != null || data[1,3].GetType() != typeof(string) || data[1, 4].GetType() != typeof(string))
            {
                return;
            }
            */
            SaveWsInitForIWR(newDaily, false);
            SaveWsInitForIWR(newYearly, true);

            
            var sum = new double[9];

            for (int i = 0; i < 9; i++)
            {
                sum[i] = 0;
            }

            for (int i = 2, j = 2; i < data.GetLength(0) && null != data[i, 1]; ++i)
            {
                cIWR.Calculate_SetCurrentDatesAndGetResult(Convert.ToString(data[i,1]) ,Convert.ToInt32(data[i, 2]), Convert.ToInt32(data[i, 3]), Convert.ToDouble(data[i, 4]), Convert.ToDouble(data[i, 5]), out double[] results);

                newDaily[i-1, (int) NameOfRowForIWR.Year] = data[i, 1];
                newDaily[i-1, (int) NameOfRowForIWR.Month] = results[0];
                newDaily[i-1, (int) NameOfRowForIWR.Date] = results[1];
                newDaily[i-1, (int) NameOfRowForIWR.Evapotranspiration] = results[2];
                newDaily[i-1, (int) NameOfRowForIWR.ConsumptiveUse] = results[3];
                newDaily[i-1, (int) NameOfRowForIWR.Precipitation] = results[4];
                newDaily[i-1, (int) NameOfRowForIWR.PondingDepth] = results[5];
                newDaily[i-1, (int) NameOfRowForIWR.EffectiveRainfall] = results[6];
                newDaily[i-1, (int) NameOfRowForIWR.IrrigationWaterRequirement] = results[7];
                newDaily[i-1, (int) NameOfRowForIWR.NetDutyOfWater] = results[8];

                for (int k = 0; k < sum.Length; ++k)
                {
                    sum[k] += results[k];
                }

                if (!(i < data.GetLength(0) && data[i, 1].Equals(data[i + 1, 1])))
                {

                    int offset1 = 2, offset2 = 3;
                    newYearly[j-1, (int) NameOfRowForIWR.Year] = data[i, 1];
                    //wsSummary.Cells[j, (int) NameOfRow.Month] = "월";
                    //wsSummary.Cells[j, (int) NameOfRow.Date] = "일";
                    newYearly[j-1, (int) NameOfRowForIWR.Evapotranspiration - offset1] = sum[2];
                    newYearly[j-1, (int) NameOfRowForIWR.ConsumptiveUse - offset1] = sum[3];
                    newYearly[j-1, (int) NameOfRowForIWR.Precipitation - offset1] = sum[4];
                    //wsSummary.Cells[j, (int) NameOfRow.PondingDepth] = average[5];
                    newYearly[j-1, (int) NameOfRowForIWR.EffectiveRainfall - offset2] = sum[6];
                    newYearly[j-1, (int) NameOfRowForIWR.IrrigationWaterRequirement - offset2] = sum[7];
                    newYearly[j++ -1, (int) NameOfRowForIWR.NetDutyOfWater - offset2] = sum[8];

                    if (i + 1 < data.GetLength(0) && null != data[i + 1, 1])
                    {
                        cIWR.Calculate_ChangeYear(Convert.ToString(data[i + 1 , 1]));
                    }

                    Array.Clear(sum, 0, sum.Length);
                }
                workDayCount++;
                //textBox_Progress.Text = data[i, 1].ToString() + ((int)results[0]).ToString("D2") + ((int)results[1]).ToString("D2");
                //textBox_Progress.Text = string.Format("IWR - {0:0000}{1:00}{2:00}", data[i, 1], results[0], results[1]);
            }
            //ABCDE FGHIJ KLMNO PQRST UVWXYZ
            wsDaily.get_Range("A1", "J" + newDaily.GetLength(0)).Value = newDaily;
            wsYearly.get_Range("A1", "G" + newYearly.GetLength(0)).Value = newYearly;
            //textBox_Progress.Text = "Done";
        }

        enum NameOfRowForIWR { Year = 0, Month, Date, Evapotranspiration, ConsumptiveUse, Precipitation, PondingDepth, EffectiveRainfall, IrrigationWaterRequirement, NetDutyOfWater }
        private void SaveWsInitForIWR(object[,] obj, bool isYeary)
        {
            if (isYeary)
            {
                int offset = -2;
                obj[0, (int) NameOfRowForIWR.Year] = "년 (Year)";
                //월
                //일
                obj[0, (int) NameOfRowForIWR.Evapotranspiration + offset] = "실제 증발산량 (ETc)(mm)";
                obj[0, (int) NameOfRowForIWR.ConsumptiveUse + offset] = "소비수량 (Consumptive Use)(mm)";
                obj[0, (int) NameOfRowForIWR.Precipitation + offset--] = "강우량 (Precipitation)(mm)";
                //담수량
                obj[0, (int) NameOfRowForIWR.EffectiveRainfall + offset] = "유효우량 (Effective rainfall)(mm)";
                obj[0, (int) NameOfRowForIWR.IrrigationWaterRequirement + offset] = "수요량 (IWR)(mm)";
                obj[0, (int) NameOfRowForIWR.NetDutyOfWater + offset] = "수요량 (IWR)(10^3 m^3)";
                return;
            }
            obj[0, (int) NameOfRowForIWR.Year] = "년 (Year)";
            obj[0, (int) NameOfRowForIWR.Month] = "월 (Monteith)";
            obj[0, (int) NameOfRowForIWR.Date] = "일 (Day)";
            obj[0, (int) NameOfRowForIWR.Evapotranspiration] = "실제 증발산량 (ETc)(mm)";
            obj[0, (int) NameOfRowForIWR.ConsumptiveUse] = "소비수량 (Consumptive Use)(mm)";
            obj[0, (int) NameOfRowForIWR.Precipitation] = "강우량 (Precipitation)(mm)";
            obj[0, (int) NameOfRowForIWR.PondingDepth] = "담수심 (Ponding depth)(mm)";
            obj[0, (int) NameOfRowForIWR.EffectiveRainfall] = "유효우량 (Effective rainfall)(mm)";
            obj[0, (int) NameOfRowForIWR.IrrigationWaterRequirement] = "수요량 (IWR)(mm)";
            obj[0, (int) NameOfRowForIWR.NetDutyOfWater] = "수요량 (IWR)(10^3 m^3)";
        }

        private static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }



        private void CroppingPeriodsResetButton_Click(object sender, EventArgs e)
        {
            string[] identityValue;
            switch (comboBox_Region.SelectedIndex)
            {

                case 0:
                    identityValue = new string[]
                    {"04170420"
                    ,"04210520"
                    ,"05210531"
                    ,"06010610"
                    ,"06110910" };
                    break;
                case 1:
                    identityValue = new string[]
                    {"04270430"
                    ,"05010530"
                    ,"05310610"
                    ,"06110620"
                    ,"06210920" };
                    break; ;
                default:
                    return;
            }

            maskedTextBox_NurseryPreparationPeriod.Text = identityValue[0];
            maskedTextBox_NurseryPeriod.Text = identityValue[1];
            maskedTextBox_NurseryAndTransplantingPeriod.Text = identityValue[2];
            maskedTextBox_TransplantingAndGrowingPeriod.Text = identityValue[3];
            maskedTextBox_GrowingPeriod.Text = identityValue[4];

            if (beforeDir == comboBox_Region.SelectedIndex)
            {
                return;
            }
            beforeDir = comboBox_Region.SelectedIndex;

            CoefficientShifter(comboBox_Region.SelectedIndex);
        }

        private void PabbyFieldDataResetButton_Click(object sender, EventArgs e)
        {
            string[] identityValue =
                {"5"
                ,"80"
                ,"20"
                ,"140"
                ,"5" };

            textBox_Infiltration.Text = identityValue[0];
            textBox_MaxPondingDepth.Text = identityValue[1];
            textBox_MinPondingDepth.Text = identityValue[2];
            textBox_RiceNurseryWater.Text = textBox_TransplantingWater.Text = identityValue[3];
            textBox_NurseryAreaRate.Text = identityValue[4];
        }

        private void ClimateDataResetButton_Click(object sender, EventArgs e)
        {
            string[] identityValue;
            switch (comboBox_Formula.SelectedIndex)
            {
                case 0:
                    identityValue = new string[]
                    {"0.56"
                    ,"0.56"
                    ,"0.56"
                    ,"0.56"
                    ,"0.56"
                    ,"0.75"
                    ,"0.95"
                    ,"1.06"
                    ,"1.09"
                    ,"1.17"
                    ,"1.39"
                    ,"1.53"
                    ,"1.58"
                    ,"1.47"
                    ,"1.42"
                    ,"1.32" };
                    break;
                case 1:
                    identityValue = new string[]
                    {"0.97"
                    ,"0.97"
                    ,"0.97"
                    ,"0.97"
                    ,"0.97"
                    ,"0.97"
                    ,"0.97"
                    ,"0.97"
                    ,"1.15"
                    ,"1.15"
                    ,"1.15"
                    ,"1.34"
                    ,"1.34"
                    ,"1.34"
                    ,"1.34"
                    ,"1.34" };
                    break; ;
                default:
                    return;
            }

            TextBox[] textBoxes =
            {IWRTextBox42
            ,IWRTextBox43
            ,IWRTextBox51
            ,IWRTextBox52
            ,IWRTextBox53
            ,IWRTextBox61
            ,IWRTextBox62
            ,IWRTextBox63
            ,IWRTextBox71
            ,IWRTextBox72
            ,IWRTextBox73
            ,IWRTextBox81
            ,IWRTextBox82
            ,IWRTextBox83
            ,IWRTextBox91
            ,IWRTextBox92
            ,IWRTextBox93};

            int regionSelectedIndex = comboBox_Region.SelectedIndex;

            if (regionSelectedIndex == 0)
            {
                textBoxes[textBoxes.Length - 1].Text = "0";
            }
            else
            {
                textBoxes[0].Text = "0";
            }

            for (int i = regionSelectedIndex, j = 0; j < identityValue.Length; ++i, ++j)
            {
                textBoxes[i].Text = identityValue[j];
            }
        }

        private int beforeDir = 0;
        private void CoefficientShifter(int dir)
        {

            TextBox[] textBoxes =
            {IWRTextBox42
            ,IWRTextBox43
            ,IWRTextBox51
            ,IWRTextBox52
            ,IWRTextBox53
            ,IWRTextBox61
            ,IWRTextBox62
            ,IWRTextBox63
            ,IWRTextBox71
            ,IWRTextBox72
            ,IWRTextBox73
            ,IWRTextBox81
            ,IWRTextBox82
            ,IWRTextBox83
            ,IWRTextBox91
            ,IWRTextBox92
            ,IWRTextBox93};

            switch (dir)
            {
                case 0:
                    for (int i = 0; i < textBoxes.Length - 1; ++i)
                    {
                        textBoxes[i].Text = textBoxes[i + 1].Text;
                    }
                    textBoxes[textBoxes.Length - 1].Text = "0";
                    break;
                case 1:
                    for (int i = textBoxes.Length - 1; i > 0; --i)
                    {
                        textBoxes[i].Text = textBoxes[i - 1].Text;
                    }
                    textBoxes[0].Text = "0";
                    break;
                default:
                    return;
            }
        }

        private void Button_OpenFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            openFileDialog1.Filter = "Excel File|*.xlsx";
            openFileDialog1.Title = "Select Data Files";
            openFileDialog1.Multiselect = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox_FileNames.Clear();
                foreach (var name in openFileDialog1.FileNames)
                {
                    string fileName = name.Substring(name.LastIndexOf("\\")+1);
                    textBox_FileNames.AppendText(fileName + Environment.NewLine);
                }
            }
        }

        //List<string> applicationNames;

        private void Idirom_FormClosing(object sender, FormClosingEventArgs e)
        {
            var processes = from p in System.Diagnostics.Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.MainWindowTitle == "")
                    process.Kill();
            }

            /*
            foreach(var name in applicationNames)
            {
                KillSpecificExcelFileProcess(name);
            }
            */
            /*
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach(var app in process)
            {
                app.Kill();
            }*/
            /*
            System.Diagnostics.Process[] mProcess = System.Diagnostics.Process.GetProcessesByName(Application.ProductName);
            foreach (System.Diagnostics.Process p in mProcess)
                p.Kill();
                */
            //System.Diagnostics.Process.GetCurrentProcess().Kill();
        }
        /*
        private void KillSpecificExcelFileProcess(string excelFileName)
        {
            var processes = from p in System.Diagnostics.Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.MainWindowTitle == "Microsoft Excel - " + excelFileName)
                    process.Kill();
            }
        }
        */
        private void Button_INFLOWExecute_Click(object sender, EventArgs e)
        {
            workDayCount = 0;
            workThreadCount++;
            aTimer.Start();

            if (string.Empty == openFileDialog1.FileName)
            {
                Button_OpenFile_Click(sender, e);
                return;
            }

            foreach (var fileName in openFileDialog1.FileNames)
            {
                _ = ThreadPool.QueueUserWorkItem(DoINFLOW, fileName);
            }
            workThreadCount--;
        }

        private void DoINFLOW(object fileName)
        {
            workThreadCount++;
            INFLOW iNFLOW = new INFLOW();

            iNFLOW.Calculate_Init();

            Excel.Application excelAppData = new Excel.Application() { DisplayAlerts = false };
            Excel.Workbook wbData = excelAppData.Workbooks.Open(fileName.ToString());
            Excel.Worksheet wsData = wbData.Worksheets.get_Item(1);

            Excel.Worksheet wsResultDaily;

            try
            {
                wsResultDaily = wbData.Worksheets["Daily INFLOW"];
            }
            catch
            {
                wsResultDaily = wbData.Worksheets.Add(After: wbData.Worksheets[wbData.Worksheets.Count]);
                wsResultDaily.Name = "Daily INFLOW";
            }

            Excel.Worksheet wsResultYearly;

            try
            {
                wsResultYearly = wbData.Worksheets["Yearly INFLOW"];
            }
            catch
            {
                wsResultYearly = wbData.Worksheets.Add(After: wbData.Worksheets[wbData.Worksheets.Count]);
                wsResultYearly.Name = "Yearly INFLOW";
            }

            OpenExcelForINFLOW(wsData, wsResultDaily, wsResultYearly, iNFLOW);

            wbData.Save();
            wbData.Close(true);
            excelAppData.Quit();
            ReleaseExcelObject(wsResultYearly);
            ReleaseExcelObject(wsResultDaily);
            ReleaseExcelObject(wsData);
            ReleaseExcelObject(wbData);
            ReleaseExcelObject(excelAppData);
            workThreadCount--;
        }

        private void OpenExcelForINFLOW(Excel.Worksheet wsData, Excel.Worksheet wsDaily, Excel.Worksheet wsYearly, INFLOW iNFLOW)
        {
            object[,] data = wsData.UsedRange.Value;

            int length = 2;
            for (int i = 2; i < data.GetLength(0) && null != data[i, 1]; ++i)
            {
                ++totalWorkDayCount;
                ++length;
            }

            object[,] newDaily = new object[length, 13];
            object[,] newYearly = new object[length / 365 + 2, 11];

            /*  검문소 다시 만들것
            if (data == null || data[1, 1].GetType() != typeof(Double) || data[1, 2] != null || data[1, 3].GetType() != typeof(string) || data[1, 4].GetType() != typeof(string))
            {
                wbData.Close(true);
                excelAppData.Quit();
                return;
            }
            */

            SaveWsInitForINFLOW(newDaily, false);
            SaveWsInitForINFLOW(newYearly, true);

            iNFLOW.InitInputData(new double[]
            {
                Convert.ToDouble(data[2,6]) / 100,
                Convert.ToDouble(data[4,6]),
                Convert.ToDouble(data[4,7]),
                Convert.ToDouble(data[4,8])
            },
            new TextBox[]{
                INFLOWTextBox011,
                INFLOWTextBox012,
                INFLOWTextBox013,
                INFLOWTextBox021,
                INFLOWTextBox022,
                INFLOWTextBox023,
                INFLOWTextBox031,
                INFLOWTextBox032,
                INFLOWTextBox033,
                INFLOWTextBox041,
                INFLOWTextBox042,
                INFLOWTextBox043,
                INFLOWTextBox051,
                INFLOWTextBox052,
                INFLOWTextBox053,
                INFLOWTextBox061,
                INFLOWTextBox062,
                INFLOWTextBox063,
                INFLOWTextBox071,
                INFLOWTextBox072,
                INFLOWTextBox073,
                INFLOWTextBox081,
                INFLOWTextBox082,
                INFLOWTextBox083,
                INFLOWTextBox091,
                INFLOWTextBox092,
                INFLOWTextBox093,
                INFLOWTextBox101,
                INFLOWTextBox102,
                INFLOWTextBox103,
                INFLOWTextBox111,
                INFLOWTextBox112,
                INFLOWTextBox113,
                INFLOWTextBox121,
                INFLOWTextBox122,
                INFLOWTextBox123
            });

            bool isFirst = true;
            double area = Convert.ToDouble(data[2, 6]) / 100;
            double[] sum = new double[10];
            Array.Clear(sum, 0, sum.Length);

            for (int i = 2, j = 2; i < data.GetLength(0) && null != data[i, 1]; ++i)
            {
                iNFLOW.SetPrecipitationAndGetResult(Convert.ToInt32(data[i, 2]), Convert.ToInt32(data[i, 3]), Convert.ToDouble(data[i, 4]), out double[] results, isFirst);

                newDaily[i-1, (int) NameOfRowForINFLOW.Year] = data[i, 1];
                newDaily[i-1, (int) NameOfRowForINFLOW.Month] = data[i, 2];
                newDaily[i-1, (int) NameOfRowForINFLOW.Date] = data[i, 3];
                newDaily[i-1, (int) NameOfRowForINFLOW.Precipitation] = results[1];
                newDaily[i-1, (int) NameOfRowForINFLOW.Infiltrtion1] = results[5];
                newDaily[i-1, (int) NameOfRowForINFLOW.Infiltrtion2] = results[9];
                newDaily[i-1, (int) NameOfRowForINFLOW.Infiltrtion3] = results[13];
                newDaily[i-1, (int) NameOfRowForINFLOW.Runoff11] = results[2];
                newDaily[i-1, (int) NameOfRowForINFLOW.Runoff12] = results[3];
                newDaily[i-1, (int) NameOfRowForINFLOW.Runoff2] = results[7];
                newDaily[i-1, (int) NameOfRowForINFLOW.Runoff3] = results[11];
                double inflow = results[2] + results[3] + results[7] + results[11];
                newDaily[i-1, (int) NameOfRowForINFLOW.Inflow] = inflow;
                newDaily[i-1, (int) NameOfRowForINFLOW.Inflowm3] = inflow * area;

                for(int q = 0; q < sum.Length; ++q)
                {
                    sum[q] += Convert.ToDouble(newDaily[i-1, q + 3]);
                }

                if ( !(i < data.GetLength(0) && data[i, 1].Equals(data[i + 1, 1])))
                {
                    newYearly[j-1, (int) NameOfRowForINFLOW.Year] = data[i, 1];

                    for (int k = 0; k < sum.Length; ++k)
                    {
                        newYearly[j-1, k + 1] = sum[k];
                    }
                    j++;
                    Array.Clear(sum, 0, sum.Length);
                }

                if (isFirst) { isFirst = false; }
                workDayCount++;
                //textBox_Progress.Text = string.Format("INFLOW - {0:0000}{1:00}{2:00}", data[i, 1], data[i, 2], data[i, 3]);
            }
            //ABCDE FGHIJ KLMNO PQRST UVWXYZ
            wsDaily.get_Range("A1", "M" + newDaily.GetLength(0)).Value = newDaily;
            wsYearly.get_Range("A1", "K" + newYearly.GetLength(0)).Value = newYearly;
            //textBox_Progress.Text = "Done";
        }

        enum NameOfRowForINFLOW
        {
            Year = 0,
            Month,
            Date,
            Precipitation,
            Infiltrtion1,
            Infiltrtion2,
            Infiltrtion3,
            Runoff11,
            Runoff12,
            Runoff2,
            Runoff3,
            Inflow,
            Inflowm3
        }
        void SaveWsInitForINFLOW(object[,] obj, bool isYear)
        {
            int offset = 0;
            if (isYear)
            {
                offset = -2;
            }
            obj[0, (int) NameOfRowForINFLOW.Year] = "년 (Year)";
            obj[0, (int) NameOfRowForINFLOW.Month] = "월 (Monteith)";
            obj[0, (int) NameOfRowForINFLOW.Date] = "일 (Day)";
            obj[0, offset + (int) NameOfRowForINFLOW.Precipitation] = "강우량(Precipitation)";
            obj[0, offset + (int) NameOfRowForINFLOW.Infiltrtion1] = "침투량1(mm)(Infiltrtion1)";
            obj[0, offset + (int) NameOfRowForINFLOW.Infiltrtion2] = "침투량2(mm)(Infiltrtion2)";
            obj[0, offset + (int) NameOfRowForINFLOW.Infiltrtion3] = "침투량3(mm)(Infiltrtion3)";
            obj[0, offset + (int) NameOfRowForINFLOW.Runoff11] = "유출량11(mm)(Runoff11)";
            obj[0, offset + (int) NameOfRowForINFLOW.Runoff12] = "유출량12(mm)(Runoff12)";
            obj[0, offset + (int) NameOfRowForINFLOW.Runoff2] = "유출량2(mm)(Runoff2)";
            obj[0, offset + (int) NameOfRowForINFLOW.Runoff3] = "유출량3(mm)(Runoff3)";
            obj[0, offset + (int) NameOfRowForINFLOW.Inflow] = "유입량 mm (Inflow)";
            obj[0, offset + (int) NameOfRowForINFLOW.Inflowm3] = "유입량 1000m^3 (Inflow)";
        }

        private void Button_WBMExecute_Click(object sender, EventArgs e)
        {
            workDayCount = 0;
            workThreadCount++;
            aTimer.Start();

            if (string.Empty == openFileDialog1.FileName)
            {
                Button_OpenFile_Click(sender, e);
                return;
            }

            foreach (var fileName in openFileDialog1.FileNames)
            {
                _ = ThreadPool.QueueUserWorkItem(DoWBM, fileName);
            }
            workThreadCount--;
        }

        private void DoWBM(object fileName)
        {
            workThreadCount++;
            WBM wBM = new WBM();
            Excel.Application excelAppData = new Excel.Application() { DisplayAlerts = false };
            Excel.Workbook wbData = excelAppData.Workbooks.Open(fileName.ToString());
            Excel.Worksheet wsData = wbData.Worksheets.get_Item(1);
            Excel.Worksheet wsIWR;
            //////////////////////////////////  IWR     ///////////////////////////
            try
            {
                wsIWR = wbData.Worksheets.Item["Daily IWR"];
            }
            catch
            {
                IWR cIWR = new IWR();

                wsIWR = wbData.Worksheets.Add(After: wbData.Worksheets[wbData.Worksheets.Count]);
                wsIWR.Name = "Daily IWR";

                Excel.Worksheet wsIWRYearly;

                try
                {
                    wsIWRYearly = wbData.Worksheets["Yearly IWR"];
                }
                catch
                {
                    wsIWRYearly = wbData.Worksheets.Add(After: wbData.Worksheets[wbData.Worksheets.Count]);
                    wsIWRYearly.Name = "Yearly IWR";
                }

                OpenExcelForIWR(wsData, wsIWR, wsIWRYearly, cIWR);

            }


            //////////////////////////////////  INFLOW    ///////////////////////////
            Excel.Worksheet wsINFLOW;
            try
            {
                wsINFLOW = wbData.Worksheets.Item["Daily INFLOW"];
            }
            catch
            {
                INFLOW iNFLOW = new INFLOW();

                iNFLOW.Calculate_Init();

                wsINFLOW = wbData.Worksheets.Add(After: wbData.Worksheets[wbData.Worksheets.Count]);
                wsINFLOW.Name = "Daily INFLOW";

                Excel.Worksheet wsINFLOWYearly;

                try
                {
                    wsINFLOWYearly = wbData.Worksheets["Yearly INFLOW"];
                }
                catch
                {
                    wsINFLOWYearly = wbData.Worksheets.Add(After: wbData.Worksheets[wbData.Worksheets.Count]);
                    wsINFLOWYearly.Name = "Yearly INFLOW";
                }
                OpenExcelForINFLOW(wsData, wsINFLOW, wsINFLOWYearly, iNFLOW);
            }

            Excel.Worksheet wsResultDaily;
            Excel.Worksheet wsResultYearly;
            try
            {
                wsResultDaily = wbData.Worksheets["Daily Water Balance Model"];
            }
            catch
            {
                wsResultDaily = wbData.Worksheets.Add(After: wbData.Worksheets[wbData.Worksheets.Count]);
                wsResultDaily.Name = "Daily Water Balance Model";
            }

            try
            {
                wsResultYearly = wbData.Worksheets["Yearly Water Balance Model"];
            }
            catch
            {
                wsResultYearly = wbData.Worksheets.Add(After: wbData.Worksheets[wbData.Worksheets.Count]);
                wsResultYearly.Name = "Yearly Water Balance Model";
            }

            Excel.Worksheet aofdf;
            try
            {
                aofdf = wbData.Worksheets["Drought frequency"];
            }
            catch
            {
                aofdf = wbData.Worksheets.Add(After: wbData.Worksheets[wbData.Worksheets.Count]);
                aofdf.Name = "Drought frequency";
            }

            OpenExcelForWBM(wsData, wsIWR, wsINFLOW, wsResultDaily, wsResultYearly, wBM);

            OpenExcelForAOFDF(wsResultYearly, aofdf);

            wbData.Save();
            wbData.Close(true);
            excelAppData.Quit();
            ReleaseExcelObject(wsResultYearly);
            ReleaseExcelObject(wsResultDaily);
            ReleaseExcelObject(wsData);
            ReleaseExcelObject(wbData);
            ReleaseExcelObject(excelAppData);
            workThreadCount--;
        }

        void SaveWsInitForWBM(object[,] obj, bool isYeary)
        {
            int offset = 0;
            if (isYeary)
            {
                offset = -2;
            }
            obj[0, 0] = "년 (Year)";
            obj[0, 1] = "월 (Monteith)";
            obj[0, 2] = "일 (Day)";
            obj[0, offset + 3] = "유역유입량(watershed Inflow)(천m3)";
            obj[0, offset + 4] = "수면강수량(Prcp of surface)( 천m3)";
            obj[0, offset + 5] = "전체유입량(Total inflow)";
            obj[0, offset + 6] = "IWR(천m3)";
            obj[0, offset + 7] = "수면증발량(ET of surface)(천m3)";
            obj[0, offset + 8] = "전체 필요수량(Total water requirement)";
            obj[0, offset + 9] = "물수지 (Water balance) (천m3)";
            if (isYeary)
            {
                obj[0, offset + 10] = "년최대 필요저수량(Yeary maximum storage requirments of reservoir) (천m3)";
            }
            else
            {
                obj[0, offset + 10] = "필요저수량(storage requirments of reservoir) (천m3)";
            }
            obj[0, offset + 11] = "저수량(Storage) (천m3) (방류전 before the discharge)";
            obj[0, offset + 12] = "저수량(Storage) (천m3) (방류후after the discharge)";
            obj[0, offset + 13] = "방류량(Discharge) (천m3)";
        }

        private void OpenExcelForWBM(Excel.Worksheet wsData, Excel.Worksheet wsiwrData, Excel.Worksheet wsinflowData, Excel.Worksheet wsDaily, Excel.Worksheet wsYearly, WBM wBM)
        {
            object[,] data = wsData.UsedRange.Value;
            object[,] iwrData = wsiwrData.UsedRange.Value;
            object[,] inflowData = wsinflowData.UsedRange.Value;

            int length = 2;
            for (int i = 2; i < data.GetLength(0) && null != data[i, 1]; ++i)
            {
                ++totalWorkDayCount;
                ++length;
            }

            object[,] newDaily = new object[length, 14];
            object[,] newYearly = new object[length / 365 + 2, 12];

            /*  검문소 다시 만들것
            if (data == null || data[1, 1].GetType() != typeof(Double) || data[1, 2] != null || data[1, 3].GetType() != typeof(string) || data[1, 4].GetType() != typeof(string))
            {
                wbData.Close(true);
                excelAppData.Quit();
                return;
            }
            */

            SaveWsInitForWBM(newDaily, false);
            SaveWsInitForWBM(newYearly, true);


            int deadIndex = 0, fullIndex = 0;

            for (int i = 1, j = 0; i < 120 && j < 2; ++i)
            {
                if (data[i, 9] != null)
                {
                    if (data[i, 9].ToString().Equals("만수위"))
                    {
                        fullIndex = i;
                        j++;
                    }
                    else if (data[i, 9].ToString().Equals("사수위"))
                    {
                        deadIndex = i;
                        j++;
                    }
                }

                if (i == 119 && j < 2)
                {
                    return;
                }
            }

            double FWLArea = Convert.ToDouble(data[fullIndex, 7]);
            double FWLInternalCumulativeVolume = 0;

            for (int i = deadIndex; i <= fullIndex; ++i)
            {
                if (Double.TryParse(Convert.ToString(data[i - 1, 6]), out double beforeHight))
                {
                    FWLInternalCumulativeVolume += Convert.ToDouble(data[i, 8]) * (Convert.ToDouble(data[i, 6]) - beforeHight);
                }
            }

            wBM.InitCalculate(FWLArea, FWLInternalCumulativeVolume);
            double[] sum = new double[13];
            double maxWBM = 0;
            for (int i = 2, j = 2; i < data.GetLength(0) && null != data[i, 1]; ++i)
            {

                if (i > inflowData.GetLength(0) || i > iwrData.GetLength(0))
                {
                    break;
                }

                wBM.SetCurrentDatasAndGetResult(Convert.ToInt32(data[i, 2]), Convert.ToInt32(data[i, 3]), Convert.ToDouble(inflowData[i, 13]), Convert.ToDouble(data[i, 4]), Convert.ToDouble(iwrData[i, 10]), out double[] results);

                newDaily[i-1, 0] = data[i, 1];
                for (int k = 0; k < results.Length; ++k)
                {
                    newDaily[i-1, k + 1] = results[k];
                }


                if (i < data.GetLength(0) && data[i, 1].Equals(data[i + 1, 1]))
                {
                    for (int k = 0; k < sum.Length; ++k)
                    {
                        sum[k] += results[k];
                    }
                    maxWBM = Math.Max(maxWBM, results[9]);
                }
                else
                {
                    for (int k = 0; k < sum.Length; ++k)
                    {
                        sum[k] += results[k];
                    }
                    maxWBM = Math.Max(maxWBM, results[9]);

                    newYearly[j-1, (int) NameOfRowForIWR.Year] = data[i, 1];

                    for (int k = 2; k < sum.Length; ++k)
                    {
                        newYearly[j-1, k-1] = sum[k];
                    }
                    newYearly[j++ -1, 8] = maxWBM;
                    maxWBM = 0;
                    Array.Clear(sum, 0, sum.Length);
                }

                workDayCount++;
                //textBox_Progress.Text = string.Format("WBM - {0:0000}{1:00}{2:00}", data[i, 1], data[i, 2], data[i, 3]);
            }
            //ABCDE FGHIJ KLMNO PQRST UVWXYZ
            wsDaily.get_Range("A1", "N" + newDaily.GetLength(0)).Value = newDaily;
            wsYearly.get_Range("A1", "L" + newYearly.GetLength(0)).Value = newYearly;
            //textBox_Progress.Text = "Done";
        }

        private void OpenExcelForAOFDF(Excel.Worksheet wsYearly, Excel.Worksheet aofdf)
        {
            object[,] aofdfData = wsYearly.UsedRange.Value;
            List<double> datas = new List<double>();

            for (int i = aofdfData.GetLength(0) - 1; i > 0; i--)
            {
                if (null != aofdfData[i, 9])
                {
                    datas.Add(Convert.ToDouble(aofdfData[i + 1, 9]));
                }
            }

            if(0 == datas.Count)
            {
                return;
            }

            AOFDF aOFDF = new AOFDF();
            CommonClass.Pair<string, double>[] AOFDFResults = aOFDF.GetAOFDF(datas.ToArray());

            for (int i = 0; i < AOFDFResults.Length; ++i)
            {
                aofdf.Cells[i + 1, 1] = AOFDFResults[i].GetKey();
                aofdf.Cells[i + 1, 2] = AOFDFResults[i].GetValue();
            }
        }
    }
}