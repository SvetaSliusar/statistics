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
using EvilDICOM.Core.Helpers;
using EvilDICOM.Core.Interfaces;
using EvilDICOM.Core.IO.Writing;
using EvilDICOM.Core.Element;
using EvilDICOM.Core;
using ClosedXML.Excel;

namespace StatisticsAnalysis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        struct DicomData
        {
            public dynamic[] PixelDataOrig;
            public int sizeRow;
            public int sizeCol;
            public string fileName;
            public double sliceLoc;
            public double instNum;
            public string fullFilePath;
        }

        DicomData[] firstDataSet;
        DicomData[] secondDataSet;
        bool firstComplete = false;
        bool secondComplete = false;

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog browserDialog = new FolderBrowserDialog();
            if (browserDialog.ShowDialog() == DialogResult.OK)
            {
                string dir = browserDialog.SelectedPath;
                string[] fullfilesPath = Directory.GetFiles(dir, "*.dcm", SearchOption.AllDirectories);
                progressBar1.Maximum = fullfilesPath.Length;
                firstDataSet = LoadFiles(fullfilesPath);
                firstComplete = true;
            }
        }

        XLWorkbook workBook = new XLWorkbook();

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel(firstDataSet, workBook);
        }

        private DicomData[] LoadFiles(string[] fullfilesPath)
        {
            DICOMObject dicom;
            DicomData[] DicomFile = new DicomData[fullfilesPath.Length];
            for (int curImg = 0; curImg < fullfilesPath.Length; curImg++)
            {
                if (!firstComplete)
                    progressBar1.Value = curImg;
                else
                    progressBar2.Value = curImg;
                string file = fullfilesPath[curImg];
                dicom = DICOMObject.Read(file);

                DicomFile[curImg].sizeRow = Convert.ToInt32(dicom.FindFirst("00280010").DData);
                DicomFile[curImg].sizeCol = Convert.ToInt32(dicom.FindFirst("00280011").DData);
                DicomFile[curImg].fileName = Path.GetFileName(fullfilesPath[curImg]);
                DicomFile[curImg].fullFilePath = fullfilesPath[curImg];

                var pixelData = dicom.FindFirst("7FE00010").DData_;
                DicomFile[curImg].PixelDataOrig = new dynamic[DicomFile[curImg].sizeRow * DicomFile[curImg].sizeCol];
                var typeData = Convert.ToInt16(dicom.FindFirst("00280100").DData);
                switch (typeData)
                {
                    case 8:
                        if (pixelData.Count != DicomFile[curImg].sizeRow * DicomFile[curImg].sizeCol)
                        {
                            dynamic[] temp = new dynamic[pixelData.Count];
                            pixelData.CopyTo(temp, 0);
                            dynamic[] temp2 = new dynamic[DicomFile[curImg].sizeRow * DicomFile[curImg].sizeCol];
                            temp2 = temp.Skip(1).Take(DicomFile[curImg].sizeRow * DicomFile[curImg].sizeCol).ToArray();
                            for (int i = 0; i < temp2.Length; i++)
                                temp2[i] = temp2[i] ?? 0;
                            temp2.CopyTo(DicomFile[curImg].PixelDataOrig, 0);
                        }
                        else
                            pixelData.CopyTo(DicomFile[curImg].PixelDataOrig, 0);
                        break;
                    case 16:
                        var PixelDataArray = new ushort[DicomFile[curImg].sizeRow * DicomFile[curImg].sizeCol * 2];
                        pixelData.CopyTo(PixelDataArray, 0);
                        int pix = 0;
                        for (int i = 0; i < DicomFile[curImg].sizeRow * DicomFile[curImg].sizeCol * 2; i += 2)
                            DicomFile[curImg].PixelDataOrig[pix++] = (ushort)(PixelDataArray[i] | PixelDataArray[i + 1] << 8);
                        break;
                }
                
                var rescaleSlope = (float)Convert.ToDouble(dicom.FindFirst("00281053").DData);
                var rescaleIntercept = (float)Convert.ToDouble(dicom.FindFirst("00281052").DData);
                for (int i=0; i<DicomFile[curImg].PixelDataOrig.Length;i++)
                {
                    DicomFile[curImg].PixelDataOrig[i] = DicomFile[curImg].PixelDataOrig[i] * rescaleSlope + rescaleIntercept;
                }
            }
            return DicomFile;
        }

        private void ExportToExcel(DicomData[] dicomDatas, XLWorkbook workBook)
        { 
            var workSheet = workBook.Worksheets.Add(methodName.SelectedItem.ToString());
            //IXLWorksheet workSheet;
            //workBook.TryGetWorksheet(methodName.SelectedItem.ToString(), out workSheet);
            int columnCount = 0;
            foreach (var data in dicomDatas)
            {
                progressBar2.Value = columnCount;
                workSheet.Cell(1, ++columnCount).Value = Path.GetFileNameWithoutExtension(data.fileName);
                for (int i = 1; i < data.PixelDataOrig.Length; i++)
                {
                    workSheet.Cell(i + 1, columnCount).SetValue<dynamic>(data.PixelDataOrig[i]);
                }                
            }
            workSheet.Columns().AdjustToContents();
            //workBook.Save();
            workBook.SaveAs(String.Format("D://{0}//{0}.xlsx", 
                dicomDatas[0].fileName.Substring(0, 12)));
            //int i = 0; 

            //foreach(var dicomData in dicomDatas)
            //{
            //    i = 0;
            //    table.Columns.Add(dicomData.fileName, typeof(float));
            //    foreach(var pixelData in dicomData.PixelDataOrig)
            //    {
            //        try
            //        {
            //            table.Rows[i][dicomData.fileName] = pixelData;
            //            i++;
            //        }
            //        catch(System.IndexOutOfRangeException)
            //        {
            //            table.Rows.Add(pixelData);
            //        }
            //    }
            //}
            //workBook.Worksheets.Add(table, "BRUTEFORCE");


            //var workbook = new XLWorkbook();
            //var worksheet = workbook.Worksheets.Add("Sample Sheet");
            //worksheet.Cell("A1").Value = "Hello World!";
            //workbook.SaveAs("D://A//hello.xlsx");
            //System.Data.DataTable table = new System.Data.DataTable();
            //table.Columns.Add("First Dataset", typeof(ushort));
            //table.Columns.Add("Second Dataset", typeof(ushort));
            //table.Columns.Add("Pearson");
            //table.Columns.Add("Slope");
            //table.Columns.Add("Intercept");
            //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            //excel.Workbooks.Add();
            //excel.Visible = false;
            //excel.Range["A1"].Value = "Pearson";
            //excel.Range["A2"].Select();
            //for (int i = 0, j = 0; i < firstDataSet[0].PixelDataOrig.Length && j < secondDataSet[0].PixelDataOrig.Length; i++, j++)
            //{
            //    excel.ActiveCell.Value = firstDataSet[0].PixelDataOrig[i];
            //    excel.ActiveCell.Offset[0, 1].Value = secondDataSet[0].PixelDataOrig[j];
            //    excel.ActiveCell.Offset[1, 0].Select();
            //}

            //WorksheetFunction worksheetFunction = excel.WorksheetFunction;
            //List<double> pearsons = new List<double>();
            //for (int i = 0, j = 0; i < firstDataSet.Length && j < secondDataSet.Length; i++, j++)
            //{
            //    try
            //    {
            //        pearsons.Add(worksheetFunction.Pearson(firstDataSet[i].PixelDataOrig, secondDataSet[j].PixelDataOrig));
            //    }
            //    catch(Exception e)
            //    {

            //    }
            //        //excel.ActiveCell.Offset[1, 0].Select();
            //}
            //excel.Columns[1].AutoFit();
            //excel.Visible = true;
            //return table;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            workBook = new XLWorkbook();
        }
    }
}
