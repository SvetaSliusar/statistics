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
            public string fullFilePath;
        }

        DicomData[] firstDataSet;

        private void button1_Click(object sender, EventArgs e)
        {
            label2.Visible = false;
            label3.Visible = false;
            FolderBrowserDialog browserDialog = new FolderBrowserDialog();
            if (browserDialog.ShowDialog() == DialogResult.OK)
            {
                string dir = browserDialog.SelectedPath;
                string[] fullfilesPath = Directory.GetFiles(dir, "*.dcm", SearchOption.AllDirectories);
                progressBar1.Maximum = fullfilesPath.Length;
                firstDataSet = LoadFiles(fullfilesPath);
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
                progressBar1.Value = curImg;
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
            progressBar1.Value = 0;
            label2.Visible = true;
            return DicomFile;
        }

        private void ExportToExcel(DicomData[] dicomDatas, XLWorkbook workBook)
        {            
            IXLWorksheet workSheet;
            workBook.TryGetWorksheet(methodName.SelectedItem.ToString(), out workSheet);
            if (workSheet == null)
                workSheet = workBook.Worksheets.Add(methodName.SelectedItem.ToString());
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
            try
            {
                workBook.Save();
            }
            catch
            {
                workBook.SaveAs(String.Format("D://{0}//{0}.xlsx",
                dicomDatas[0].fileName.Substring(0, 12)));
                progressBar2.Value = 0;
                label3.Visible = true;
            }                     
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
