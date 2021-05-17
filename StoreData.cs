using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NationalInstruments.Examples.ContAcqVoltageSamples_IntClk
{
    class StoreData
    {
        List<double[]> Data = new List<double[]>();

        public void createCSVfile(int[] rowCount, double[] time, double[][] data)
        {
            try
            {
                IWorkbook wb = new XSSFWorkbook();
                ISheet ws = wb.CreateSheet("Class");
                string datetime = DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss");
                ////建立Excel 2007檔案
                //IWorkbook wb = new XSSFWorkbook();
                //ISheet ws = wb.CreateSheet("Class");

                ws.CreateRow(0);//第一行為欄位名稱
                ws.GetRow(0).CreateCell(0).SetCellValue("name");
                ws.GetRow(0).CreateCell(1).SetCellValue("score");

                //Data.Add(data);

                foreach (var rowIndex in rowCount)
                {
                    ws.CreateRow(Convert.ToInt32(rowIndex));                
                }
                for (int j = 0; j < data.Length; j++)
                {
                    for (int i = 0; i < data[0].Length; i++)
                    {
                        ws.GetRow(0).CreateCell(i).SetCellValue(time[i]);
                        ws.GetRow(j+1).CreateCell(i).SetCellValue(data[j][i]);
                    }
                }


                string filepath = @"C:\Data\npoi";
                FileStream file;
                if (File.Exists(filepath))
                {
                    file = new FileStream(filepath + datetime + ".csv", FileMode.Create);//產生檔案
                    wb.Write(file);
                    file.Close();
                }
                else
                {
                    file = new FileStream(filepath + datetime + ".csv", FileMode.Create);//產生檔案
                    wb.Write(file);
                    file.Close();
                }

            }
            catch (Exception)
            {

                throw;
            }
        
        }

    }
}
