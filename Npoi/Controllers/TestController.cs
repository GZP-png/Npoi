using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;


namespace Npoi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TestController : ControllerBase
    {
        [HttpGet]
        public void Test()
        {
            FileStream file;
            ISheet sheet;
            IWorkbook workbook = null;
            try
            {
                string filepath = @"C:\Users\53572\Desktop\测试.xlsx";
                // 2007版本及以上版本
                if (filepath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(file);
                // 2003版本
                else if (filepath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(file);
                if (workbook != null)
                {

                    sheet = workbook.CreateSheet("sheet10");
                    for (int i = 0; i < 4; i++)
                    {
                        sheet.CreateRow(i); //创建行
                        for (int j = 0; j < 3; j++)
                        {
                            sheet.GetRow(i).CreateCell(j);//创建单元格
                            sheet.GetRow(i).GetCell(j).SetCellValue($"行{i}列{j}");
                        }
                    }
                    for (int i = 0; i < 4; i++)
                    {
                        for (int j = 0; j < 3; j++)
                        {
                            //写入值
                            sheet.GetRow(i).GetCell(j).SetCellValue($"行{i}列{j}");
                        }
                    }
                    //写入到文件中
                    file = new FileStream(filepath, FileMode.Open, FileAccess.Write);
                    workbook.Write(file);
                    file.Close();
                    workbook.Close();
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }
    }
}
