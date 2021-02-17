using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ExcelEPPlus
{
    public class ExcelService
    {
        public void ExportToExcel<T>(Stream stream, List<T> model)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; 
            //取得首行文字
            var excelTitle = GetPropertyInfos<T>();
            using (var xlPackage = new ExcelPackage(stream))
            {
                
                //設定工作表
                var worksheet = xlPackage.Workbook.Worksheets.Add("Sheet1");

                for (int i = 0; i < excelTitle.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = excelTitle[i].DisplayName;
                    worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                }
                int row = 2;
                foreach (T item in model.ToList())
                {
                    for (int i = 1; i <= excelTitle.Count; i++)
                    {
                        worksheet.Cells[row, i].Value = item.GetType().GetProperty(excelTitle[i - 1].Name).GetValue(item, null);
                        if (item.GetType().GetProperty(excelTitle[i - 1].Name).PropertyType.IsValueType)
                        {
                            worksheet.Cells[row, i].Style.Numberformat.Format = "#0";
                        }

                    }
                    row++;
                }
                xlPackage.Save();
            }
        }

        /// <summary>
        /// 從模型取出屬性相關內容
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        List<ExcelTitleModel> GetPropertyInfos<T>()
        {
            PropertyInfo[] vPropertyInfos = typeof(T).GetProperties();
            var excelTitle = new List<ExcelTitleModel>();
            foreach (PropertyInfo Item in vPropertyInfos)
            {
                var ignore = Item.CustomAttributes.Where(x => x.AttributeType == typeof(ExportIgnoreAttribute)).Any();
                if (ignore)
                {
                    continue;
                }
                var titleModel = new ExcelTitleModel();
                titleModel.Name = Item.Name;
                var order = Item.CustomAttributes.Where(x => x.AttributeType == typeof(OrderByAttribute)).Single();
                if (order != null)
                {
                    titleModel.Orderby = Convert.ToInt16(order.ConstructorArguments[0].Value);
                }
                var display = Item.CustomAttributes.Where(x => x.AttributeType == typeof(DisplayNameAttribute)).Single();
                if (display != null)
                {
                    titleModel.DisplayName = Item.GetCustomAttribute<DisplayNameAttribute>().DisplayName;
                }

                excelTitle.Add(titleModel);
            }
            excelTitle = excelTitle.OrderBy(o => o.Orderby).ToList();
            return excelTitle;
        }
    }
}
