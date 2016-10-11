using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlusDemo
{
    class ExcelHelper
    {
        private static readonly ExcelHelper excelHelper = new ExcelHelper();
        private ExcelHelper() { }

        public static ExcelHelper CreateInstance()
        {
            return excelHelper;
        }

        /// <summary>
        /// 从Excel读取数据
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public List<Student> Import(string fileName)
        {
            var students = new List<Student>();
            var excel = new FileInfo(fileName);
            var package = new ExcelPackage(excel);
            var worksheet = package.Workbook.Worksheets[2];//选定 指定页
            int maxColumnNum = worksheet.Dimension.End.Column;//最大列
            int minColumnNum = worksheet.Dimension.Start.Column;//最小列

            int maxRowNum = worksheet.Dimension.End.Row;//最小行
            int minRowNum = worksheet.Dimension.Start.Row;//最大行

            for (int i = 2; i <= maxRowNum; i++)
            {
                var student = new Student()
                {
                    Name = worksheet.Cells[i, 1].Value.ToString(),
                    Age = Convert.ToInt32(worksheet.Cells[i, 2].Value),
                    Sex = worksheet.Cells[i, 3].Value.ToString(),
                    Course = worksheet.Cells[i, 4].Value.ToString()
                };
                students.Add(student);
            }
            return students;
        }

        /// <summary>
        /// 把数据导出到Excel
        /// </summary>
        /// <param name="fileName"></param>
        public void Export(string fileName)
        {
            var sex = new string[] { "Boy", "Girl" };
            var course = new string[] { "C", "C++", "C#", "Java" };

            var file = new FileInfo(fileName);
            var package = new ExcelPackage(file);
            // 数据源表单  
            var source = package.Workbook.Worksheets.Add("Source");
            source.Cells[1, 1].Style.Font.Bold = true;
            source.Cells[2, 1].Style.Font.Bold = true;
            source.Cells[1, 1].Value = "Sex";
            source.Cells[2, 1].Value = "Course";
            for (int i = 0; i < sex.Length; i++)
            {
                source.Cells[1, 2 + i].Value = sex[i];
            }
            for (int i = 0; i < course.Length; i++)
            {
                source.Cells[2, 2 + i].Value = course[i];
            }

            //添加数据源
            package.Workbook.Names.Add("sex", source.Cells[1, 2, 1, 3]);
            package.Workbook.Names.Add("course", source.Cells[2, 2, 2, 5]);

            //下拉列表的示例表单
            var student = package.Workbook.Worksheets.Add("Student");
            student.Cells["A1"].Style.Font.Bold = true;
            student.Cells["B1"].Style.Font.Bold = true;
            student.Cells["B1"].Style.Font.Bold = true;
            student.Cells["D1"].Style.Font.Bold = true;
            student.Cells["A1"].Value = "Name";
            student.Cells["B1"].Value = "Age";
            student.Cells["C1"].Value = "Sex";
            student.Cells["D1"].Value = "Course";

            //数据有效性或者数据验证  
            //性别  
            var validationSex = student.DataValidations.AddListValidation("C2:C65535");
            validationSex.ShowErrorMessage = true;
            validationSex.ErrorStyle = ExcelDataValidationWarningStyle.warning;
            validationSex.ErrorTitle = "Error";
            validationSex.Error = "输入的值无效";
            validationSex.Formula.ExcelFormula = "sex"; //绑定数据源
                                                        //课程  
            var validationCourse = student.DataValidations.AddListValidation("D2:D65535");
            validationCourse.ShowErrorMessage = true;
            validationCourse.ErrorStyle = ExcelDataValidationWarningStyle.warning;
            validationCourse.ErrorTitle = "Error";
            validationCourse.Error = "输入的值无效";
            validationCourse.Formula.ExcelFormula = "course"; //绑定数据源

            //填充Excel
            for (int i = 0; i < 20; i++)
            {
                student.Cells[2 + i, 1].Value = $"student{i + 1}";
                student.Cells[2 + i, 2].Value = 10 + i % 5;
                student.Cells[2 + i, 3].Value = sex[i % sex.Length];
                student.Cells[2 + i, 4].Value = course[i % course.Length];
            }

            package.SaveAs(file);
        }
    }
}
