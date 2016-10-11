using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Windows;

namespace EPPlusDemo
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Import_Click(object sender, RoutedEventArgs e)
        {
            this.Import.IsEnabled = false;
            try
            {
                var students = ExcelHelper.CreateInstance().Import("Student.xlsx");

                this.dgStudent.ItemsSource = null;
                this.dgStudent.ItemsSource = students;

                this.Import.IsEnabled = true;
                MessageBox.Show("完成");
            }
            catch (Exception ex)
            {
                this.Import.IsEnabled = true;
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Export_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "Excel 2007(*.xlsx)|*.xlsx|Excel 2003(*.xls)|*.xls";
            if(!saveFile.ShowDialog().Value)
            {
                return;
            }
            this.Export.IsEnabled = false;
            try
            {
                ExcelHelper.CreateInstance().Export(saveFile.FileName);

                this.Export.IsEnabled = true;
                MessageBox.Show("完成");
            }
            catch (Exception ex)
            {
                this.Export.IsEnabled = true;
                MessageBox.Show(ex.Message);
            }
        }


    }
}
