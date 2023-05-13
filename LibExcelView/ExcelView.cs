using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace LibExcelView
{
    /// <summary>
    /// 按照步骤 1a 或 1b 操作，然后执行步骤 2 以在 XAML 文件中使用此自定义控件。
    ///
    /// 步骤 1a) 在当前项目中存在的 XAML 文件中使用该自定义控件。
    /// 将此 XmlNamespace 特性添加到要使用该特性的标记文件的根
    /// 元素中:
    ///
    ///     xmlns:MyNamespace="clr-namespace:LibExcelView"
    ///
    ///
    /// 步骤 1b) 在其他项目中存在的 XAML 文件中使用该自定义控件。
    /// 将此 XmlNamespace 特性添加到要使用该特性的标记文件的根
    /// 元素中:
    ///
    ///     xmlns:MyNamespace="clr-namespace:LibExcelView;assembly=LibExcelView"
    ///
    /// 您还需要添加一个从 XAML 文件所在的项目到此项目的项目引用，
    /// 并重新生成以避免编译错误:
    ///
    ///     在解决方案资源管理器中右击目标项目，然后依次单击
    ///     “添加引用”->“项目”->[浏览查找并选择此项目]
    ///
    ///
    /// 步骤 2)
    /// 继续操作并在 XAML 文件中使用控件。
    ///
    ///     <MyNamespace:ExcelView/>
    ///
    /// </summary>
    public class ExcelView : Control
    {
        static ExcelView()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(ExcelView), new FrameworkPropertyMetadata(typeof(ExcelView)));
#if DEBUG
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; //指明商业应用
#endif
        }

        private ExcelPackage? loadedPackage;
        private Dictionary<string, ExcelWorksheet>? worksheetMappings;

        private void LoadExcel()
        {
            if (string.IsNullOrWhiteSpace(FileName))
            {
                ClearExcel();
                return;
            }

            if (!File.Exists(FileName))
                return;

            try
            {
                loadedPackage = new ExcelPackage(FileName);
                worksheetMappings = loadedPackage.Workbook.Worksheets
                    .ToDictionary(v => v.Name);
                SheetNames = loadedPackage.Workbook.Worksheets
                    .Select(sheet => sheet.Name)
                    .ToArray();

                if (string.IsNullOrWhiteSpace(SheetName))
                    SheetName = loadedPackage.Workbook.Worksheets
                        .Select(sheet => sheet.Name)
                        .FirstOrDefault();
                        
                ErrorMessage = null;
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
            }
        }

        private void LoadSheet()
        {
            if (string.IsNullOrWhiteSpace(SheetName))
            {
                ClearExcel();
                return;
            }

            if (worksheetMappings == null)
                return;

            if (worksheetMappings.TryGetValue(SheetName, out ExcelWorksheet? sheet))
            {
                RenderExcel(sheet);
            }
        }

        private void ClearExcel()
        {
            Data = null;
            loadedPackage?.Dispose();
        }

        private void RenderExcel(ExcelWorksheet sheet)
        {
            if (sheet.Cells.Value is not object[,] rs)
            {
                Data = null;
                return;
            }

            var rowLen = rs.GetLength(0);
            var colLen = rs.GetLength(1);

            var data = new DataTable();

            data.Columns.Add(
                new DataColumn("\\", typeof(string)) 
                {
                    ReadOnly = true, 
                });

            foreach (string header in Utils.EnumerateLetters().Take(colLen))
                data.Columns.Add(header);

            for (int row = 0; row < rowLen; row++)
            {
                var rowData = data.NewRow();
                rowData[0] = $"{row + 1}";
                data.Rows.Add(rowData);
            }

            for (int row = 0; row < rowLen; row++)
            {
                var rowData = data.Rows[row];
                for (int col = 0; col < colLen; col++)
                {
                    rowData[col + 1] = rs[row, col];
                }
            }

            Data = new DataView(data);
        }


        public object? Data
        {
            get { return (object)GetValue(DataProperty); }
            set { SetValue(DataProperty, value); }
        }

        public string? FileName
        {
            get { return (string?)GetValue(FileNameProperty); }
            set { SetValue(FileNameProperty, value); }
        }

        public string? SheetName
        {
            get { return (string)GetValue(SheetNameProperty); }
            set { SetValue(SheetNameProperty, value); }
        }

        public string? ErrorMessage
        {
            get { return (string?)GetValue(ErrorMessageProperty); }
            set { SetValue(ErrorMessageProperty, value); }
        }

        public string[]? SheetNames
        {
            get { return (string[])GetValue(SheetNamesProperty); }
            private set { SetValue(SheetNamesProperty, value); }
        }





        public static readonly DependencyProperty FileNameProperty =
            DependencyProperty.Register("FileName", typeof(string), typeof(ExcelView), new PropertyMetadata(null, ExcelFileNameChanged));

        public static readonly DependencyProperty SheetNameProperty =
            DependencyProperty.Register("SheetName", typeof(string), typeof(ExcelView), new PropertyMetadata(null, ExcelSheetChanged));

        public static readonly DependencyProperty DataProperty =
            DependencyProperty.Register("Data", typeof(object), typeof(ExcelView), new PropertyMetadata(null));
        
        public static readonly DependencyProperty ErrorMessageProperty =
            DependencyProperty.Register("ErrorMessage", typeof(string), typeof(ExcelView), new PropertyMetadata(null));

        public static readonly DependencyProperty SheetNamesProperty =
            DependencyProperty.Register("SheetNames", typeof(string[]), typeof(ExcelView), new PropertyMetadata(null));

        private static void ExcelFileNameChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not ExcelView excelView)
                return;

            excelView.LoadExcel();
        }
        private static void ExcelSheetChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not ExcelView excelView)
                return;

            excelView.LoadSheet();
        }
    }
}
