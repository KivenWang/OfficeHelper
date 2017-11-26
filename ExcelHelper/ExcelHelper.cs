using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Drawing;
using System.Data;

namespace Nanolighting.BIRCH.BIRCHCommonClsLib
{
    /// <summary>
    /// 导出信息到Excel模板
    /// </summary>
    public class ExcelHelper : IDisposable
    {
        /// <summary>
        /// 标识对象资源是否释放
        /// </summary>
        private bool _disposed = false;
        /// <summary>
        /// 模板文件
        /// </summary>
        private string _templateFile { get; set; }
        /// <summary>
        /// 导出文件
        /// </summary>
        private string _outputFile { get; set; }
        /// <summary>
        /// Excel应用程序
        /// </summary>
        private Excel.Application _excelApp = new Excel.Application();
        /// <summary>
        /// 工作薄
        /// </summary>
        private Excel.Workbook _workbook;
        /// <summary>
        /// 工作表
        /// </summary>
        private Excel.Worksheet _worksheet;

        /// <summary> 
        /// 导出信息到Excel模板
        /// 请在外部进行异常处理
        /// </summary> 
        /// <param name="templetFilePath"> Excel模板完整文件路径 </param> 
        /// <param name="outputFilePath"> 输出Excel完整文件路径 </param> 
        public ExcelHelper(string templateFilePath, string outputFilePath)
        {
            if (templateFilePath == null)
                throw new Exception(" Excel template file is null! ");
            if (!File.Exists(templateFilePath))
                throw new Exception(" Excel template file is not't exist! ");
            this._templateFile = templateFilePath;
            this._outputFile = outputFilePath;
            Excel.Workbooks workbooks = _excelApp.Workbooks;
            //加载模板
            this._workbook = workbooks.Add(templateFilePath);
            Excel.Sheets sheets = this._workbook.Sheets;
            //第一个工作薄
            this._worksheet = (Excel.Worksheet)sheets.get_Item(1);
        }
        /// <summary>
        /// 向Excel中添加图片
        /// </summary>
        /// <param name="fileName">需要添加图片的详细路径</param>
        /// <param name="startCellRow">单元格起始行</param>
        /// <param name="startCellColumn">单元格起始列</param>
        /// <param name="endCellRow">单元格结束行</param>
        /// <param name="endCellColumn">单元格结束列</param>
        public void AddPicToExcel(string fileName, int startCellRow, int startCellColumn, int endCellRow, int endCellColumn)
        {
            if (!File.Exists(fileName))
                throw new Exception(string.Format("{} is not exist!", fileName));
            //获取单元格范围
            Excel.Range oRange = _worksheet.Range[_worksheet.Cells[startCellRow, startCellColumn],
                _worksheet.Cells[endCellRow, endCellColumn]];
            float left = (float)(oRange.Left);
            float top = (float)(oRange.Top);
            float width = (float)(oRange.Width);
            float height = (float)(oRange.Height);
            //添加图片
            _worksheet.Shapes.AddPicture(fileName, Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoCTrue, left, top, width, height);

        }
        /// <summary>
        /// 向Excel中添加文本
        /// </summary>
        /// <param name="text">文本</param>
        /// <param name="rowCell">单元格行</param>
        /// <param name="columnCell">单元格列</param>
        public void AddTextToExcel(string text, int rowCell, int columnCell)
        {
            _worksheet.Cells[rowCell, columnCell] = text;
            //单元格自适应
            //worksheet.Cells.EntireColumn.AutoFit();
        }

        /// <summary>
        /// 将DataTable的内容导出到Excel
        /// </summary>
        /// <param name="dataTable">原始DataTable数据</param>
        /// <param name="startRow">单元格开始行</param>
        /// <param name="startColumn">单元格开始列</param>
        /// <param name="isGridLine">是否需要绘制边框线</param>
        public void DataTableToExcel(DataTable dataTable, int startRow, int startColumn, bool isGridLine)
        {
            int i, j = 0;
            for (i = 0; i < dataTable.Rows.Count; i++)
            {
                for (j = 0; j < dataTable.Columns.Count; j++)
                {
                    _worksheet.Cells[startRow + i, startColumn + j] = dataTable.Rows[i][j];
                }
            }
            //居中对齐
            _worksheet.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            if (isGridLine)
            {
                Excel.Range datatableRange = _worksheet.Range[_worksheet.Cells[startRow, startColumn],
                    _worksheet.Cells[startRow + i - 1, startColumn + j - 1]];
                //边框线
                datatableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }
        }
        /// <summary>
        /// 保存当前Excel文件
        /// </summary>
        public void SaveExcel()
        {
            object misValue = System.Reflection.Missing.Value;
            this._workbook.SaveAs(_outputFile, misValue, misValue, misValue, misValue, misValue,
                Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //关闭工作薄
            this._workbook.Close(true, misValue, misValue);
            //退出Excel应用程序
            this._excelApp.Quit();
        }

        # region 强制结束Excel进程
        [System.Runtime.InteropServices.DllImport("User32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        /// <summary>
        /// 对Excel的进程进行处理
        /// </summary>
        /// <param name="excel"></param>
        public static void ExcelProecssKill(Excel.Application excel)
        {
            if (excel != null)
            {
                //得到这个句柄，具体作用是得到这块内存入口
                IntPtr t = new IntPtr(excel.Hwnd);
                int k = 0;
                //得到本进程唯一标志k
                GetWindowThreadProcessId(t, out k);
                //得到对进程k的引用
                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
                //关闭进程k
                p.Kill();
            }
        }
        # endregion
        #region IDisposable Members

        //实现IDisposable接口方法,MSDN推荐的非托管资源释放方法
        //由类的使用者，在外部显示调用，释放类资源
        //托管资源通常由GC自动回收,但非托管资源占用较大,建议显示调用回收
        public void Dispose()
        {
            Dispose(true);
            //将对象从垃圾回收器链表中移除，
            // 从而在垃圾回收器工作时，只释放托管资源，而不执行此对象的析构函数
            GC.SuppressFinalize(this);
        }

        //由垃圾回收器调用，释放非托管资源
        ~ExcelHelper()
        {
            Dispose(false);
        }

        //参数为true表示释放所有资源，只能由使用者调用
        //参数为false表示释放非托管资源，只能由垃圾回收器自动调用
        //如果子类有自己的非托管资源，可以重载这个函数，添加自己的非托管资源的释放
        //但是要记住，重载此函数必须保证调用基类的版本，以保证基类的资源正常释放
        protected virtual void Dispose(bool disposing)
        {
            if (!this._disposed)
            {
                if (disposing)
                {
                    // 释放托管资源                    
                }
                // 释放非托管资源
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this._worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this._workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this._excelApp);
                this._worksheet = null;
                this._workbook = null;
                this._excelApp = null;
            }
            this._disposed = true;
        }
        #endregion
    }

}