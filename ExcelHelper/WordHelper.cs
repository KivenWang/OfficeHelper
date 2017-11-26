using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace Nanolighting.BIRCH.BIRCHCommonClsLib
{
    /// <summary>
    /// 需要导出的单条信息
    /// </summary>
    public class WordExportInfo
    {
        /// <summary>
        /// 需要导出的文本
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// 需要导出的图片路径
        /// </summary>
        public string ImagePath { get; set; }
    }

    /// <summary>
    /// 支持从Word模板导出Word到指定路径.
    /// </summary>
    public class WordHelper
    {
        /// <summary>
        /// 通过Word模板标签生成word
        /// </summary>
        /// <param name="templateFileName">模板文件路径</param>
        /// <param name="exportFileName">新文件路径</param>
        /// <param name="info">需要导出的信息键值对,key为模板标签名,value为需要导出的信息</param>
        /// <param name="startPageIndex">word页数</param>
        public static bool ExportWord(string templateFileName, string exportFileName, Dictionary<string, WordExportInfo> info)
        {
            //生成documnet对象
            Word._Document doc = new Microsoft.Office.Interop.Word.Document();
            //生成word程序对象
            Word.Application app = new Word.Application();
            //模板文件
            //模板文件拷贝到新文件
            File.Copy(templateFileName, exportFileName);

            object Obj_FileName = exportFileName;
            object Visible = false;
            object ReadOnly = false;
            object missing = System.Reflection.Missing.Value;

            try
            {
                //打开文件
                doc = app.Documents.Open(ref Obj_FileName, ref missing, ref ReadOnly, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref Visible,
                    ref missing, ref missing, ref missing,
                    ref missing);
                doc.Activate();

                #region 将文本或图片插入到模板中对应标签
                if (info.Count > 0)
                {
                    object what = Word.WdGoToItem.wdGoToBookmark;
                    object WordMarkName;
                    foreach (var item in info)
                    {
                        WordMarkName = item.Key;
                        //光标转到书签的位置
                        doc.ActiveWindow.Selection.GoTo(ref what, ref missing, ref missing, ref WordMarkName);
                        //插入的内容，插入位置是word模板中书签定位的位置
                        if (item.Value.ImagePath == null)
                        {
                            doc.ActiveWindow.Selection.TypeText(item.Value.Text);
                        }
                        else
                        {
                            //注意此处需要对应模板文件的图片处的书签
                            object oStart = doc.Bookmarks.get_Item("Image");
                            Object linkToFile = false;       //图片是否为外部链接   
                            Object saveWithDocument = true;  //图片是否随文档一起保存    
                            object range = doc.Bookmarks.get_Item(ref oStart).Range;//图片插入位置       
                            FileInfo filePhotoInfo = new FileInfo(item.Value.ToString());
                            if (filePhotoInfo.Exists == false)
                                break;
                            doc.InlineShapes.AddPicture(item.Value.ImagePath, ref linkToFile, ref saveWithDocument, ref range);
                            doc.Application.ActiveDocument.InlineShapes[1].Width = 60;   //设置图片宽度               
                            doc.Application.ActiveDocument.InlineShapes[1].Height = 70;  //设置图片高度  
                        }
                        //设置当前定位书签位置插入内容的格式,建议直接在模板中设置
                        //doc.ActiveWindow.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                }
                #endregion

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                //输出完毕后关闭doc对象
                object IsSave = true;
                doc.Close(ref IsSave, ref missing, ref missing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                doc = null;
            }
        }
    }
}
