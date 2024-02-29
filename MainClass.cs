
//2020年7月28日更新：将数据直接输出到Excel文件中
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


namespace Dll类库_输出多段线坐标
{
    public class MainClass
    {
        public static string DirPath = "D:";
        public static string OutTxtFileName = "out.txt";
        public static string OutExcelFileName = "out.xls";

        public static string NotePadPath = @"C:\Windows\System32\notepad.exe";
        [CommandMethod("PLCGQ")]//PLineVertexCoordsGet
        public void PLCGQ()
        {
            string filePath = @"D:\LICENSE.txt";
            #region 验证
            if (File.Exists(filePath))
            {

            }
            else
            {

            }
                #endregion



                //需要访问Database的操作 需首先将该文档进行锁定，操作完成后，在最后进行释放
                DocumentLock docLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument();
            // 对话框窗口
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            // 数据库对象
            Database db = HostApplicationServices.WorkingDatabase;

            UsersInputEntity entities = GetPolyline();

            Polyline pLine = entities.Polyline;
            Point3d BasePoint = entities.BasePonint;
            if ((pLine == null) || (BasePoint == null) || (entities.isSelected == false))
            {
                //ed.WriteMessage("\n未正确选择多段线或基点！");
                //ed.WriteMessage("\n未正确选择多段线或基点！");
                return;
            }

            bool isclosed = pLine.Closed;

            bool isLeftToRight = true;//从左到右方向

            //bool isConvex = true;//上凸

            if (!isclosed)
            {
                if (pLine.StartPoint.X >= pLine.EndPoint.X)
                {
                    isLeftToRight = false;
                }
                else
                {
                    isLeftToRight = true;
                }
            }
            else
            {
                ed.WriteMessage("\n多段线闭合");
                return;
            }


            int vertexNum = pLine.NumberOfVertices;//顶点vertex 数 

            List<Point3d> vertex_List = new List<Point3d>();//多段线顶点
            List<Point3d> midPoint_List = new List<Point3d>();//多段线各段中点
            List<double> vBulgeList = new List<double>();//多段线凸度  前进方向右边为正   不凸为0

            for (int i = 0; i < vertexNum; i++)
            {
                vertex_List.Add(pLine.GetPoint3dAt(i));
                vBulgeList.Add(pLine.GetBulgeAt(i));
            }
            for (int i = 0; i < vertexNum - 1; i++)
            {
                midPoint_List.Add(pLine.GetPointAtParameter(i + 0.5));
            }

            #region 判断是否从左至右边，若不是则反转
            List<OutData> outDataList = new List<OutData>();////输出数据

            if (!isLeftToRight)
            {
                //midPoint_List.Reverse();
                //vertex_List.Reverse();
                //vBulgeList.Reverse();
                //for (int i = 0; i < vBulgeList.Count; i++) //凸度为负数
                //{
                //    vBulgeList[i] = -1.0 * vBulgeList[i];
                //}
                //for (int j = 0; j < vBulgeList.Count; j++)
                //{
                //    if (vBulgeList[j] != 0)
                //    {

                //        double R = Getradius(vertex_List[j - 1], midPoint_List[j], vertex_List[j]);
                //        Point3d point = GetCossPoint(vertex_List[j-1], midPoint_List[j], vertex_List[j]);

                //        outDataList.Add(new OutData(vertex_List[j].X, vertex_List[j].Y, 0.0));
                //        outDataList.Add(new OutData(point.X, point.Y, R));
                //    }
                //    else
                //    {
                //        outDataList.Add(new OutData(vertex_List[j].X, vertex_List[j].Y, 0.0));
                //    }
                //}
                for (int j = 0; j < vBulgeList.Count; j++)
                {
                    if (vBulgeList[j] != 0)
                    {
                        double R = Getradius(vertex_List[j + 1], midPoint_List[j], vertex_List[j]);
                        Point3d point = GetCossPoint(vertex_List[j + 1], midPoint_List[j], vertex_List[j]);
                        outDataList.Add(new OutData(vertex_List[j].X, vertex_List[j].Y, 0.0));
                        outDataList.Add(new OutData(point.X, point.Y, R));
                    }
                    else
                    {
                        outDataList.Add(new OutData(vertex_List[j].X, vertex_List[j].Y, 0.0));
                    }
                }
                outDataList.Reverse();//反转坐标
            }
            else
            {
                for (int j = 0; j < vBulgeList.Count; j++)
                {
                    if (vBulgeList[j] != 0)
                    {
                        double R = Getradius(vertex_List[j], midPoint_List[j], vertex_List[j + 1]);
                        Point3d point = GetCossPoint(vertex_List[j], midPoint_List[j], vertex_List[j + 1]);
                        outDataList.Add(new OutData(vertex_List[j].X, vertex_List[j].Y, 0.0));
                        outDataList.Add(new OutData(point.X, point.Y, R));
                    }
                    else
                    {
                        outDataList.Add(new OutData(vertex_List[j].X, vertex_List[j].Y, 0.0));
                    }
                }
            }
            #endregion

            ///在命令行打印结果
            foreach (var item in outDataList)
            {
                //ed.WriteMessage("\nX:{0}   Y:{1}   R:{2}", item.X- BasePoint.X, item.Y - BasePoint.Y, item.R);
                ed.WriteMessage("\nX:{0}   Y:{1}   R:{2}", item.X, item.Y, item.R);
            }
            ed.WriteMessage("\nIsLeftToRight:{0}", isLeftToRight.ToString());
            ed.WriteMessage("\nDeveloped by CGQ");            


            //换算基点坐标
            for (int i = 0; i < outDataList.Count; i++)
            {
                outDataList[i].X = outDataList[i].X - BasePoint.X;
                outDataList[i].Y = outDataList[i].Y - BasePoint.Y;
            }

            ///输出
            docLock.Dispose();//解锁文档
            //string FullPath = DirPath + "\\" + OutTxtFileName;

            string FullPath = DirPath + "\\" + OutExcelFileName;
            OutPutData(outDataList, FullPath);

            if (File.Exists(NotePadPath))
            {
                System.Diagnostics.Process.Start("notepad.exe", FullPath);
            }

            //OutPutDataToExcel(outDataList, FullPath);   //2021年4月3日修改  由于输出excel不稳定，改回输出txt

        }

        /// <summary>
        /// 得到用户选择的多段线和基点
        /// </summary>
        /// <returns></returns>
        private static UsersInputEntity GetPolyline()
        {
            Polyline pLine = new Polyline();
            Point3d basePoint = new Point3d();

            //需要访问Database的操作 需首先将该文档进行锁定，操作完成后，在最后进行释放
            DocumentLock docLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument();
            // 对话框窗口
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            // 数据库对象
            Database db = HostApplicationServices.WorkingDatabase;
            bool getPolyline = false;
            bool getbasePoint = false;


            try
            {
                // 开启事务处理
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    // 选择多段线
                    while (!getPolyline)
                    {
                        PromptEntityOptions entOpts = new PromptEntityOptions("\n请选择多段线");
                        PromptEntityResult entResult = ed.GetEntity(entOpts);
                        // 判断选择是否成功
                        if (entResult.Status == PromptStatus.OK)
                        {
                            // 获取多段线实体对象
                            pLine = trans.GetObject(entResult.ObjectId, OpenMode.ForRead) as Polyline;
                            if (pLine != null)
                            {
                                getPolyline = true;
                            }
                            //多线段是否闭合 pline.Closed
                            //string isclosed = pLine.Closed.ToString();
                            //多线段起始点 pline.StartPoint
                            //多线段结束点 pline.EndPoint
                            //多段线顶点数                        
                            //int vertexNum = pLine.NumberOfVertices;
                            //Point3d point;
                            //// 遍历获取多段线顶点坐标
                            //for (int i = 0; i < vertexNum; i++)
                            //{
                            //    point = pLine.GetPoint3dAt(i);
                            //    ed.WriteMessage("\n" + point.ToString());
                            //}
                            //ed.WriteMessage("\n" + pLine.StartPoint.ToString());
                            //ed.WriteMessage("\n" + pLine.EndPoint.ToString());
                        }
                        else if (entResult.Status == PromptStatus.Cancel)
                        {
                            return new UsersInputEntity(pLine, basePoint, false);
                        }
                    }

                    while (!getbasePoint)
                    {
                        PromptPointOptions entOpts_2 = new PromptPointOptions("\n请选择原点");
                        PromptPointResult entResult_2 = ed.GetPoint(entOpts_2);
                        if (entResult_2.Status == PromptStatus.OK)
                        {
                            basePoint = entResult_2.Value;
                            getbasePoint = true;
                        }
                        else if (entResult_2.Status == PromptStatus.Cancel)
                        {
                            return new UsersInputEntity(pLine, basePoint, false);
                        }
                    }
                    trans.Commit();
                }
            }

            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(e.Message);
            }
            // 解锁文档
            docLock.Dispose();
            return new UsersInputEntity(pLine, basePoint, true);
        }

        /// <summary>
        /// 求半径R
        /// </summary>
        /// <param name="startPoint"></param>
        /// <param name="midPoint"></param>
        /// <param name="endPoint"></param>
        /// <returns></returns>
        private double Getradius(Point3d startPoint, Point3d midPointOfArc, Point3d endPoint)
        {
            double X1 = startPoint.X; double Y1 = startPoint.Y;
            double X2 = midPointOfArc.X; double Y2 = midPointOfArc.Y;
            double X3 = endPoint.X; double Y3 = endPoint.Y;
            Point3d midPoint = new Point3d((X1 + X3) / 2, (Y1 + Y3) / 2, 0);
            double b = Math.Sqrt(Math.Pow((X1 - X3), 2.0) + Math.Pow((Y1 - Y3), 2.0)) / 2.0;//distanceOf X1_ X3
            double a = Math.Sqrt(Math.Pow((midPoint.X - X2), 2.0) + Math.Pow((midPoint.Y - Y2), 2.0));//distance Of mid_X2

            double radius = (Math.Pow(a, 2.0) + Math.Pow(b, 2.0)) / (2 * a);
            return radius;
        }

        /// <summary>
        /// 求两切线交点
        /// </summary>
        /// <param name="startPoint"></param>
        /// <param name="midPoint"></param>
        /// <param name="endPoint"></param>
        /// <returns></returns>
        private Point3d GetCossPoint(Point3d startPoint, Point3d midPointOfArc, Point3d endPoint)
        {


            double R = Getradius(startPoint, midPointOfArc, endPoint);

            double X1 = startPoint.X; double Y1 = startPoint.Y;
            double X2 = midPointOfArc.X; double Y2 = midPointOfArc.Y;
            double X3 = endPoint.X; double Y3 = endPoint.Y;
            Point3d midPoint = new Point3d((X1 + X3) / 2, (Y1 + Y3) / 2, 0);

            double delta_X = midPoint.X - X2;
            double delta_Y = Y2 - midPoint.Y;
            double a = Math.Sqrt(Math.Pow((midPoint.X - X2), 2.0) + Math.Pow((midPoint.Y - Y2), 2.0));//distance Of mid_X2
            double b = Math.Sqrt(Math.Pow((X1 - X3), 2.0) + Math.Pow((Y1 - Y3), 2.0)) / 2.0;//distanceOf X1_ X3

            double sin_Sita = delta_Y / a;
            double cos_Sita = delta_X / a;


            double c = Math.Sqrt(Math.Pow(R, 2.0) - Math.Pow(b, 2.0));//distanceOf circleCenter_ midPoint
            double d = Math.Pow(R, 2.0) / c - a - c;

            double delta_X_fromX2 = d * cos_Sita;
            double delta_Y_fromY2 = d * sin_Sita;

            Point3d CrossPoint = new Point3d(X2 - delta_X_fromX2, Y2 + delta_Y_fromY2, 0);
            return CrossPoint;
        }

        private void OutPutData(List<OutData> _outDataList, string outpath)
        {
            string filePath = Path.GetDirectoryName(outpath);

            if ((Directory.Exists(filePath) == false))
            {
                return;
            }
            //保存到本地的路径
            System.IO.StreamWriter mStreamWriter = new System.IO.StreamWriter(outpath, false, System.Text.Encoding.UTF8);

            //mStreamWriter.WriteLine("序号\t\tX\t\tY\t\tR");
            mStreamWriter.WriteLine("☆★☆★ Developed by CGQ ☆★☆★\n");

            mStreamWriter.WriteLine("X\tY\tZ\tR");


            for (int i = 0; i < _outDataList.Count; i++)
            {

                //mStreamWriter.WriteLine("{0}#{1}#{2}", Arrayblocks[i, j].centroid.X.ToString(), Arrayblocks[i, j].centroid.Y.ToString(), Arrayblocks[i, j].Area.ToString());// X  Y Area

                //mStreamWriter.WriteLine("{0}\t\t{1}\t\t{2}\t\t{3}", (i+1).ToString(), Math.Round(_outDataList[i].X,2).ToString(), Math.Round(_outDataList[i].Y,2).ToString(), Math.Round(_outDataList[i].R,2).ToString());
                //2020年7月24日修改输出格式 匹配midas  X  Y  Z  R  将Y列输出成0，去掉序号
                mStreamWriter.WriteLine("{0}\t{1}\t{2}\t{3}", Math.Round(_outDataList[i].X, 5).ToString(), 0.ToString(), Math.Round(_outDataList[i].Y, 5).ToString(), Math.Round(_outDataList[i].R, 5).ToString());

            }
            //用完StreamWriter的对象后一定要及时销毁
            mStreamWriter.Close();
            mStreamWriter.Dispose();
            mStreamWriter = null;
            return;
        }

        /// <summary>
        /// 设置输出目录
        /// </summary>

        [CommandMethod("SetDir")]
        public void SetDir()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            PromptStringOptions optionsx = new PromptStringOptions("\n请输入数据输出目录[选择目录(A)]:");
            //optionsx.Keywords.Add("A");
            //optionsx.Keywords.Add("M");
            //optionsx.Keywords.Default = "M";
            optionsx.AllowSpaces = true;
            optionsx.DefaultValue = DirPath;
            //optionsx.UseBasePoint = true; //允许使用基准点
            //optionsx.BasePoint = ptPrevious;//设置基准点
            optionsx.AppendKeywordsToMessage = false;//不将关键字列表添加到提示信息中

            PromptResult resultx = doc.Editor.GetString(optionsx);
            switch (resultx.Status)
            {
                case PromptStatus.OK:
                    if ((resultx.StringResult == "A") || (resultx.StringResult == "a"))
                    {
                        FolderBrowserDialog path = new FolderBrowserDialog();
                        //path.ShowDialog();
                        if (path.ShowDialog() == DialogResult.OK)
                        {
                            string selectpath = path.SelectedPath;
                            DirPath = selectpath;
                            // 所选文件夹selectFolderDialog.SelectedPath;
                        }
                    }
                    else
                    {
                        if ((Directory.Exists(Path.GetDirectoryName(resultx.StringResult)) == true))
                        {
                            DirPath = resultx.StringResult;
                        }
                        else
                        {
                            doc.Editor.WriteMessage("\n指定目录不存在");
                        }
                    }
                    doc.Editor.WriteMessage("\n输出目录：" + DirPath);
                    break;
                case PromptStatus.None:// 空输入
                    DirPath = optionsx.DefaultValue;
                    //doc.Editor.WriteMessage("\n输出目录 = " + Common.DirPath);
                    break;
                case PromptStatus.Cancel:
                    DirPath = optionsx.DefaultValue;
                    //doc.Editor.WriteMessage("\n输出目录 = " + Common.DirPath);
                    break;
                case PromptStatus.Keyword:
                    break;
                default:
                    //ed.WriteMessage(Common.DirPath.ToString());
                    break;
            }


        }

        private void OutPutDataToExcel(List<OutData> _outDataList, string outpath)
        {        
            string filePath = Path.GetDirectoryName(outpath);
            
            if ((Directory.Exists(filePath) == false))
            {
                return;
            }

            if (File.Exists(outpath))
            {
                File.Delete(outpath);
            }
            try
            {
                Microsoft.Office.Interop.Excel.Application xlsApp;
                Excel.Worksheet worksheet;
                Excel.Workbook workbook;

                xlsApp = new Excel.Application();
                xlsApp.Visible = true;
                workbook = xlsApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

                worksheet.Name = "Data";
                worksheet.Cells[1, 1] = "☆★☆★ Developed by CGQ ☆★☆★";
                //Range rangeProgram = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, 4]);//获取需要合并的单元格的范围
                //rangeProgram.Application.DisplayAlerts = false;
                //rangeProgram.Merge(Missing.Value);
                //rangeProgram.Application.DisplayAlerts = true;
                //worksheet.Columns.EntireColumn.AutoFit();//列宽自适应
                worksheet.Range["A1:D1"].Merge(); //纵向合并 
                worksheet.Cells[2, 1] = "X";
                worksheet.Cells[2, 2] = "Y";
                worksheet.Cells[2, 3] = "Z";
                worksheet.Cells[2, 4] = "R";
                worksheet.Cells[2, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
                worksheet.Cells[2, 1].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中
                worksheet.Cells[2, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
                worksheet.Cells[2, 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中
                worksheet.Cells[2, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
                worksheet.Cells[2, 3].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中
                worksheet.Cells[2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
                worksheet.Cells[2, 4].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中

                for (int i = 0; i < _outDataList.Count; i++)
                {
                    worksheet.Cells[i + 3, 1] = Math.Round(_outDataList[i].X, 4).ToString();

                    worksheet.Cells[i + 3, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
                    worksheet.Cells[i + 3, 1].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中  

                    worksheet.Cells[i + 3, 2] = "0";
                    worksheet.Cells[i + 3, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
                    worksheet.Cells[i + 3, 2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中 

                    worksheet.Cells[i + 3, 3] = Math.Round(_outDataList[i].Y, 4).ToString();
                    worksheet.Cells[i + 3, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
                    worksheet.Cells[i + 3, 3].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中 

                    worksheet.Cells[i + 3, 4] = Math.Round(_outDataList[i].R, 4).ToString();
                    worksheet.Cells[i + 3, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
                    worksheet.Cells[i + 3, 4].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中 
                                                                                                                         //mStreamWriter.WriteLine("{0}\t{1}\t{2}\t{3}", Math.Round(_outDataList[i].X, 3).ToString(), 0.ToString(), Math.Round(_outDataList[i].Y, 3).ToString(), Math.Round(_outDataList[i].R, 3).ToString());

                }
                //workbook.SaveAs(outpath);
                //    workbook.Close();
                //    xlsApp.Quit();
            }
            finally { }
  

        }

        /// <summary>
        /// 验证注册码
        /// </summary>
        public bool Verification()
        {
            string licensetime = "";
            string time = "";//真实时间
            string filePath = @"D:\LICENSE.txt";
            string timeNow = "";
            try
            {
                if (File.Exists(filePath))
                {
                    licensetime = File.ReadAllText(filePath);
                    byte[] mybyte = Encoding.UTF8.GetBytes(licensetime);
                    licensetime = Encoding.UTF8.GetString(mybyte);
                    string timeTemp = Decode(licensetime);
                    time = timeTemp.Substring(0, timeTemp.Length - 4);
                    timeNow = DateTime.Now.ToString("yyyyMMdd") + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString();        // 08:05:57;
                    int.TryParse(timeNow, out int timeNowNum);
                    int.TryParse(time, out int timeNum);

                    if (timeNowNum > timeNum)
                    {
                        //return false;
                    }
                }
                else
                {
                    //return false;
                }

            }
            catch (System.Exception ex)
            {
            }

            //验证

            return true;
        }

        /// <summary>
        /// <函数：Decode>
        ///作用：将16进制数据编码转化为字符串，是Encode的逆过程
        /// </summary>
        /// <param name="strDecode"></param>
        /// <returns></returns>
        public static string Decode(string strDecode)
        {
            string sResult = "";
            for (int i = 0; i < strDecode.Length / 4; i++)
            {
                sResult += (char)short.Parse(strDecode.Substring(i * 4, 4), global::System.Globalization.NumberStyles.HexNumber);
            }
            return sResult;
        }

        /// <summary>
        /// <函数：Encode>
        /// 作用：将字符串内容转化为16进制数据编码，其逆过程是Decode
        /// 参数说明：
        /// strEncode 需要转化的原始字符串
        /// 转换的过程是直接把字符转换成Unicode字符,比如数字"3"-->0033,汉字"我"-->U+6211
        /// 函数decode的过程是encode的逆过程.
        /// </summary>
        /// <param name="strEncode"></param>
        /// <returns></returns>
        public static string Encode(string strEncode)
        {
            string strReturn = "";//  存储转换后的编码
            foreach (short shortx in strEncode.ToCharArray())
            {
                strReturn += shortx.ToString("X4");
            }
            return strReturn;
        }
    }
}
