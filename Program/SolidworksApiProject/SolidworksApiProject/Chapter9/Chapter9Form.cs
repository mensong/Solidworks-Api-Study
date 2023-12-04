using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;

namespace SolidworksApiProject.Chapter9
{
    public partial class Chapter9Form : Form
    {
        SldWorks swApp = null;
        ModelDoc2 DrawModleDoc = null;
        string ModleRoot = @"D:\正式版机械工业出版社出书\SOLIDWORKS API 二次开发实例详解\ModleAsbuit";
        public Chapter9Form()
        {
            InitializeComponent();

            ModleRoot = Path.GetDirectoryName(this.GetType().Assembly.Location) + @"\..\..\..\..\..\ModleAsbuit";
            ModleRoot = Path.GetFullPath(ModleRoot);
        }

        public void open_swfile(string filepath, int x, string pgid)
        {
            //SldWorks swApp;//启动程序接口//放到公共变量去
            if (x == 0)//无进程-->新建
            {
                MessageBox.Show("当前无启动中的Solidworks应用");
            }
            else if (x == 1)//有进程,得到进程的应用对象
            {
                System.Type swtype = System.Type.GetTypeFromProgID(pgid);
                swApp = (SldWorks)System.Activator.CreateInstance(swtype);
                DrawModleDoc = (ModelDoc2)swApp.ActiveDoc;
            }
        }

        //获取进程
        public void openswfile(string filepath,int x, string pgid)
        {
            if (x == 0)
            {
                MessageBox.Show("当前没有进程");
            } else if (x == 1)
            {
                System.Type swtype = System.Type.GetTypeFromProgID(pgid);
                swApp = (SldWorks)System.Activator.CreateInstance(swtype);
                DrawModleDoc = (ModelDoc2)swApp.ActiveDoc;
            }
        }

        public int getProcesson(string processName)//找此进程是否存在,0为无进程，1为存在进程，但是在由1变为0时，可能会有延迟
        {
            int x = 0;
            System.Diagnostics.Process myproc = new System.Diagnostics.Process();
            //得到所有打开的进程
            try
            {
                foreach (Process thisproc in Process.GetProcessesByName(processName))
                //循环查找
                {
                    //thisproc.Kill();
                    x = 1;//有进程打开
                    return x;
                }
            }
            catch
            {
                //Memo1.Text += "杀死" + processName + "失败！";
            }
            return x;
        }

        #region 杀SW进程
        public static void KILLSW()//清除SW进程
        {
            if (ProcessExited())//是否存在SW进程
            {
                do
                {
                    DoKillOnce();//先杀进程
                } while (ProcessExited());//再次检查是否还有进程
                System.Windows.Forms.MessageBox.Show("Soldiworks进程清理完成!");
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("无可清理的SW进程!");
            }
        }
        private static void DoKillOnce()
        {
            Process[] processes = Process.GetProcessesByName("SLDWORKS");
            foreach (Process process in processes)
            {
                if (process.ProcessName == "SLDWORKS")
                {
                    try//防止系统延迟，在杀的时候Process才没
                    {
                        process.Kill();
                    }
                    catch
                    {
                        //说明进程已经消失
                    }
                }
            }
        }
        private static bool ProcessExited()
        {
            Process[] processes = Process.GetProcessesByName("SLDWORKS");
            foreach (Process process in processes)
            {
                if (process.ProcessName == "SLDWORKS")
                {
                    return true;
                }
            }
            return false;
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            string rootpath = ModleRoot+@"\工程图模板";
            string SheetName = "A";
            string ViewTestName1 = "工程图视图1";
            string ViewTestName2 = "工程图视图2";
            string TableTempName1 = rootpath+@"\TopRightBaseTable.sldtbt";
            string TableTempName2 = rootpath + @"\TopLeftBaseTable.sldtbt";
            string Block1 = rootpath + @"\1比1右上基点.SLDBLK";
            string Block2 = rootpath + @"\1比1左下基点.SLDBLK";
            string Bom1 = rootpath + @"\BomTopRight.sldbomtbt";
            string Bom2 = rootpath + @"\BomTopLeft.sldbomtbt";

            open_swfile("",getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawModleDoc = swApp.ActiveDoc;
            SketchManager SwSketchMrg = DrawModleDoc.SketchManager;//草图选择器
            SelectionMgr SwSelMrg = DrawModleDoc.SelectionManager;//图纸选择器
            MathUtility SwMathUtility =swApp.GetMathUtility();
            MathPoint SwMathPoint = null;
            DrawingDoc SwDrawing = null;
            if (DrawModleDoc.GetType() == 3)//说明是工程图
            {
                SwDrawing = (DrawingDoc)DrawModleDoc;
            }
            else
            {
                return;
            }
            Sheet SwSheet = SwDrawing.Sheet[SheetName];
            SwDrawing.ActivateSheet(SheetName);
           
            #region 得到视图对象
            SolidWorks.Interop.sldworks.View SwView1 = null;
            SolidWorks.Interop.sldworks.View SwView2 = null;
            Feature SwFeat1 = SwDrawing.FeatureByName(ViewTestName1);
            SwFeat1.Select2(false, 0);
            Feature SwFeat2 = SwDrawing.FeatureByName(ViewTestName2);
            SwFeat2.Select2(true, 0);
            if (SwSelMrg.GetSelectedObjectCount2(-1) > 0)
            {
                for (int i = 1; i <= SwSelMrg.GetSelectedObjectCount2(-1); i++)
                {
                    SolidWorks.Interop.sldworks.View SwView = SwSelMrg.GetSelectedObjectsDrawingView2(i, -1);
                    if (SwView != null)//说明选中的是视图
                    {
                        if (SwView.GetName2() == ViewTestName1)
                        {
                            SwView1 = SwView;
                        }
                        else if (SwView.GetName2() == ViewTestName2)
                        {
                            SwView2 = SwView;
                        }
                    }
                }
            }
            #endregion

            #region 图纸比例1:1
            SwSheet.SetScale(1, 2, false, false);
            SwSketchMrg.AddToDB = true;
            SwMathPoint = SwMathUtility.CreatePoint(new double[] { 25 / 1000.0, 25 / 1000.0, 0 });

            #region [图纸格式]中创建x=50,y=50的直线
            SwDrawing.EditTemplate();
            SwSketchMrg.CreateLine(0, 0, 0, 50 / 1000.0, 50 / 1000.0, 0);//图纸左下角为原点，线画在图纸坐标系中，比例按照图纸格式比例
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, Block2, false, 1, 0);//按照图纸格式坐标系，图块定义时候的基点将与方法中的坐标点重合，按图纸格式比例1：1
            SwDrawing.EditSheet();
            #endregion

            #region [图纸]
            SwSketchMrg.CreateLine(0, 0, 0, 50 / 1000.0, 50 / 1000.0, 0);//图纸左下角为原点，线画在图纸坐标系中，比例按照图纸比例
         
            #region 获得视图相对图纸的位置坐标,并绘制图纸原点到视图位置的直线
            double[] ViewPos1 = SwView1.Position;
            double[] ViewPos2 = SwView2.Position;

            //string kk = ViewPos1[0].ToString();
            //string kk1 = ViewPos1[1].ToString();
            //StringBuilder sb = new StringBuilder(ViewPos1[0].ToString());
            //sb.Append(ViewPos1[1].ToString());
            //MessageBox.Show(sb.ToString());
            //return;

            SwSketchMrg.CreateLine(0, 0, 0, ViewPos1[0], ViewPos1[1], 0);
            SwSketchMrg.CreateLine(0, 0, 0, ViewPos2[0], ViewPos2[1], 0);
            #endregion

            #region 图纸插入注解
            Note SwNote = SwDrawing.CreateText2("图纸注解", 25 / 1000.0, 25 / 1000.0, 0, 0.003, 0);//证明按图纸坐标系，比例按照图纸格式比例
            #endregion

            return;

            #region 图纸插入普通表格
            TableAnnotation swTable = SwDrawing.InsertTableAnnotation2(false, 25 / 1000.0, 25 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopRight, TableTempName1, 3, 0);//证明按图纸格式坐标系,方法中Anchor设置无用与表格模板设置有关，模板定义的Anchor点将于方法中的坐标点重合，比例按照图纸格式比例
            #endregion

            #region 图纸插入图块
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, Block1, false, 1, 0);//按照图纸坐标系，图块定义时候的基点将与方法中的坐标点重合，位置按，比例按照图纸比例
            #endregion
            #endregion

            #region [视图]
            #region 激活的视图上画线
            SwDrawing.ActivateView(SwView1.Name);
            SwSketchMrg.CreateLine(0, 0, 0, 50 / 1000.0, 50 / 1000.0, 0);//视图中心为原点，线画在视图坐标系中，比例按照视图比例
            SwNote = SwDrawing.CreateText2("视图注解1XX", 25 / 1000.0, 25 / 1000.0, 0, 0.003, 0);//插入注解，证明按图纸格式坐标系，比例按照图纸格式比例
            swTable = SwDrawing.InsertTableAnnotation2(false, 25 / 1000.0, 25 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft, TableTempName1, 3, 0);//证明按照图纸格式坐标系,方法中Anchor设置无用与表格模板设置有关，模板定义的Anchor点将于方法中的坐标点重合，比例按照图纸格式比例
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, Block1, false, 1, 0);//按照视图坐标系，图块定义时候的基点将与方法中的坐标点重合，比例按照视图比例
            SwView1.InsertBomTable4(false, 25 / 1000.0, 25 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft, (int)swBomType_e.swBomType_TopLevelOnly, "默认 ", Bom1, false, (int)swNumberingType_e.swNumberingType_Flat, false);//图纸格式坐标系，Anchor与模板制作定义的无关，在此方法上定义的点与给定的坐标重合，比例按照图纸格式比例

            SwDrawing.ActivateView(SwView2.Name);
            SwSketchMrg.CreateLine(0, 0, 0, 50 / 1000.0, 50 / 1000.0, 0);//视图中心为原点，线画在视图坐标系中，比例按照视图比例
            SwNote = SwDrawing.CreateText2("视图注解2YYYYYY", 25 / 1000.0, 25 / 1000.0, 0, 0.003, 0);//插入注解，证明按图纸格式坐标系，比例按照图纸格式比例
            swTable = SwDrawing.InsertTableAnnotation2(false, 25 / 1000.0, 25 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomLeft, TableTempName2, 3, 0);//证明按图纸格式坐标系,方法中Anchor设置无用与表格模板设置有关，模板定义的Anchor点将于方法中的坐标点重合，比例按照图纸格式比例
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, Block2, false, 1, 0);//按照视图坐标系，图块定义时候的基点将与方法中的坐标点重合，比例按照视图比例
            SwView2.InsertBomTable4(false, 25 / 1000.0, 25 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopRight, (int)swBomType_e.swBomType_TopLevelOnly, "默认", Bom2, false, (int)swNumberingType_e.swNumberingType_Flat, false);//图纸格式坐标系，Anchor与模板制作定义的无关，在此方法上定义的点与给定的坐标重合，比例按照图纸格式比例
            #endregion

            #endregion
            SwSketchMrg.AddToDB = false;
            #endregion

            //三大坐标系图纸格式坐标系（坐标比例1：1），图纸坐标系（坐标比例为图纸比例），视图坐标系（坐标比例为视图比例）。
            //图纸格式坐标系与图纸坐标系原点一致，视图使用以视图中心为原点的坐标系
            //DrawingDoc插入的都是按照图纸格式坐标系
            //SketchManager插入的按照选择环境所在三大坐标系之一
            //View插入的Bom表按照图纸格式坐标系
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string rootpath = ModleRoot + @"\第9章\9.3.1";
            string TemplateName = rootpath+ @"\A1模板.DRWDOT";//工程图模板路径--包含了属性和文档设置信息
            string DwgFormatePath = rootpath+ @"\A1图纸格式.slddrt";//图纸格式模板
            string DwgFormateAddedPath  = rootpath+ @"\A3图纸格式.slddrt";//图纸格式模板
            string DrawStdPath = rootpath + @"\绘图标准.sldstd";

            #region 使用工程图模板新建工程图文件
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawModleDoc = swApp.NewDocument(TemplateName, 10, 0, 0);//新建图纸模板
            MessageBox.Show(DrawModleDoc.GetTitle() + "新建成功!");
            #endregion
            return;

            #region 设置绘图标准
            DrawModleDoc.Extension.LoadDraftingStandard(DrawStdPath);
            MessageBox.Show("文档总绘图标准修改为：绘图标准");
            #endregion

            DrawingDoc SwDraw = (DrawingDoc)DrawModleDoc;
            string[] SheetNames = SwDraw.GetSheetNames();//获得工程图文件中所有的图纸名
            #region 设置图纸格式
            SwDraw.SetupSheet5(SheetNames[0], 12, 12, 1, 10, true, DwgFormatePath, 0.841, 0.594, "默认", true);
            MessageBox.Show("图纸格式设置完毕");
            #endregion

            #region 添加图纸
            SwDraw.NewSheet3("新建图纸1", 12, 12, 1, 1, true, DwgFormateAddedPath, 0.42, 0.297, "默认");
            MessageBox.Show("新增图纸成功!");
            #endregion

            #region 重新激活第一张图
            SwDraw.ActivateSheet(SheetNames[0]);
            MessageBox.Show("图纸" + SheetNames[0] + "被再次激活!");
            #endregion
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DoInsert();
        }

        double[] LineStartPointXY =new double[2];
        public void DoInsert()
        {
            string SheetName = "A";
            string ViewName = "插头视图";
            string CompName = "PlugTopBox";//
            string ViewName1 = "接线板俯视";//

            string rootpath = ModleRoot+@"\工程图模板";
            string TableTempName1 = rootpath + @"\TopRightBaseTable.sldtbt";
            string TableTempName2 = rootpath + @"\TopLeftBaseTable.sldtbt";
            string Block1 = rootpath + @"\1比1右上基点.SLDBLK";
            string Block2 = rootpath + @"\1比1左下基点.SLDBLK";

            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            SketchManager SwSketchMrg = DrawModleDoc.SketchManager;
            DrawingDoc SwDrawing = null;
            SelectionMgr SwSelMrg = DrawModleDoc.SelectionManager;
            //SelectData sd = SwSelMrg.CreateSelectData();
            if (DrawModleDoc.GetType() == 3)//说明是工程图
            {
                SwDrawing = (DrawingDoc)DrawModleDoc;
            }
            else
            {
                return;
            }

            Sheet SwSheet = SwDrawing.Sheet[SheetName];
            SwDrawing.ActivateSheet(SheetName);

            SwSketchMrg.AddToDB = true;//开启直接写入数据库
            #region 注解
            #region 普通注解
            CreateNote(SwDrawing, SwSketchMrg, "无属性联动注解", 50 / 1000.0, 50 / 1000.0);
            //Note SwNote1 = SwDrawing.CreateText2("无属性联动注解", 50 / 1000.0, 50 / 1000.0, 0, 0.003, 0);//左上点定位
            //double[] NoteSizes1 = SwNote1.GetExtent();//得到左下和右上点
            //SwSketchMrg.CreateLine(0, 0, 0, NoteSizes1[0], NoteSizes1[4], NoteSizes1[2]);//注解左上点坐标
            #endregion

            #region 联动图纸属性
            CreateNote(SwDrawing, SwSketchMrg, "$PRPSHEET:\"注解联动\"", 100 / 1000.0, 100 / 1000.0);
            //Note SwNote2 = SwDrawing.CreateText2("$PRPSHEET:\"注解联动\"", 100 / 1000.0, 100 / 1000.0, 0, 0.003, 0);
            //double[] NoteSizes2 = SwNote2.GetExtent();//得到左下和右上点
            //SwSketchMrg.CreateLine(NoteSizes1[0], NoteSizes1[4], NoteSizes1[2], NoteSizes2[0], NoteSizes2[4], NoteSizes2[2]);//注解左上点坐标
            #endregion

            #region 联动视图属性
            SwDrawing.ActivateView(ViewName);
            CreateNote(SwDrawing, SwSketchMrg, "$PRPVIEW:\"注解联动\"", 150 / 1000.0, 150 / 1000.0);
            //Note SwNote3 = SwDrawing.CreateText2("$PRPVIEW:\"注解联动\"", 150 / 1000.0, 150 / 1000.0, 0, 0.003, 0);
            //double[] NoteSizes3 = SwNote3.GetExtent();//得到左下和右上点
            //SwDrawing.EditSheet();//要重新激活回图纸，否则直线就会画在视图上，从属视图坐标和比例
            //SwSketchMrg.CreateLine(NoteSizes2[0], NoteSizes2[4], NoteSizes2[2], NoteSizes3[0], NoteSizes3[4], NoteSizes3[2]);//注解左上点坐标
            #endregion

            #region 联动视图中指定的模型
            CreateNote(SwDrawing, SwSketchMrg, "$PRPSMODEL:\"注解联动\"" + " $COMP:\"PowerStrip-2@接线板俯视/PlugTopBox-1@PowerStrip\"", 200 / 1000.0, 200 / 1000.0);
            //Note SwNote4 = SwDrawing.CreateText2("$PRPSMODEL:\"注解联动\"" + " $COMP:\"PowerStrip-2@接线板俯视/PlugTopBox-1@PowerStrip\"", 200 / 1000.0, 200 / 1000.0, 0, 0.003, 0);
            //double[] NoteSizes4 = SwNote4.GetExtent();//得到左下和右上点
            //SwSketchMrg.CreateLine(NoteSizes3[0], NoteSizes3[4], NoteSizes3[2], NoteSizes4[0], NoteSizes4[4], NoteSizes4[2]);//注解左上点坐标
            #endregion

            #region 联动工程图文档属性
            CreateNote(SwDrawing, SwSketchMrg, "$PRP:\"注解联动\"", 250 / 1000.0, 250 / 1000.0);
            //Note SwNote5 = SwDrawing.CreateText2("$PRP:\"注解联动\"", 250 / 1000.0, 250 / 1000.0, 0, 0.003, 0);
            //double[] NoteSizes5 = SwNote5.GetExtent();//得到左下和右上点
            //SwSketchMrg.CreateLine(NoteSizes4[0], NoteSizes4[4], NoteSizes4[2], NoteSizes5[0], NoteSizes5[4], NoteSizes5[2]);//注解左上点坐标
            #endregion
            #endregion

            #region 插入图块
            SwDrawing.EditSheet();
            MathUtility SwMathUtility = swApp.GetMathUtility();
            MathPoint SwMathPoint = null;
            SwMathPoint = SwMathUtility.CreatePoint(new double[] { 300 / 1000.0, 100 / 1000.0, 0 });
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, Block1, false, 1, 0);
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, Block2, false, 1, 0);
            #endregion

            #region 插入表格
            TableAnnotation swTable1 = SwDrawing.InsertTableAnnotation2(false, 500 / 1000.0, 200 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopRight, TableTempName1, 3, 0);//证明按图纸格式坐标系,方法中Anchor设置无用与表格模板设置有关，模板定义的Anchor点将于方法中的坐标点重合，比例按照图纸格式比例
            //TableAnnotation swTable2 = SwDrawing.InsertTableAnnotation2(false, 500 / 1000.0, 200 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopRight, TableTempName2, 3, 0);//证明按图纸格式坐标系,方法中Anchor设置无用与表格模板设置有关，模板定义的Anchor点将于方法中的坐标点重合，比例按照图纸格式比例
            #endregion

            SwSketchMrg.AddToDB = false;//关闭直接写入数据库
        }
        public void CreateNote(DrawingDoc SwDrawing, SketchManager SwSketchMrg, string inputtext, double x, double y)
        {
            Note SwNote = SwDrawing.CreateText2(inputtext, x, y, 0, 0.03, 0);
            double[] NoteSizes = SwNote.GetExtent();//得到左下和右上点
            SwDrawing.EditSheet();//要重新激活回图纸，否则直线就会画在视图上，从属视图坐标和比例
            SwSketchMrg.CreateLine(LineStartPointXY[0], LineStartPointXY[1], 0, NoteSizes[0], NoteSizes[4], NoteSizes[2]);//注解左上点坐标
            LineStartPointXY[0] = NoteSizes[0];
            LineStartPointXY[1] = NoteSizes[4];
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string rootpath = ModleRoot + @"\工程图模板";
            string TemplateName = rootpath + @"\A3模板.DRWDOT";//工程图模板路径--包含了属性和文档设置信息
            string DwgFormatePath = rootpath + @"\A3图纸格式.slddrt";//图纸格式模板
            string DrawStdPath = rootpath + @"\绘图标准.sldstd";
            string ModleRootPath = ModleRoot + @"\第9章\9.3.3\RectanglePlug";
            string Bom1 = rootpath + @"\BomTopRight.sldbomtbt";
            string Bom2 = rootpath + @"\BomTopLeft.sldbomtbt";

            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawModleDoc = swApp.NewDocument(TemplateName, 10, 0, 0);//新建图纸
            DrawModleDoc.Extension.LoadDraftingStandard(DrawStdPath);
            DrawingDoc SwDraw = (DrawingDoc)DrawModleDoc;
            SketchManager SwSketchMrg = DrawModleDoc.SketchManager;
            string[] SheetNames = SwDraw.GetSheetNames();//获得工程图文件中所有的图纸名
            SwDraw.SetupSheet5(SheetNames[0], 12, 12, 1, 2, true, DwgFormatePath, 0.841, 0.594, "默认", true);
            DrawModleDoc.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swViewDisplayHideAllTypes, true);

            double[] view1pos = new double[] { 120 / 1000.0, 100 / 1000.0 };//定义SwView1插入点
            double[] view2pos = new double[] { 120 / 1000.0, 200 / 1000.0 };//定义SwView2插入点
            double[] view3pos = new double[] { 180 / 1000.0, 250 / 1000.0 };//定义SwView3插入点
          
            //创建视图
            SolidWorks.Interop.sldworks.View SwView1 = SwDraw.CreateDrawViewFromModelView3(ModleRootPath + @"\PowerStrip.SLDASM", "*上视", view1pos[0], view1pos[1], 0);
            SolidWorks.Interop.sldworks.View SwView2 = SwDraw.CreateDrawViewFromModelView3(ModleRootPath + @"\PowerStrip.SLDASM", "*前视", view2pos[0], view2pos[1], 0);
            SolidWorks.Interop.sldworks.View SwView3 = SwDraw.CreateDrawViewFromModelView3(ModleRootPath + @"\PlugHead\PlugHead.SLDASM", "*下视", view3pos[0], view3pos[1], 0);

            //插入表格
            SwView1.InsertBomTable4(false, 412 / 1000.0, 70 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight, (int)swBomType_e.swBomType_TopLevelOnly, "默认", Bom1, false, (int)swNumberingType_e.swNumberingType_Flat, false);
            SwView3.InsertBomTable4(false, 412 / 1000.0, 170 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight, (int)swBomType_e.swBomType_TopLevelOnly, "默认", Bom2, false, (int)swNumberingType_e.swNumberingType_Flat, false);
        
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //使用9.3.2 实例分析中的模型PowerStrip.SLDDRW
            string SheetName = "A";
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");

            DrawingDoc SwDrawing = null;
            if (DrawModleDoc.GetType() == 3)//说明是工程图
            {
                SwDrawing = (DrawingDoc)DrawModleDoc;
            }
            else
            {
                return;
            }
            Sheet SwSheet = SwDrawing.Sheet[SheetName];

            MessageBox.Show("当前图纸属性来源于视图:" + SwSheet.CustomPropertyView);

            MessageBox.Show("当前图纸名称为:" + SwSheet.GetName());
            SwSheet.SetName("AX");
            MessageBox.Show("修改后图纸名称为:" + SwSheet.GetName());

            double SheetWidth = -1;
            double SheetHeight = -1;
            int template = SwSheet.GetSize(ref SheetWidth,ref SheetHeight);
            MessageBox.Show("图纸枚举尺寸为：" + ((swDwgPaperSizes_e)template).ToString() + "\r\n宽度：" + SheetWidth.ToString() + "\r\n高度：" + SheetHeight.ToString());

            object[] ObjViews = SwSheet.GetViews();
            foreach (object ObjView in ObjViews)
            {
                SolidWorks.Interop.sldworks.View SwView = (SolidWorks.Interop.sldworks.View)ObjView;
                if (SwView.Name.Contains("*"))
                {
                    continue;
                }
                SwDrawing.ActivateView(SwView.Name);
                MessageBox.Show(SwView.Name + "被激活,图纸锁焦状态" + SwSheet.FocusLocked.ToString());
                SwSheet.FocusLocked = true;
                MessageBox.Show("图纸锁焦状态" + SwSheet.FocusLocked.ToString());
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //使用9.3.2 实例分析中的模型PowerStrip.SLDDRW
            DoView();//实例分析：视图自身属性的获得与设置
        }
        public void DoView()
        {
            string ViewName = "插头视图";
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawingDoc SwDrawing = null;
            if (DrawModleDoc.GetType() == 3)//说明是工程图
            {
                SwDrawing = (DrawingDoc)DrawModleDoc;
            }
            else
            {
                return;
            }
            SolidWorks.Interop.sldworks.View SwView = GetView(ViewName, SwDrawing);//根据视图名称获取视图标签
            string ViewDataStrs = GetViewData(SwView);//获取视图数据
            MessageBox.Show(ViewDataStrs);
            SetViewData(SwView);//重新设置视图比例位置等
            ViewDataStrs = GetViewData(SwView);
            MessageBox.Show(ViewDataStrs);
        }

        public SolidWorks.Interop.sldworks.View GetView(string ViewName, DrawingDoc SwDrawing)//得到指定名称的视图
        {
            SelectionMgr SwSelMrg = DrawModleDoc.SelectionManager;
            SolidWorks.Interop.sldworks.View SwView = null;
            Feature SwFeat = (Feature)SwDrawing.FeatureByName(ViewName);
            if (SwFeat == null)
                return null;
            SwFeat.Select2(false, 0);
            if (SwSelMrg.GetSelectedObjectCount2(-1) > 0)
            {
                for (int i = 1; i <= SwSelMrg.GetSelectedObjectCount2(-1); i++)
                {
                    SolidWorks.Interop.sldworks.View TempSwView = SwSelMrg.GetSelectedObjectsDrawingView2(i, -1);
                    if (TempSwView != null)//说明选中的是视图
                    {
                        if (TempSwView.GetName2() == ViewName)
                        {
                            SwView = TempSwView;
                        }
                    }
                }
            }
            return SwView;
        }

        public string GetViewData(SolidWorks.Interop.sldworks.View SwView)
        {
            StringBuilder sb = new StringBuilder("");
            sb.Append("视图名称：" + SwView.GetName2().ToString() + "\r\n");
            sb.Append("视图角度=" + SwView.Angle.ToString()+"\r\n");
            sb.Append("视图锁焦状态=" + SwView.FocusLocked.ToString() + "\r\n");
            double[] pos = SwView.Position;
            sb.Append("视图位置坐标{" + pos[0].ToString() +","+pos[1].ToString()+ "}\r\n");
            sb.Append("视图使用的模型配置=" + SwView.ReferencedConfiguration + "\r\n");
            sb.Append("视图小数比例=" + SwView.ScaleDecimal.ToString() + "\r\n");
            double[] scal = SwView.ScaleRatio;
            sb.Append("视图比例模式=" + scal[0].ToString() + "：" + scal[1].ToString() + "\r\n");
            sb.Append("视图所在图纸：" + SwView.Sheet.GetName() + "\r\n");
            sb.Append("视图类型：" + ((swDrawingViewTypes_e)SwView.Type).ToString() + "\r\n");
            sb.Append("视图可见性：" + SwView.GetVisible().ToString() + "\r\n");
            return sb.ToString();
        }

        public void SetViewData(SolidWorks.Interop.sldworks.View SwView)
        {
            SwView.Position = new double[] { 250 / 1000.0, 100 / 1000.0 };
            DrawModleDoc.EditRebuild3();
            MessageBox.Show("视图位置已变更");
            SwView.Angle = Math.PI / 4.0;//转45度
            SwView.FocusLocked = true;
            SwView.ScaleRatio = new double[] { 1, 4 };
            SwView.SetName2("视图重命名");
            DrawModleDoc.EditRebuild3();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ViewModleData();//实例分析：提取视图中模型数据
        }

        public void ViewModleData()
        {
            string OutputFilePath = Application.StartupPath+@"\BomOupput.txt";
            string ViewName = "接线板俯视";
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawingDoc SwDrawing = null;
            if (DrawModleDoc.GetType() == 3)//说明是工程图
            {
                SwDrawing = (DrawingDoc)DrawModleDoc;
            }
            else
            {
                return;
            }

            SolidWorks.Interop.sldworks.View SwView = GetView(ViewName, SwDrawing);
            DrawingComponent SwDrawComp = SwView.RootDrawingComponent;//获得视图中
            List<string> CompsData = new List<string>();
            GetChilerenComp(SwDrawComp, "", CompsData);

            StreamWriter sw = File.AppendText(OutputFilePath);
            foreach(string aa in CompsData)
            {
                sw.WriteLine(aa);
            }
            sw.WriteLine("******结束******");
            sw.Flush();
            sw.Close();
            Process.Start(OutputFilePath);
            //小提示，得到ModleDoc2还能修改尺寸,相同零件统计数量，这里不展开，因为完全属于编程算法，这里仅针对SW的API应用
        }
        public void GetChilerenComp(DrawingComponent RootComp,string RootPartNo,List<string> CompsData)
        {
            int i = 1;
            string PartNo="";
            string temp="";
            string BomName = "";
            string Material = "";
            object[] ObjChildDrawComps = RootComp.GetChildren();
            foreach (object ObjChildComp in ObjChildDrawComps)
            {
                DrawingComponent ChildComp = (DrawingComponent)ObjChildComp;
                Component2 SwComp = ChildComp.Component;
                ModelDoc2 CompDoc = SwComp.GetModelDoc2();
                if (RootPartNo == "")
                {
                    PartNo = i.ToString().Trim();
                }
                else
                {
                    PartNo = RootPartNo + "-" + i.ToString().Trim();
                }
                CompDoc.Extension.CustomPropertyManager[""].Get3("名称", true, out temp, out BomName);
                CompDoc.Extension.CustomPropertyManager[""].Get3("材料", true, out temp, out Material);
                CompsData.Add("件号" + PartNo + "-->规格为:" + BomName + ",材料为:" + Material);
                GetChilerenComp(ChildComp, PartNo, CompsData);
                i = i+1;
            }       
        }

        //实例分析：图纸部件的设置
        private void button8_Click(object sender, EventArgs e)
        {
            DoDrawingComponent();
        }
        public void DoDrawingComponent()
        {
            string DwgCompName = "PlugPinHead";
            string ViewName = "接线板俯视";
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            SelectionMgr SwSelMrg = DrawModleDoc.SelectionManager;
            SelectData SwSelData = SwSelMrg.CreateSelectData();
            DrawingDoc SwDrawing = (DrawingDoc)DrawModleDoc;
            SolidWorks.Interop.sldworks.View SwView = GetView(ViewName, SwDrawing);
            DrawingComponent SwRootDrawComp = SwView.RootDrawingComponent;//获得视图中
            Component2 SwRootComp = SwRootDrawComp.Component;
            ModelDoc2 SwRootDoc = SwRootComp.GetModelDoc2();
            AssemblyDoc SwRootAssem = null;
            if (SwRootDoc.GetType() == 2)
            {
                SwRootAssem = (AssemblyDoc)SwRootDoc;
            }
            Component2 SwCompNeed = null;
            for (int i = 1; i < 21; i++)
            {
                SwCompNeed = SwRootAssem.GetComponentByName(DwgCompName + "-" + i.ToString().Trim());
                if (SwCompNeed != null)
                {
                    break;
                }
            }
            if (SwCompNeed == null)
            {
                return;
            }
            DrawingComponent SwNeedDrawComp = SwCompNeed.GetDrawingComponent(SwView);
            SwNeedDrawComp.Select(false, SwSelData);
            MessageBox.Show("接线板俯视 中 PlugPinHead 被选中。\r\nDrawingComponentName为：" + SwNeedDrawComp.Name);
            string datastr = GetDrawCompData(SwNeedDrawComp, SwView.GetName2());
            MessageBox.Show(datastr);
            SwNeedDrawComp.Layer = "Divide";
            SwNeedDrawComp.Style = (int)swLineStyles_e.swLinePHANTOM;//双点划线
            SwNeedDrawComp.Width = (int)swLineWeights_e.swLW_THICK;
            DrawModleDoc.EditRebuild3();
            datastr = GetDrawCompData(SwNeedDrawComp, SwView.GetName2());
            SwNeedDrawComp.SetLineStyle((int)swDrawingComponentLineFontOption_e.swDrawingComponentLineFontHidden, (int)swLineStyles_e.swLineCENTER);
            SwNeedDrawComp.SetLineThickness((int)swDrawingComponentLineFontOption_e.swDrawingComponentLineFontHidden, (int)swLineWeights_e.swLW_THICK2, 6);
            DrawModleDoc.EditRebuild3();
            datastr = GetDrawCompData(SwNeedDrawComp, SwView.GetName2());
            MessageBox.Show(datastr);
            SwNeedDrawComp.Visible = false;       
        }
        public string GetDrawCompData(DrawingComponent DrawComp,string viewname)//获取 DrawingComponent部件的数据
        {
            StringBuilder sb = new StringBuilder("");
            sb.Append("视图" + viewname+"中，部件"+DrawComp.Name+"的信息如下：\r\n");
            sb.Append("所在图层：" + DrawComp.Layer+"\r\n");
            sb.Append("可见性：" + DrawComp.Visible.ToString() + "\r\n");
            sb.Append("属性方式获得的线型：" + ((swLineStyles_e)DrawComp.Style).ToString() + "\r\n");
            sb.Append("属性方式获得的线宽：" + ((swLineWeights_e)DrawComp.Width).ToString() + "\r\n");
            sb.Append("方法获得隐藏边线的线型：" + ((swLineStyles_e)DrawComp.GetLineStyle((int)swDrawingComponentLineFontOption_e.swDrawingComponentLineFontHidden)).ToString() + "\r\n");
            double outthick = -1;
            string x = ((swLineWeights_e)DrawComp.GetLineThickness((int)swDrawingComponentLineFontOption_e.swDrawingComponentLineFontHidden, out outthick)).ToString();
            sb.Append("方法获得隐藏边线的线宽：枚举线宽" + x + ",线宽值:"+outthick.ToString()+"\r\n");//GetLineThickness Method 
            return sb.ToString();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            DoLayer();//实例分析：图纸部件的设置
        }

        public void DoLayer()
        {
            string rootpath = ModleRoot+@"\工程图模板";
            string TemplateName = rootpath + @"\A3模板.DRWDOT";//工程图模板路径--包含了属性和文档设置信息
            string DwgFormatePath = rootpath + @"\A3图纸格式.slddrt";//图纸格式模板
            string DrawStdPath = rootpath + @"\绘图标准.sldstd";
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawModleDoc = swApp.NewDocument(TemplateName, 10, 0, 0);
            DrawingDoc SwDrawing = (DrawingDoc)DrawModleDoc;
            DrawModleDoc.Extension.LoadDraftingStandard(DrawStdPath);
            LayerMgr SwLayerMgr = DrawModleDoc.GetLayerManager();
            string[] ExistedLayers = SwLayerMgr.GetLayerList();
            StringBuilder sb = new StringBuilder("当前存在的图层如下：\r\n");
            foreach (string bb in ExistedLayers)
            {
                sb.Append(bb + "、");
            }
            MessageBox.Show(sb.ToString());

            Layer SwLayer = SwLayerMgr.GetLayer("Draw");//Draw
            string Layerdata = GetLayerData(SwLayer);
            MessageBox.Show(Layerdata, "Draw图层信息");

            SwLayerMgr.AddLayer("NewLayer", "示例新建图层", 255, (int)swLineStyles_e.swLineCENTER, (int)swLineWeights_e.swLW_THICK4);//图层颜色设置为红色
            Layerdata = GetLayerData(SwLayerMgr.GetLayer("NewLayer"));
            MessageBox.Show(Layerdata,"图层新建成功");

            MessageBox.Show("当前激活的图层为" + SwLayerMgr.GetCurrentLayer());
            Note SwNote1 = SwDrawing.CreateText2("TestLsyer", 0.05,0.05, 0, 0.03, 0);
            SwLayerMgr.SetCurrentLayer("Dim");
            MessageBox.Show("当前激活的图层为" + SwLayerMgr.GetCurrentLayer());
            Note SwNote2 = SwDrawing.CreateText2("TestLsyer2", 0.1, 0.1, 0, 0.03, 0);

            SwLayerMgr.DeleteLayer("NewLayer");
        }
        public string GetLayerData(Layer SwLayer)
        {
            StringBuilder sb = new StringBuilder("");
            sb.Append("图层名称" + SwLayer.Name+"\r\n");
            sb.Append("图层描述" + SwLayer.Description + "\r\n");
            sb.Append("图层颜色" + SwLayer.Color.ToString() +"\r\n");
            sb.Append("图层线型" + ((swLineStyles_e)SwLayer.Style).ToString() + "\r\n");
            sb.Append("图层线粗" + ((swLineWeights_e)SwLayer.Width).ToString() + "\r\n");
            sb.Append("图层可打印性" + SwLayer.Printable.ToString() + "\r\n");
            sb.Append("图层可见性" + SwLayer.Visible.ToString() + "\r\n");
            return sb.ToString();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            DoTable();//9.11.1 表格的获得及数据写入
        }

        public void DoTable()
        {
            string rootpath = ModleRoot + @"\工程图模板";
            string TableTempName1 = rootpath + @"\TopRightBaseTable.sldtbt";
            string TemplateName = rootpath + @"\A3模板.DRWDOT";//工程图模板路径--包含了属性和文档设置信息
            string DwgFormatePath = rootpath + @"\A3图纸格式.slddrt";//图纸格式模板
            string DrawStdPath = rootpath + @"\绘图标准.sldstd";
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawModleDoc = swApp.NewDocument(TemplateName, 10, 0, 0);
            DrawingDoc SwDrawing = (DrawingDoc)DrawModleDoc;
            TableAnnotation swTable = SwDrawing.InsertTableAnnotation2(false, 500 / 1000.0, 200 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopRight, TableTempName1, 3, 0);
            MessageBox.Show("表格新建成功!行数=" + swTable.RowCount.ToString() + ",列数=" + swTable.ColumnCount.ToString());

            #region 对新建的表格做标记
            for (int i = 0; i < swTable.ColumnCount; i++)
            {
                swTable.Text[0, i] = i.ToString();
            }
            for (int i = 0; i < swTable.RowCount; i++)
            {
                swTable.Text[i, 0] = i.ToString();
            }
            #endregion

            #region 表格边框线设置
            swTable.BorderLineWeight = (int)swLineWeights_e.swLW_THIN;
            swTable.GridLineWeight = (int)swLineWeights_e.swLW_THICK;
            MessageBox.Show("表格边框线设置完毕!");
            #endregion

            #region 插入行与列
            swTable.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_Before, 3, "3Before", (int)swInsertTableColumnWidthStyle_e.swInsertColumn_DefaultWidth);
            swTable.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_After, 1, "1After", (int)swInsertTableColumnWidthStyle_e.swInsertColumn_DefaultWidth);
            swTable.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_Last, 3, "3Last", (int)swInsertTableColumnWidthStyle_e.swInsertColumn_DefaultWidth);
            //列名是看不见的，同编程中的列表名
            for (int i = 0; i < 10; i++)//插入10行
            {
                swTable.InsertRow((int)swTableItemInsertPosition_e.swTableItemInsertPosition_Last, 1);
                swTable.Text[swTable.RowCount - 1, 0] = (swTable.RowCount - 1).ToString();
            }
            #endregion

            #region 解锁列宽，设置列宽
            if (swTable.GetLockColumnWidth(2))
            {
                swTable.SetLockColumnWidth(2, false);
            }
            swTable.SetColumnWidth(2, 20/1000.0, (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
            swTable.SetLockColumnWidth(2, true);
            #endregion

            #region 得到表格标题与标题位置
            swTable.Title = "表格标题";
            swTable.TitleVisible = true;
            swTable.SetHeader((int)swTableHeaderPosition_e.swTableHeader_Bottom, 3);//将表格表头设置到下方
            #endregion

            #region 合并单元格
            swTable.MergeCells(3,2,5,4);
            MessageBox.Show("单元格合并成功");
            #endregion

            #region 撤销单元格合并--仅需要选择的单元格在合并的单元格之内即可
            swTable.UnmergeCells(4,3);
            //swTable.UnmergeCells(3, 2);
            MessageBox.Show("单元格解除成功");
            #endregion

            #region 表格分割
            TableAnnotation SwSplitTable1 = SplitTable(swTable);
            StringBuilder sb = new StringBuilder("分割前swTable表格首列数值:\r\n");
            for (int i = 0; i < swTable.RowCount; i++)
            {
                sb.Append(swTable.Text[i, 0] + ",");
            }
            sb.Append("\r\n首次分割后SwSplitTable1首列值：\r\n");
            for (int i = 0; i < SwSplitTable1.RowCount; i++)
            {
                sb.Append(SwSplitTable1.Text[i, 0] + ",");
            }
            MessageBox.Show(sb.ToString(), "表格分割完成");
            #endregion

            #region 表格合并
            swTable.Merge((int)swTableMergeLocations_e.swTableMerge_WithNext);
            MessageBox.Show("前两表合并成功");
            SwSplitTable1.Merge((int)swTableMergeLocations_e.swTableMerge_WithNext);
            MessageBox.Show("后两表合并成功");
            swTable.Merge((int)swTableMergeLocations_e.swTableMerge_All);
            SwSplitTable1 = SplitTable(swTable);
            swTable.Merge((int)swTableMergeLocations_e.swTableMerge_All);
            MessageBox.Show("全部合并");
            SwSplitTable1 = SplitTable(swTable);
            SwSplitTable1.Merge((int)swTableMergeLocations_e.swTableMerge_All);
            MessageBox.Show("SwSplitTable1全部合并效果");
            #endregion
        }

        public TableAnnotation SplitTable(TableAnnotation swTable)
        {
            TableAnnotation SwSplitTable1 = swTable.Split((int)swTableSplitLocations_e.swTableSplit_AfterRow, 3);//第4行之后分割
            swTable.Split((int)swTableSplitLocations_e.swTableSplit_BeforeRow, 2);//第3行之前分割
            SwSplitTable1.Split((int)swTableSplitLocations_e.swTableSplit_BeforeRow, 7);//第8行之前分割     
            return SwSplitTable1;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Do9_11_2();
        }
        public void Do9_11_2()
        {
            string rootpath = ModleRoot + @"\工程图模板";
            string NewTableName = "NewTable";
            string BomTemple = rootpath + @"\9.11.2表格模板.sldtbt";
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawingDoc SwDraw = (DrawingDoc)DrawModleDoc;

            #region 表格新建，重名，分割，表格特征与表格注解的关系--新建时给表特征重命名有助于后期指定名称直接获得对象
            TableAnnotation SwTable = SwDraw.InsertTableAnnotation2(false, 0.412, 0.284, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopRight, BomTemple, 6, 4);
            GeneralTableFeature SwTableFeat = SwTable.GeneralTableFeature;
            Feature SwFeat = SwTableFeat.GetFeature();
            SwFeat.Name = NewTableName;
            MessageBox.Show("表格特征中表格数为" + SwTableFeat.GetTableAnnotationCount().ToString());
            SwTable.Split((int)swTableSplitLocations_e.swTableSplit_AfterRow, 2);
            MessageBox.Show("表格分割后,表格特征中表格数为" + SwTableFeat.GetTableAnnotationCount().ToString());
            //GeneralTableFeature为1张表特征，若表格被分割则一个GeneralTableFeature中有多个TableAnnotation
            #endregion
            //////////////
            #region 获得表格
            TableAnnotation SwTableToGet = null;//清空原对象，使用名字获得表格
            Feature SwFeatGet = SwDraw.FeatureByName(NewTableName);
            if (SwFeatGet != null)
            {
                SwFeatGet.Select2(false, 0);
                GeneralTableFeature gtf = DrawModleDoc.SelectionManager.GetSelectedObject6(1, -1);
                object[] ObjTables = gtf.GetTableAnnotations();
                SwTableToGet = (TableAnnotation)ObjTables[0];//随便得到一个即可
            }
            #endregion

            #region 单元格字体
            int RowToDo = 1;
            int ColToDo = 1;
            TextFormat SwTF = GetCellData(SwTableToGet, RowToDo, ColToDo);//获得
            #region 设置
            SwTableToGet.Text[RowToDo, ColToDo] = "$PRPSMODEL:\"名称\" $COMP:\"PowerStrip-2@接线板俯视/PlugSlotA-1@PowerStrip\"";
            SwTableToGet.CellTextHorizontalJustification[RowToDo, ColToDo] = (int)swTextJustification_e.swTextJustificationCenter;
            SwTableToGet.CellTextVerticalJustification[RowToDo, ColToDo] = (int)swTextAlignmentVertical_e.swTextAlignmentMiddle;
            SwTF.Bold = true;//设置字体对象参数
            SwTF.Italic = true;//设置字体对象参数
            SwTF.Underline = true;//设置字体对象参数
            SwTableToGet.SetCellTextFormat(RowToDo, ColToDo, false, SwTF);//最终将修改完的字体对象赋予此单元格
            #endregion
            SwTF = GetCellData(SwTableToGet, RowToDo, ColToDo);//再次获得变更后的信息
            #endregion

            #region 移动行-将第二行移动到当前第四行之后，即新的第四行
            SwTableToGet.MoveRow(RowToDo, (int)swTableItemInsertPosition_e.swTableItemInsertPosition_After, RowToDo + 2);//index为0起点
            MessageBox.Show("第" + (RowToDo + 1).ToString() + "行已经移动至第" + (RowToDo + 3).ToString() + "行");//实际行数是index+1
            #endregion

            #region 导出表中数据
            SwTableToGet.SaveAsText(@"E:\SwTableOutput.txt", ",");//以逗号分割单元格,注意被隐藏的行列不会被输出
            #endregion

            #region 隐藏行
            SwTableToGet.RowHidden[RowToDo + 2] = true;//index为0起点
            MessageBox.Show("第" + (RowToDo + 3).ToString() + "行已经被隐藏");//实际行数是index+1
            #endregion      
        }
        public TextFormat GetCellData(TableAnnotation SwTableToGet, int row, int col)
        {
            StringBuilder sb = new StringBuilder("单元格(" + row.ToString() + "," + col.ToString() + ")的信息：\r\n");
            sb.Append("表达式文本Text=" + SwTableToGet.Text[row, col].ToString() + "\r\n");
            sb.Append("显示文本DisplayedText=" + SwTableToGet.DisplayedText[row, col].ToString() + "\r\n");
            sb.Append("水平对齐方式:" + (swTextJustification_e)SwTableToGet.CellTextHorizontalJustification[row, col] + "\r\n");
            sb.Append("垂直对齐方式:" + (swTextAlignmentVertical_e)SwTableToGet.CellTextVerticalJustification[row, col] + "\r\n");
            TextFormat SwTF = SwTableToGet.GetCellTextFormat(row, col);
            sb.Append("是否粗体:" + SwTF.Bold.ToString() + "\r\n");
            sb.Append("是否斜体:" + SwTF.Italic.ToString() + "\r\n");
            sb.Append("是否下划线:" + SwTF.Underline.ToString() + "\r\n");
            MessageBox.Show(sb.ToString());
            return SwTF;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            DoBom();//例子：PowerStrip.SLDDRW
        }
        public void DoBom()//9.13.1 实例分析：明细表的插入获得与设置
        {
            string rootpath = ModleRoot + @"\工程图模板";
            string Bom1 = rootpath + @"\BomTopRight.sldbomtbt";
            string NewBomName = "接线板总装明细表";
            string ViewName = "接线板俯视";
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawingDoc SwDraw = (DrawingDoc)DrawModleDoc;

            #region 插入表并重命名
            SolidWorks.Interop.sldworks.View SwView = GetView(ViewName, SwDraw);
            BomTableAnnotation SwBomTableAnn = SwView.InsertBomTable4(false, 0.412, 0.01, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight, (int)swBomType_e.swBomType_TopLevelOnly, "默认", Bom1, false, (int)swNumberingType_e.swNumberingType_None, false);
            BomFeature SwBomFeat = SwBomTableAnn.BomFeature;
            DrawModleDoc.EditRebuild3();
            MessageBox.Show("新建明细表的名字为:" + SwBomFeat.Name);
            Feature SwBF = SwBomFeat.GetFeature();
            SwBF.Name = NewBomName;
            MessageBox.Show("明细表重命名为:" + SwBomFeat.Name);
            #endregion

            SwBomTableAnn = null; //清空此对象尝试重新获得
            SwBomFeat = null;
            SwBF = null;
            MessageBox.Show("所有关于表格的对象清空完成，下面进入[按指定名称]获得“接线板总装明细表”过程");

            #region 获得表格
            SwBF = SwDraw.FeatureByName(NewBomName);//根据名称获取表格
            SwBomFeat = SwBF.GetSpecificFeature2();
            GetBomFeatureData(SwBomFeat);
            object[] ObjBomTableAnns = SwBomFeat.GetTableAnnotations();
            SwBomTableAnn = (BomTableAnnotation)ObjBomTableAnns[0];
            string CustomerPro = SwBomTableAnn.GetColumnCustomProperty(2);
            MessageBox.Show("表第三列关联的模型属性为" + CustomerPro, "成功获得" + SwBomFeat.Name);
            #endregion

            SwBomTableAnn = null; //清空此对象尝试重新获得
            #region 遍历方式获得需要的插座Bom表,11章将具体讲解feature
            string ModleNameNeedToFoundBom = "PlugHead.SLDASM";
            Feature FeatForScan = DrawModleDoc.FirstFeature();
            BomFeature PlugHeadBom = null;
            while (FeatForScan != null)//遍历特征树上所有特征
            {
                if (FeatForScan.GetTypeName2() == "BomFeat")//说明是BomFeature
                {
                    BomFeature SwBomFeatForScan = (BomFeature)FeatForScan.GetSpecificFeature2();
                    if (SwBomFeatForScan != null)//说明是明细表中的特征。
                    {
                        if (SwBomFeatForScan.GetReferencedModelName().Contains(ModleNameNeedToFoundBom))//找到了模型路径一致
                        {
                            PlugHeadBom = SwBomFeatForScan;
                            break;
                        }
                    }
                }
                FeatForScan = FeatForScan.GetNextFeature();
            }
            if (PlugHeadBom != null)//说明找到了需要的BOM表
            {
                GetBomFeatureData(PlugHeadBom);
                ObjBomTableAnns = PlugHeadBom.GetTableAnnotations();
                SwBomTableAnn = (BomTableAnnotation)ObjBomTableAnns[0];
                CustomerPro = SwBomTableAnn.GetColumnCustomProperty(2);
                MessageBox.Show("表第三列关联的模型属性为" + CustomerPro, "成功获得" + PlugHeadBom.Name);
            }
            #endregion
        }
        public void GetBomFeatureData(BomFeature SwBomFeat)
        {
            StringBuilder sb = new StringBuilder("BomFeature特征信息:\r\n");
            sb.Append("是否按装配特征排序" + SwBomFeat.FollowAssemblyOrder2.ToString() + "\r\n");
            sb.Append("类型" + ((swBomType_e)SwBomFeat.TableType).ToString() + "\r\n");
            sb.Append("参考模型名字:" + SwBomFeat.GetReferencedModelName() + "\r\n");
            object x = new object();
            string[] aa = SwBomFeat.GetConfigurations(true, ref x);//当前使用的
            string[] bb = SwBomFeat.GetConfigurations(false, ref x);//所有可用的
            sb.Append("BOM当前正在使用的配置:");
            for (int i = 0; i < aa.Length; i++)
            {
                sb.Append(aa[i]);
            }
            sb.Append("\r\n");

            sb.Append("BOM当前可使用的配置:");
            for (int i = 0; i < bb.Length; i++)
            {
                sb.Append(bb[i]);
            }
            sb.Append("\r\n");
            MessageBox.Show(sb.ToString(), SwBomFeat.Name+"特征信息");     
        }

        //9.13.2 实例分析：明细栏内容的获取
        private void button13_Click(object sender, EventArgs e)
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawingDoc SwDwg = (DrawingDoc)DrawModleDoc;
            Feature SwBF = SwDwg.FeatureByName("材料明细表1");

            BomFeature SwBomFeat = SwBF.GetSpecificFeature2();//操作明细表特征属性

            #region Bom表关联数据获得
            StringBuilder sb = new StringBuilder("Bom表行数据获取信息：\r\n");
            BomTableAnnotation SwBomTableAnn = SwBomFeat.GetTableAnnotations()[0];//操作明细表内容中的特性
            string ItemNumber = "";
            string PartNumber = "";
            object[] aa = SwBomTableAnn.GetModelPathNames(8, out ItemNumber, out PartNumber);//零件-底盒
            sb.Append("第九行Bom数据：ItemNumber=" + ItemNumber + ",PartNumber=" + PartNumber+",模型路径如下：\r\n");
            foreach (string s in aa)
            {
                sb.Append(s + "\r\n");
            }
            sb.Append("\r\n");
            object[] bb = SwBomTableAnn.GetModelPathNames(4, out ItemNumber, out PartNumber);//组件-插头组件
            sb.Append("第五行Bom数据：ItemNumber=" + ItemNumber + ",PartNumber=" + PartNumber + ",模型路径如下：\r\n");
            foreach (string s in bb)
            {
                sb.Append(s + "\r\n");
            }
            MessageBox.Show(sb.ToString());
            #endregion

            #region 通过Bom表控修改零件尺寸
            object[] ObjRowComp = SwBomTableAnn.GetComponents2(7, "默认");//得到顶盒
            Component2 SwComp = (Component2)ObjRowComp[0];
            MessageBox.Show("获得模型部件Component2，名为:" + SwComp.Name2);
            ModelDoc2 topcoverDoc = SwComp.GetModelDoc2();
            topcoverDoc.Parameter("L@SketchRec").SystemValue=0.3;
            topcoverDoc.EditRebuild3();
            DrawModleDoc.EditRebuild3();
            MessageBox.Show( SwComp.Name2+"的模型尺寸已经修改!");
            #endregion

            #region 纯表模式设置格式及获得BOM明细
            TableAnnotation SwTableAnn = (TableAnnotation)SwBomTableAnn;
            SwTableAnn.SetRowHeight(5, 20 / 1000.0, (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
            MessageBox.Show(SwTableAnn.Text[5, 2]);
            #endregion

        }

        private void button14_Click(object sender, EventArgs e)
        {
            GetAnnotations();
        }

        public void GetAnnotations()
        { 
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawingDoc SwDwg = (DrawingDoc)DrawModleDoc;
            Annotation swAnn;

            #region 从视图获得Annotation
            SolidWorks.Interop.sldworks.View SwView = GetView("接线板俯视", SwDwg);
            Note swNoteGeted;
            object[] swAnns = SwView.GetAnnotations();
            foreach (object c in swAnns)
            {
                swAnn = (Annotation)c;
                if (swAnn.GetType() == (int)swAnnotationType_e.swNote)
                {
                    swNoteGeted = swAnn.GetSpecificAnnotation();
                    MessageBox.Show(swNoteGeted.GetName() + ":" + swNoteGeted.GetText());
                    swNoteGeted.SetName("俯视图描述");
                    swNoteGeted.SetText("接线板俯视图");
                    MessageBox.Show("修改后：注解名称："+swNoteGeted.GetName() + "\r\n注解值为：" + swNoteGeted.GetText());
                }
                else if (swAnn.GetType() == (int)swAnnotationType_e.swDisplayDimension)
                {
                    DisplayDimension SwDisplayDim = swAnn.GetSpecificAnnotation();
                    Dimension SwDim = SwDisplayDim.GetDimension2(0);
                    MessageBox.Show("原始尺寸名："+SwDim.FullName);
                    SwDim.Name =SwDim.Name+ "XX";
                    MessageBox.Show("修改后尺寸名：" + SwDim.FullName + "\r\n尺寸值为:" + DrawModleDoc.Parameter(SwDim.FullName).SystemValue.ToString());
                }
                else if (swAnn.GetType() == (int)swAnnotationType_e.swTableAnnotation)
                {
                    MessageBox.Show("获得表格:" + swAnn.GetName());
                }
                else if (swAnn.GetType() == (int)swAnnotationType_e.swCenterLine)
                {
                    MessageBox.Show("获得中心线:" + swAnn.GetName());
                }
            }
            #endregion  
        }
    }
}
