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


namespace SolidworksApiProject.Chapter10
{
    public partial class Chapter10Form : Form
    {
        SldWorks swApp = null;
        ModelDoc2 SwModleDoc = null;
        string ModleRoot = @"D:\正式版机械工业出版社出书\SOLIDWORKS API 二次开发实例详解\ModleAsbuit";
        public Chapter10Form()
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
                SwModleDoc = (ModelDoc2)swApp.ActiveDoc;
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
            DoPart();//10.3 实例分析：零件的自动绘制
        }

        public void DoPart()
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");

            SketchManager skm = SwModleDoc.SketchManager;
            SelectionMgr SwSelMrg = SwModleDoc.SelectionManager;
            SelectData sd = SwSelMrg.CreateSelectData();
            Feature TopPlane = ((PartDoc)SwModleDoc).FeatureByName("上视基准面");
            TopPlane.Select2(false, 0);
            Sketch SwSketch = null;
            try
            {
                swApp.SetUserPreferenceToggle(10, false);//关闭弹出尺寸修改对话框
                #region 建立凸台
                skm.InsertSketch(true);
                SketchSegment Line1 = skm.CreateLine(0, 0, 0, 0.2, 0, 0);
                SketchSegment Arc1 = skm.CreateTangentArc(0.2, 0, 0, 0.25, 0.05, 0, (int)swTangentArcTypes_e.swForward);
                SketchSegment Line2 = skm.CreateLine(0.25, 0.05, 0, 0, 0.05, 0);
                SketchSegment Line3 = skm.CreateLine(0, 0.05, 0, 0, 0, 0);
                Line1.Select4(false, sd);
                //因为有对象，方便直接选择，但是像图纸中的轮廓线，则需要使用坐标选中，非常复杂
                DisplayDimension DisplayDim1 = SwModleDoc.AddDimension2(0.1, -0.02, 0);
                Dimension Dim1 = DisplayDim1.GetDimension2(0);
                Dim1.Name = "L1";
                Dim1.SetValue3(300, (int)swSetValueInConfiguration_e.swSetValue_InAllConfigurations, "");
                Arc1.Select4(false, sd);
                DisplayDimension DisplayDim2 = SwModleDoc.AddDimension2(0.25, 0.02, 0);
                Dimension Dim2 = DisplayDim2.GetDimension2(0);
                Dim2.Name = "R1";
                Dim2.SetValue3(100, (int)swSetValueInConfiguration_e.swSetValue_InAllConfigurations, "");
                //可以看到，自动添加了几何关系
                SwSketch = skm.ActiveSketch;//获得当前激活的此草图对象，编辑草图状态能够获得
                Feature SwFeature = (Feature)SwSketch;
                SwFeature.Name = "凸台草图";//重命名，有利于后期使用时捕捉
                skm.InsertSketch(true);//同样退出草图
                SwModleDoc.ClearSelection2(true);
                SwSketch = skm.ActiveSketch;//获得当前激活的此草图对象，退出编辑草图状态，不能获得
                if (SwSketch == null)
                {
                    MessageBox.Show("当前无激活的草图!");
                }

                SwFeature = ((PartDoc)SwModleDoc).FeatureByName("凸台草图");//重新获得特征
                SwFeature.Select2(false, 0);
                SwFeature = SwModleDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.03, 0, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, 0, 0, false);
                SwFeature.Name = "新建凸台";
                #endregion

                #region 挖槽
                SwModleDoc.ClearSelection2(true);
                SwModleDoc.Extension.SelectByID2("", "FACE", 0.01, 0.03, -0.01, false, 0, null, 0);//通过坐标点选择面
                //讲解捕捉需要坐标，很困难，工程图不建议从无到有生成
                skm.InsertSketch(true);
                object[] Rectangle1 = skm.CreateCornerRectangle(0, 0, 0, 0.1, 0.01, 0);//因为由4个SketchSegment组成
                //MessageBox.Show("矩形草图实体已经添加");
                SwSketch = skm.ActiveSketch;//获得当前激活的此草图对象，编辑草图状态能够获得
                SwFeature = (Feature)SwSketch;
                SwFeature.Name = "切割槽草图";//重命名，有利于后期使用时捕捉

                SwModleDoc.ClearSelection2(true);
                SwModleDoc.Extension.SelectByID2("", "SKETCHSEGMENT", 0.09, 0.03, 0, false, 0, null, 0);//参照的不是草图坐标，还是零件坐标
                DisplayDimension DisplayDim3 = SwModleDoc.AddDimension2(0.05, 0.03, 0.02);//参照的不是草图坐标，还是零件坐标
                Dimension Dim3 = DisplayDim3.GetDimension2(0);
                Dim3.Name = "CL1";
                Dim3.SetValue3(110, (int)swSetValueInConfiguration_e.swSetValue_InAllConfigurations, "");

                SwModleDoc.ClearSelection2(true);
                SwModleDoc.Extension.SelectByID2("", "SKETCHSEGMENT", 0.11, 0.03, -0.005, false, 0, null, 0);//一定要是唯一通过此点的
                DisplayDimension DisplayDim4 = SwModleDoc.AddDimension2(0.13, 0.03, -0.002);
                Dimension Dim4 = DisplayDim4.GetDimension2(0);
                Dim4.Name = "CH1";
                Dim4.SetValue3(20, (int)swSetValueInConfiguration_e.swSetValue_InAllConfigurations, "");

                skm.InsertSketch(true);//同样退出草图

                SwFeature = ((PartDoc)SwModleDoc).FeatureByName("切割槽草图");//重新获得特征
                SwFeature.Select2(false, 0);
                SwFeature = SwModleDoc.FeatureManager.FeatureCut3(true, false, false, 0, 0, 0.01, 0.03, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, false, true, true, true, true, false, 0, 0, false);
                SwFeature.Name = "槽口";
                #endregion

                #region 编辑已存在的草图
                SwFeature = ((PartDoc)SwModleDoc).FeatureByName("凸台草图");//重新获得特征
                SwFeature.Select2(false, 0);
                SwModleDoc.EditSketch();
                SwSketch = skm.ActiveSketch;
                object[] ObjSketsegs = SwSketch.GetSketchSegments();
                SwModleDoc.ClearSelection2(true);
                foreach (object aa in ObjSketsegs)//循环选中每个草图片段实体
                {
                    SketchSegment zz = (SketchSegment)aa;
                    zz.Select4(true, sd);
                    MessageBox.Show("选中");//
                }
                skm.InsertSketch(true);//同样退出草图
                #endregion

                #region 选择边线
                SwModleDoc.ClearSelection2(true);
                SwModleDoc.Extension.SelectByID2("", "EDGE", 0.15, 0.03, 0, false, 0, null, 0);
                #endregion
            }
            catch
            {
                //报出异常进入这里
            }
            finally
            {
                //不管是否报异常，finally块内总是能运行到这里
                swApp.SetUserPreferenceToggle(10, true);//打开恢复正常
                skm.DisplayWhenAdded = true;//实时显示绘制内容
                skm.AddToDB = false;//非常复杂的草图需要开启
                skm.AutoSolve = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DoSketchBlock();//10.5 实例分析：块的插入与块数据的获取
            //在工程图中插入标题栏信息，填写相对应的属性
        }

        public void DoSketchBlock()
        {
            string rootpath = ModleRoot+@"\工程图模板";
            string Block1 = rootpath + @"\BLOCK_TITLE_ASSEM.SLDBLK";
            string Block23 = rootpath + @"\BLOCK_TITLE_PART.SLDBLK";

            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            DrawingDoc SwDrawing = (DrawingDoc)SwModleDoc;
            SketchManager smg = SwModleDoc.SketchManager;//获取草图管理对象
            MathUtility SwMathUtility = swApp.GetMathUtility();
            SelectionMgr SwSelMrg = SwModleDoc.SelectionManager;
            SelectData SwSelData = SwSelMrg.CreateSelectData();

            double[] InsertPoint = new double[] { 0.412, 0.01, 0 };
            MathPoint SwMathPoint = SwMathUtility.CreatePoint(InsertPoint);//创建点

            string[] SheetNames = SwDrawing.GetSheetNames();//获取图纸名称数组
            SketchBlockDefinition SwBlockDefinition1 = null;
            SketchBlockDefinition SwBlockDefinition2 = null;
            SketchBlockInstance SwSketchBlockInstance3 = null;
            foreach (string SheetName in SheetNames)
            {
                SwDrawing.ActivateSheet(SheetName);//激活不同的图纸
                if (SheetName == "1")
                {
                    SwBlockDefinition1 = smg.MakeSketchBlockFromFile(SwMathPoint, Block1, false, 1, 0);
                }
                else if (SheetName == "2")
                {
                    SwBlockDefinition2 = smg.MakeSketchBlockFromFile(SwMathPoint, Block23, false, 1, 0);
                }
                else if (SheetName == "3")
                {
                    SwSketchBlockInstance3 = smg.InsertSketchBlockInstance(SwBlockDefinition2, SwMathPoint, 1, 0);
                }
            }

            object[] Objs = (object[])smg.GetSketchBlockDefinitions();//获取草图块的信息
            foreach (object ObjSbd in Objs)
            {
                SketchBlockDefinition SwSbd = (SketchBlockDefinition)ObjSbd;//转换
                string BlockFileName = SwSbd.FileName;//文件名称
                //string aa = BlockFileName.Substring(BlockFileName.LastIndexOf(@"\") + 1, BlockFileName.Length - BlockFileName.LastIndexOf(@"\") - 1);
                object[] ObjNotesInBlock = SwSbd.GetNotes();
                object[] ObjSketchBlockInstances = SwSbd.GetInstances();
                StringBuilder Sb = new StringBuilder("块文件为:" + BlockFileName + "\r\n");
                Sb.Append("注释数量:" + ObjNotesInBlock.Length.ToString() + "\r\n");
                Sb.Append("实例数:" + ObjSketchBlockInstances.Length.ToString() + "\r\n");
                Sb.Append("******注释文本非空内容*******\r\n");
                foreach (object ObjNote in ObjNotesInBlock)
                {
                    Note SwNote = (Note)ObjNote;
                    if (SwNote.GetText().Trim() != "")
                    {
                        Sb.Append(SwNote.GetName() + "的值为：" + SwNote.GetText() + "\r\n");
                    }
                }
                Sb.Append("******块实例*******\r\n");
                foreach (object ObjInstance in ObjSketchBlockInstances)
                {
                    SketchBlockInstance SwSketchBlockInstance = (SketchBlockInstance)ObjInstance;
                    Sb.Append(SwSketchBlockInstance.Name + "\r\n");
                    object[] ObjAtts = SwSketchBlockInstance.GetAttributes();
                    foreach (object ObjAtt in ObjAtts)
                    {
                        Note SwNote = (Note)ObjAtt;
                        Sb.Append("[Attribute]" + SwNote.TagName + "的值为：" + SwSketchBlockInstance.GetAttributeValue(SwNote.TagName) + "\r\n");
                    }
                }
                MessageBox.Show(Sb.ToString());
            }

            Feature SwFeature = SwBlockDefinition1.GetFeature();
            SwFeature.Name = SwFeature.Name + "特征";
            SwModleDoc.ClearSelection2(true);
            MessageBox.Show("块定义重命名成功!");

            SwSketchBlockInstance3.Select(false, SwSelData);
            double[] InsertPoint5 = new double[] { 0.2, 0.01, 0 };
            MathPoint SwMathPoint5 = SwMathUtility.CreatePoint(InsertPoint5);
            SwSketchBlockInstance3.InstancePosition = SwMathPoint5;//移动
        }
    }
}
