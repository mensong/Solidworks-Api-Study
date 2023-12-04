using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Diagnostics;

namespace SolidworksApiProject.Chapter15
{
    public partial class Chapter15Form : Form
    {
        SldWorks swApp = null;
        ModelDoc2 AssemModleDoc = null;
        ModelDoc2 DrawingModleDoc = null;
        string ModleRoot = @"D:\正式版机械工业出版社出书\SOLIDWORKS API 二次开发实例详解\ModleAsbuit";
        string SrcPath ="";
        string TargetPath = "";
        string rootpath = "";
        
        int DoMode = 1;//1=新建，2=修改

        #region 尺寸数据,单位mm
        double PlugOD =0;


        #endregion


        public Chapter15Form()
        {
            InitializeComponent();

            ModleRoot = Path.GetDirectoryName(this.GetType().Assembly.Location) + @"\..\..\..\..\..\ModleAsbuit";
            ModleRoot = Path.GetFullPath(ModleRoot);
            //MessageBox.Show();

            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            if (AssemModleDoc == null)
            {
                DoMode = 1;
            }
            else
            { 
                  //判断是否是这个程序对应的模型

                DoMode = 2;
            }
            SrcPath = ModleRoot + @"\第15章\Source";
            TargetPath = ModleRoot + @"\第15章\Result";
            rootpath = ModleRoot + @"\工程图模板";
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
                AssemModleDoc = (ModelDoc2)swApp.ActiveDoc;
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

        private void btn_domodle_Click(object sender, EventArgs e)
        {
            DoModle();
            OutPutDrawing();
        }

        #region 建模部分方法
        public void DoModle()
        {
            PlugOD = double.Parse(txt_od.Text.Trim());

            if (DoMode == 1)//新建
            {
                #region 复制需要的插座组件A，B,C
                string[] Names = new string[3] { "A", "B", "C" };
                string SourceTop = TargetPath + @"\SlotUnit\PlugSlot.SLDASM";
                string TargetTop = "";
                string[] SourceChildrenPaths=new string[1]{TargetPath + @"\SlotUnit\InnerPluge.SLDPRT"};
                int copyopt = (int)swMoveCopyOptions_e.swMoveCopyOptionsCreateNewFolder;
                foreach (string aa in Names)
                {
                    TargetTop = TargetPath + @"\SlotUnit" + aa + @"\PlugSlot" + aa + ".SLDASM";
                    string[] TargetChildrenPaths = new string[1] { TargetPath + @"\SlotUnit" + aa + @"\InnerPluge" + aa + ".SLDPRT" };
                    swApp.CopyDocument(SourceTop, TargetTop, SourceChildrenPaths, TargetChildrenPaths, copyopt);
                }
                #endregion

                #region 新建接线板主装配体
                int IntError = -1;
                int IntWraning = -1;
                AssemModleDoc = swApp.OpenDoc6(TargetPath + @"\PowerStrip.SLDASM", (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
                SelectionMgr SwSelMrg = AssemModleDoc.SelectionManager;
                SelectData SwSelData = SwSelMrg.CreateSelectData();
                FeatureManager SwFeatMrg = AssemModleDoc.FeatureManager;
                #endregion

                AssemModleDoc.ShowNamedView2("*等轴测", (int)swStandardViews_e.swIsometricView);
                AssemModleDoc.ViewZoomtofit2();

                #region 装配接线板主体
                DoAssem(AssemModleDoc, TargetPath + @"\PlugBottomBox.SLDPRT", (int)swDocumentTypes_e.swDocPART);//装配底壳
                DoAssem(AssemModleDoc, TargetPath + @"\PlugTopBox.SLDPRT", (int)swDocumentTypes_e.swDocPART);//装配顶壳  
                DoAssem(AssemModleDoc, TargetPath + @"\PlugWire.SLDPRT", (int)swDocumentTypes_e.swDocPART);//装配线缆          
                DoAssem(AssemModleDoc, TargetPath + @"\PlugButton.SLDPRT", (int)swDocumentTypes_e.swDocPART);//装配按钮
                DoAssem(AssemModleDoc, TargetPath + @"\PlugLED.SLDPRT", (int)swDocumentTypes_e.swDocPART);//装配指示灯
                DoAssem(AssemModleDoc, TargetPath + @"\PlugHead\PlugHead.SLDASM", (int)swDocumentTypes_e.swDocASSEMBLY);//装配插头组件
                #endregion
                //////////////
                #region 由内到外依次装二眼和五眼，再装组USB口装配所有插槽
                DoAssem(AssemModleDoc, TargetPath + @"\SlotUnitB\PlugSlotB.SLDASM", (int)swDocumentTypes_e.swDocASSEMBLY);//装二眼SlotUnitB
                DoAssem(AssemModleDoc, TargetPath + @"\SlotUnitA\PlugSlotA.SLDASM", (int)swDocumentTypes_e.swDocASSEMBLY);//装五眼SlotUnitA
                DoAssem(AssemModleDoc, TargetPath + @"\SlotUnitC\PlugSlotC.SLDASM", (int)swDocumentTypes_e.swDocASSEMBLY);//装USB口，SlotUnitC
                #endregion

                #region 圆周阵列插座A,B
                AssemModleDoc.EditRebuild3();
                AssemModleDoc.ClearSelection2(true);
                AssemModleDoc.Extension.SelectByID2("CenterAxis@PlugBottomBox-1@PowerStrip", "AXIS", 0, 0, 0, true, 2, null, 0);
                AssemModleDoc.Extension.SelectByID2("PlugSlotB-1@PowerStrip", "COMPONENT", 0, 0, 0, true, 1, null, 0);
                AssemModleDoc.Extension.SelectByID2("PlugSlotA-1@PowerStrip", "COMPONENT", 0, 0, 0, true, 1, null, 0);
                Feature SwFeat = SwFeatMrg.FeatureCircularPattern5(int.Parse(txt_slotABcount.Text.Trim()), Math.PI*2, false, "NULL", false, true, false, false, false, false, 1, 0.26179938779915, "NULL", false);
                SwFeat.Name = "SlotABPattern";
                #endregion

                #region 线性插座C
                AssemModleDoc.EditRebuild3();
                AssemModleDoc.ClearSelection2(true);
                double Temp = double.Parse(txt_toph.Text.Trim()) + double.Parse(txt_bottomh.Text.Trim()) - double.Parse(txt_tn.Text.Trim()) * 2 + 5;
                AssemModleDoc.Extension.SelectByID2("", "EDGE", (double.Parse(txt_btnr.Text.Trim()) - 5) / 1000.0, Temp / 1000.0, 0, true, 2, null, 0);
                AssemModleDoc.Extension.SelectByID2("", "EDGE", (double.Parse(txt_btnr.Text.Trim()) + 5) / 1000.0, Temp / 1000.0, 0, true, 4, null, 0);
                AssemModleDoc.Extension.SelectByID2("PlugSlotC-1@PowerStrip", "COMPONENT", 0, 0, 0, true, 1, null, 0);
                SwFeat = SwFeatMrg.FeatureLinearPattern5(2, 0.02, 2, 0.02, false, false, "NULL", "NULL", false, false, false, false, false, false, true, true, false, false, 0, 0, true, false);
                SwFeat.Name = "SlotCPattern";
                #endregion
                //AssemModleDoc.Save2(true);
                AssemModleDoc.EditRebuild3();

                CutTopBox(AssemModleDoc, SwSelMrg ,SwSelData,SwFeatMrg);//对顶盒开插座孔
            }
            else if (DoMode == 2)//修改
            { 
            
            
            
            
            }
        }
        public void DoAssem(ModelDoc2 AssemDoc, string FileName, int FileType, ModelDoc2 CompDoc=null)
        {
            SelectionMgr SwSelMrg = AssemDoc.SelectionManager;
            SelectData SwSelData = SwSelMrg.CreateSelectData();

            AssemblyDoc SwAssem = (AssemblyDoc)AssemDoc;
            #region 装配
            int IntError = -1;
            int IntWraning = -1;
            ModelDoc2 SwPartDoc = CompDoc;
            if (SwPartDoc == null)//说明部件还没打开
            {
                SwPartDoc = swApp.OpenDoc6(FileName, FileType, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
            }
            Component2 SwComp = ((AssemblyDoc)AssemDoc).AddComponent5(SwPartDoc.GetPathName(), 0, "", false, "", 0, 0, 0);
            
            if (SwComp.IsFixed())
            {
                SwComp.Select4(false, SwSelData, false);
                SwAssem.UnfixComponent();
                AssemDoc.ClearSelection2(true);
            }
            swApp.CloseDoc(SwPartDoc.GetTitle());//关闭文档
            if (((ModelDoc2)swApp.ActiveDoc).GetPathName() != AssemDoc.GetPathName())//判断关闭后，激活的文档是否是需要操作配合的文档
            {
                swApp.ActivateDoc3(AssemDoc.GetPathName(), true, 2, ref IntError);
            }
            
            DoMate(AssemDoc, FileName, SwComp);//装配零件
            AssemModleDoc.ShowNamedView2("*等轴测", (int)swStandardViews_e.swIsometricView);//轴测图便于观察
            AssemModleDoc.ViewZoomtofit2(); 
            RevisePart(SwComp,SwSelData);//修改模型尺寸
            #endregion
        }
        public void DoMate(ModelDoc2 AssemDoc, string FileName, Component2 SwComp)//装配
        {
            string AssemPathForSelectById2 = "";
            string CompPathForSelectById2 = "";
            string TempFeatureType = "";

            string CompFeatureName = "";
            string AssemFeatureName = "";
            #region 主体结构部件
            if (FileName.Contains(@"\PlugBottomBox.SLDPRT"))
            {
                CompPathForSelectById2 = "@" + SwComp.Name2 + "@" + AssemDoc.GetTitle().Substring(0, AssemDoc.GetTitle().IndexOf("."));
                swMateDis(AssemDoc, "前视基准面" + CompPathForSelectById2, "前视基准面" + AssemPathForSelectById2, 0, 0, "BottomBoxFrontMate",false);
                swMateDis(AssemDoc, "上视基准面" + CompPathForSelectById2, "上视基准面" + AssemPathForSelectById2, 0, 0, "BottomBoxTopMate", false);
                swMateDis(AssemDoc, "右视基准面" + CompPathForSelectById2, "右视基准面" + AssemPathForSelectById2, 0, 0, "BottomBoxRightMate", false);
            }
            else if (FileName.Contains(@"\PlugTopBox.SLDPRT"))
            {
                CompPathForSelectById2 = "@" + SwComp.Name2 + "@" + AssemDoc.GetTitle().Substring(0, AssemDoc.GetTitle().IndexOf("."));
                Component2 CompToMate = FindComp(AssemDoc, "PlugBottomBox");
                AssemPathForSelectById2 = "@" + CompToMate.Name2 + "@" + AssemDoc.GetTitle().Substring(0, AssemDoc.GetTitle().IndexOf("."));
                swMateDis(AssemDoc, "前视基准面" + CompPathForSelectById2, "前视基准面" + AssemPathForSelectById2, 0, 0, "TopBoxFrontMate", false);
                swMateDis(AssemDoc, "ConnectFace" + CompPathForSelectById2, "ConnectFace" + AssemPathForSelectById2, 0, 0, "TopBoxConnectMate", false);
                swMateDis(AssemDoc, "右视基准面" + CompPathForSelectById2, "右视基准面" + AssemPathForSelectById2, 0, 0, "TopBoxRightMate", false);
            }
            else if (FileName.Contains( @"\PlugWire.SLDPRT"))
            {
                CompPathForSelectById2 = "@" + SwComp.Name2 + "@" + AssemDoc.GetTitle().Substring(0, AssemDoc.GetTitle().IndexOf("."));
                Component2 CompToMate = FindComp(AssemDoc, "PlugBottomBox");
                AssemPathForSelectById2 = "@" + CompToMate.Name2 + "@" + AssemDoc.GetTitle().Substring(0, AssemDoc.GetTitle().IndexOf("."));
                swMateAXIS(AssemDoc, "CenterAxis" + CompPathForSelectById2, "CenterAxis" + AssemPathForSelectById2, 0, "WireAxiMate");
                swMateDis(AssemDoc, "WireCenter" + CompPathForSelectById2, "ConnectFace" + AssemPathForSelectById2, 1, 0, "WireConnectMate", false);
                swMateAng(AssemDoc, "CenterH" + CompPathForSelectById2, "BoxCenterV" + AssemPathForSelectById2, 0, Math.PI/2.0,false,"WireAngMate");
            }
            else if (FileName.Contains( @"\PlugButton.SLDPRT"))
            {
                CompPathForSelectById2 = "@" + SwComp.Name2 + "@" + AssemDoc.GetTitle().Substring(0, AssemDoc.GetTitle().IndexOf("."));
                Component2 CompToMate = FindComp(AssemDoc, "PlugBottomBox");
                AssemPathForSelectById2 = "@" + CompToMate.Name2 + "@" + AssemDoc.GetTitle().Substring(0, AssemDoc.GetTitle().IndexOf("."));
                swMateAXIS(AssemDoc, "CenterAxis" + CompPathForSelectById2, "CenterAxis" + AssemPathForSelectById2, 1, "ButtonAxiMate");
                swMateAng(AssemDoc, "CenterH" + CompPathForSelectById2, "BoxCenterV" + AssemPathForSelectById2, 0, Math.PI / 2.0,true, "ButtonAngMate");
                CompToMate = FindComp(AssemDoc, "PlugTopBox");
                AssemPathForSelectById2 = "@" + CompToMate.Name2 + "@" + AssemDoc.GetTitle().Substring(0, AssemDoc.GetTitle().IndexOf("."));
                swMateDis(AssemDoc, "ButtonTop" + CompPathForSelectById2, "BoxInnerTop" + AssemPathForSelectById2, 0, (double.Parse(txt_tn.Text.Trim())+3) / 1000.0, "ButtonFaceMate", false);
            }
            else if (FileName.Contains( @"\PlugLED.SLDPRT"))
            {
                CompPathForSelectById2 = "@" + SwComp.Name2 + "@" + AssemDoc.GetTitle().Substring(0, AssemDoc.GetTitle().IndexOf("."));
                Component2 CompToMate = FindComp(AssemDoc, "PlugBottomBox");
                AssemPathForSelectById2 = "@" + CompToMate.Name2 + "@" + AssemDoc.GetTitle().Substring(0, AssemDoc.GetTitle().IndexOf("."));
                swMateAXIS(AssemDoc, "CenterAxis" + CompPathForSelectById2, "CenterAxis" + AssemPathForSelectById2,1, "LEDAxiMate");
                swMateAng(AssemDoc, "CenterH" + CompPathForSelectById2, "BoxCenterV" + AssemPathForSelectById2, 0, Math.PI / 2.0, true, "LEDAngMate");
                CompToMate = FindComp(AssemDoc, "PlugTopBox");
                AssemPathForSelectById2 = "@" + CompToMate.Name2 + "@" + AssemDoc.GetTitle().Substring(0, AssemDoc.GetTitle().IndexOf("."));
                swMateDis(AssemDoc, "LEDTop" + CompPathForSelectById2, "BoxInnerTop" + AssemPathForSelectById2, 0, double.Parse(txt_tn.Text.Trim()) / 1000.0, "LEDFaceMate", false);
            }
            else if (FileName.Contains(@"\PlugHead.SLDASM"))//路径比较深的时候拼接路径不方便，可以使用GetNameForSelection方法获得名称
            {
                Component2 CompToMate = FindComp(AssemDoc, "PlugPinHead");
                Component2 AssemToMate = FindComp(AssemDoc, "PlugWire");

                CompFeatureName = CompToMate.FeatureByName("WireConnectFace").GetNameForSelection(out TempFeatureType);
                AssemFeatureName = AssemToMate.FeatureByName("EndFace").GetNameForSelection(out TempFeatureType);
                swMateDis(AssemDoc, CompFeatureName, AssemFeatureName, 1, 0, "PlugHeadEndMate", true);

                CompFeatureName = CompToMate.FeatureByName("ConnectFace").GetNameForSelection(out TempFeatureType);
                AssemFeatureName = AssemToMate.FeatureByName("WireCenter").GetNameForSelection(out TempFeatureType);
                swMateDis(AssemDoc, CompFeatureName, AssemFeatureName, 0, 10/1000.0, "PlugHeadHCenterMate",true);

                CompFeatureName = CompToMate.FeatureByName("CenterV").GetNameForSelection(out TempFeatureType);
                AssemFeatureName = AssemToMate.FeatureByName("EndV").GetNameForSelection(out TempFeatureType);
                swMateDis(AssemDoc, CompFeatureName, AssemFeatureName, 0, 0, "PlugHeadVCenterMate", true);
            }
            #endregion

            #region 插座
            else if (FileName.Contains(@"\PlugSlot"))//因为源头来自一个源文件模板，所以装配方式一致
            {
                string SlotMark = FileName.Substring(FileName.LastIndexOf(".") - 1, 1);//提取A，B,C
                Component2 CompToMate = FindComp(AssemDoc, "InnerPluge"+SlotMark);
                Component2 AssemToMate = FindComp(AssemDoc, "PlugBottomBox");

                CompFeatureName = CompToMate.FeatureByName("CenterAxis").GetNameForSelection(out TempFeatureType);
                AssemFeatureName = AssemToMate.FeatureByName("CenterAxis").GetNameForSelection(out TempFeatureType);
                swMateAXIS(AssemDoc, CompFeatureName, AssemFeatureName, 1, "InnerPluge"+SlotMark+"_AxiMate");

                double Ang = 0;
                if (SlotMark == "A" || SlotMark == "B")
                {
                    Ang = 360.0 / int.Parse(txt_slotABcount.Text.Trim());
                    Ang = Ang / 2.0;
                }
                else if (SlotMark == "C")
                {
                    Ang = 180;
                }
                Ang = Ang * Math.PI / 180.0;//转化为弧度

                CompFeatureName = CompToMate.FeatureByName("SlotCenterV").GetNameForSelection(out TempFeatureType);
                AssemFeatureName = AssemToMate.FeatureByName("BoxCenterH").GetNameForSelection(out TempFeatureType);
                swMateAng(AssemDoc, CompFeatureName, AssemFeatureName, 1, Ang, false, "InnerPluge" + SlotMark + "_AngMate");

                CompFeatureName = CompToMate.FeatureByName("InnerBottom").GetNameForSelection(out TempFeatureType);
                AssemFeatureName = AssemToMate.FeatureByName("BoxInnerBottom").GetNameForSelection(out TempFeatureType);
                swMateDis(AssemDoc, CompFeatureName, AssemFeatureName, 0, 0, "InnerPluge" + SlotMark + "_BaseMate", true);
            }
            #endregion
        }
        public void RevisePart(Component2 SwComp,SelectData SwSelData)//修改零件大小
        {
            Dictionary<string, string> EquationStrDic = new Dictionary<string, string>();//方程式修改，Dictionary<等式左边, 整个公式>
            ModelDoc2 CompDoc = SwComp.GetModelDoc2();
            Component2 SwChildComp = null;
            Feature SwFeat = null;
            ModelDoc2 DocForEquation =CompDoc;//记录方程式操作所用文档对象
            Component2 CompForEquation = SwComp;//记录方程式操作所用部件
            CustomPropertyManager SwCusp = CompDoc.Extension.CustomPropertyManager[""];
            string DocTitle = "";
            if (SwComp.Name2.Contains("PlugBottomBox")||SwComp.Name2.Contains("PlugTopBox"))
            {
                SwFeat = ((PartDoc)CompDoc).FeatureByName("Rectangle");
                SwFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration,"");
                SwFeat = ((PartDoc)CompDoc).FeatureByName("Circle");
                SwFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, "");
                CompDoc.Parameter("D1@SketchCircle").SystemValue = PlugOD / 1000.0;
                double tn=double.Parse(txt_tn.Text.Trim());
                if (SwComp.Name2.Contains("PlugBottomBox"))
                {
                    EquationStrDic.Add("H", "\"H\"="+(double.Parse(txt_bottomh.Text.Trim())-tn).ToString().Trim());
                    EquationStrDic.Add("Tn", "\"Tn\"="+tn.ToString().Trim());

                    DocTitle = CompDoc.GetTitle();
                    SwCusp.Add3("图号", (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    string Name = "底盒 Φ" + "\"D1@SketchCircle@" + DocTitle + "\"" + "X" + "\"D2@CircleBox@" + DocTitle + "\"t";
                    SwCusp.Add3("名称", (int)swCustomInfoType_e.swCustomInfoText, Name, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, "PVC", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("单重", (int)swCustomInfoType_e.swCustomInfoText, "\"SW-Mass@" + DocTitle + "\"", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("类型", (int)swCustomInfoType_e.swCustomInfoText, "顶盒", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                }
                else if (SwComp.Name2.Contains("PlugTopBox"))
                {
                    EquationStrDic.Add("H", "\"H\"=" + (double.Parse(txt_toph.Text.Trim()) - tn).ToString().Trim());
                    EquationStrDic.Add("Tn", "\"Tn\"="+tn.ToString().Trim());

                    DocTitle = CompDoc.GetTitle();
                    SwCusp.Add3("图号", (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    string Name = "顶盒 Φ" + "\"D1@SketchCircle@" + DocTitle + "\"" + "X" + "\"D2@CircleBox@" + DocTitle + "\"t";
                    SwCusp.Add3("名称", (int)swCustomInfoType_e.swCustomInfoText, Name, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, "PVC", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("单重", (int)swCustomInfoType_e.swCustomInfoText, "\"SW-Mass@" + DocTitle + "\"", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("类型", (int)swCustomInfoType_e.swCustomInfoText, "顶盒", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                }
            }
            else if (SwComp.Name2.Contains("PlugWire"))
            {
                CompDoc.Parameter("R@Circle").SystemValue =(( PlugOD/2.0 )-3)/1000.0;

                DocTitle = CompDoc.GetTitle();
                SwCusp.Add3("图号", (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                string Name = "线缆 Φ" + "\"D1@草图1@" + DocTitle + "\"";
                SwCusp.Add3("名称", (int)swCustomInfoType_e.swCustomInfoText, Name, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                SwCusp.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, "黄铜", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                SwCusp.Add3("单重", (int)swCustomInfoType_e.swCustomInfoText, "\"SW-Mass@" + DocTitle + "\"", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                SwCusp.Add3("类型", (int)swCustomInfoType_e.swCustomInfoText, "线缆", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
            }
            else if (SwComp.Name2.Contains("PlugButton"))
            {
                CompDoc.Parameter("R@Circle").SystemValue = double.Parse(txt_btnr.Text.Trim()) / 1000.0;

                DocTitle = CompDoc.GetTitle();
                SwCusp.Add3("图号", (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                string Name = "按钮 Φ" + "\"D2@SRectangle@" + DocTitle + "\"" + "X" + "\"D1@SRectangle@" + DocTitle + "\"" + " H=" + "\"D1@FRectangle@" + DocTitle + "\"";
                SwCusp.Add3("名称", (int)swCustomInfoType_e.swCustomInfoText, Name, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                SwCusp.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, "PVC", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                SwCusp.Add3("单重", (int)swCustomInfoType_e.swCustomInfoText, "\"SW-Mass@" + DocTitle + "\"", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                SwCusp.Add3("类型", (int)swCustomInfoType_e.swCustomInfoText, "按钮", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
            }
            else if (SwComp.Name2.Contains("PlugLED"))
            {
                CompDoc.Parameter("R@Circle").SystemValue = double.Parse(txt_ledr.Text.Trim()) / 1000.0;

                DocTitle = CompDoc.GetTitle();
                SwCusp.Add3("图号", (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                string Name = "LED指示灯 Φ" + "\"D2@SRectangle@" + DocTitle + "\"" + "X" + "\"D1@SRectangle@" + DocTitle + "\"" ;
                SwCusp.Add3("名称", (int)swCustomInfoType_e.swCustomInfoText, Name, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                SwCusp.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, "外购件", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                SwCusp.Add3("单重", (int)swCustomInfoType_e.swCustomInfoText, "\"SW-Mass@" + DocTitle + "\"", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                SwCusp.Add3("类型", (int)swCustomInfoType_e.swCustomInfoText, "指示灯", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
            }
            else if (SwComp.Name2.Contains("PlugSlot"))//插座组件-->需要找到它的子件
            {
                string SlotMark=SwComp.Name2.Substring(SwComp.Name2.LastIndexOf("-")-1,1);//提取A,B,C
                //SwChildComp = ((AssemblyDoc)SwComp.GetModelDoc2()).GetComponentByName("InnerPluge" + SlotMark + "-1");
                SwChildComp = (Component2)(SwComp.FeatureByName("InnerPluge" + SlotMark + "-1")).GetSpecificFeature2();
                ModelDoc2 SwChildDoc = SwChildComp.GetModelDoc2();
                DocForEquation =SwChildDoc;
                CompForEquation = SwChildComp;
                double R = 0;//中心距
                if (SlotMark == "B")
                {
                    R = double.Parse(txt_br.Text.Trim());
                    SwFeat = ((PartDoc)SwChildDoc).FeatureByName("FiveSlotTP1");
                    SwFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, "");
                    SwFeat = ((PartDoc)SwChildDoc).FeatureByName("TwoSlot");
                    SwFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, "");
                    SwFeat = ((PartDoc)SwChildDoc).FeatureByName("UsbSlot");
                    SwFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, "");

                    DocTitle = CompDoc.GetTitle();
                    SwCusp.Add3("图号", (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("名称", (int)swCustomInfoType_e.swCustomInfoText, "两孔插座组件", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, "组合件", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("单重", (int)swCustomInfoType_e.swCustomInfoText, "\"SW-Mass@" + DocTitle + "\"", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("类型", (int)swCustomInfoType_e.swCustomInfoText, "两孔插座", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);

                    SwCusp = SwChildDoc.Extension.CustomPropertyManager[""];
                    DocTitle = SwChildDoc.GetTitle();
                    SwCusp.Add3("图号", (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("名称", (int)swCustomInfoType_e.swCustomInfoText, "两孔基座", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, "ABS", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("单重", (int)swCustomInfoType_e.swCustomInfoText, "\"SW-Mass@" + DocTitle + "\"", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("类型", (int)swCustomInfoType_e.swCustomInfoText, "两孔基座", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                }
                else if (SlotMark == "A")
                {
                    R = double.Parse(txt_ar.Text.Trim());
                    SwFeat = ((PartDoc)SwChildDoc).FeatureByName("FiveSlotTP1");
                    SwFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, "");
                    SwFeat = ((PartDoc)SwChildDoc).FeatureByName("TwoSlot");
                    SwFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, "");
                    SwFeat = ((PartDoc)SwChildDoc).FeatureByName("UsbSlot");
                    SwFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, "");

                    DocTitle = CompDoc.GetTitle();
                    SwCusp.Add3("图号", (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("名称", (int)swCustomInfoType_e.swCustomInfoText, "五孔插座组件", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, "组合件", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("单重", (int)swCustomInfoType_e.swCustomInfoText, "\"SW-Mass@" + DocTitle + "\"", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("类型", (int)swCustomInfoType_e.swCustomInfoText, "五孔插座", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);

                    SwCusp = SwChildDoc.Extension.CustomPropertyManager[""];
                    DocTitle = SwChildDoc.GetTitle();
                    SwCusp.Add3("图号", (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("名称", (int)swCustomInfoType_e.swCustomInfoText, "五孔基座", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, "ABS", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("单重", (int)swCustomInfoType_e.swCustomInfoText, "\"SW-Mass@" + DocTitle + "\"", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("类型", (int)swCustomInfoType_e.swCustomInfoText, "五孔基座", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                }
                else if (SlotMark == "C")
                {
                    R = double.Parse(txt_cr.Text.Trim());
                    SwFeat = ((PartDoc)SwChildDoc).FeatureByName("FiveSlotTP1");
                    SwFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, "");
                    SwFeat = ((PartDoc)SwChildDoc).FeatureByName("TwoSlot");
                    SwFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, "");
                    SwFeat = ((PartDoc)SwChildDoc).FeatureByName("UsbSlot");
                    SwFeat.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swAllConfiguration, "");

                    DocTitle = CompDoc.GetTitle();
                    SwCusp.Add3("图号", (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("名称", (int)swCustomInfoType_e.swCustomInfoText, "USB插座组件", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, "组合件", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("单重", (int)swCustomInfoType_e.swCustomInfoText, "\"SW-Mass@" + DocTitle + "\"", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("类型", (int)swCustomInfoType_e.swCustomInfoText, "USB插座", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);

                    SwCusp = SwChildDoc.Extension.CustomPropertyManager[""];
                    DocTitle = SwChildDoc.GetTitle();
                    SwCusp.Add3("图号", (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("名称", (int)swCustomInfoType_e.swCustomInfoText, "USB基座", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, "ABS", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("单重", (int)swCustomInfoType_e.swCustomInfoText, "\"SW-Mass@" + DocTitle + "\"", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    SwCusp.Add3("类型", (int)swCustomInfoType_e.swCustomInfoText, "USB基座", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                }
                if (R == 0)
                {
                    R = 0.001;//防止报错
                }
                SwChildDoc.Parameter("R@Circle").SystemValue = R/ 1000.0;
                double BaseH = double.Parse(txt_bottomh.Text.Trim()) + double.Parse(txt_toph.Text.Trim()) - 2 * double.Parse(txt_tn.Text.Trim()) - 1;
                EquationStrDic.Add("BaseH", "\"BaseH\"=" + BaseH.ToString().Trim());  
            }
            /////////////
            if (EquationStrDic.Count > 0)//说明有方程式需要修改
            {
                bool sc = CompForEquation.Select4(false, SwSelData, false);//选中编辑零部件状态才能更新方程
                int x = -1;
                //MessageBox.Show(CompForEquation.Name2);
                ((AssemblyDoc)AssemModleDoc).EditPart2(true, true, ref x);
                List<string> EquationDone=new List<string>();//记录完成的方程
                EquationMgr SwEquationMgr = DocForEquation.GetEquationMgr();
                string Left = "";//记录等号左边部分
                for (int i = 0; i < SwEquationMgr.GetCount(); i++)
                {
                    Left = SwEquationMgr.Equation[i].Substring(1, SwEquationMgr.Equation[i].IndexOf("=")-2);//去掉双引号
                    if (EquationStrDic.Keys.Contains(Left))
                    {
                        SwEquationMgr.Equation[i] = EquationStrDic[Left];//修改
                        EquationDone.Add(Left);//记录完成
                    }
                }
                /////
                if (EquationStrDic.Count > EquationDone.Count)//说明还有新增的方程式
                {
                    foreach (string key in EquationStrDic.Keys)
                    {
                        if (EquationDone.Contains(key) == false)//找到需要新增的方程式
                        {
                            SwEquationMgr.Add3(-1, EquationStrDic[key], true, (int)swInConfigurationOpts_e.swAllConfiguration, DocForEquation.GetConfigurationNames());
                        }
                    }
                }
                SwEquationMgr.EvaluateAll();
                DocForEquation.EditRebuild3();
                ((AssemblyDoc)AssemModleDoc).EditAssembly();
            }
            //CompDoc.Save2(true);
            CompDoc.EditRebuild3();
        }
        public void CutTopBox(ModelDoc2 AssemDoc, SelectionMgr SwSelMrg ,SelectData SwSelData, FeatureManager SwFeatMrg)
        {
            SketchManager SwSketch=AssemDoc.SketchManager;
            Dictionary<string, string> SketchsForCut = new Dictionary<string, string>();
            SketchsForCut.Add("PlugButton", "RectangleCut");
            SketchsForCut.Add("PlugLED", "RectangleCut");
            SketchsForCut.Add("InnerPlugeB", "STwoSlot");
            SketchsForCut.Add("InnerPlugeA", "SFiveSlotTP1");
            SketchsForCut.Add("InnerPlugeC", "SUsbSlot");

            Component2 TopBoxComp = ((AssemblyDoc)AssemDoc).GetComponentByName("PlugTopBox-1");
            TopBoxComp.Select4(false, SwSelData, false);
            ((AssemblyDoc)AssemDoc).EditPart();
        
            Feature SwSketchFace = ((AssemblyDoc)AssemDoc).GetComponentByName("PlugTopBox-1").FeatureByName("ConnectFace");

            object[] ObjComps = ((AssemblyDoc)AssemModleDoc).GetComponents(true);
            foreach (object objComp in ObjComps)
            {
                AssemDoc.ClearSelection2(true);
                Component2 SwComp = (Component2)objComp;
                Feature CutSketchFeat = null;
               
                if (SwComp.Name2.Contains("PlugButton"))
                {
                    CutSketchFeat = SwComp.FeatureByName(SketchsForCut["PlugButton"]);
                }
                else if (SwComp.Name2.Contains("PlugLED"))
                {
                    CutSketchFeat = SwComp.FeatureByName(SketchsForCut["PlugLED"]);
                }
                else if (SwComp.Name2.Contains("PlugSlotA"))
                {
                    CutSketchFeat = ((Component2)(SwComp.FeatureByName("InnerPlugeA-1").GetSpecificFeature2())).FeatureByName(SketchsForCut["InnerPlugeA"]);
                }
                else if (SwComp.Name2.Contains("PlugSlotB"))
                {
                    CutSketchFeat = ((Component2)(SwComp.FeatureByName("InnerPlugeB-1").GetSpecificFeature2())).FeatureByName(SketchsForCut["InnerPlugeB"]);
                }
                else if (SwComp.Name2.Contains("PlugSlotC"))
                {
                    CutSketchFeat = ((Component2)(SwComp.FeatureByName("InnerPlugeC-1").GetSpecificFeature2())).FeatureByName(SketchsForCut["InnerPlugeC"]);
                }
                ////
                if (CutSketchFeat == null)//说明不是需要切除上壳的零件
                {
                    continue;
                }
                SwSketchFace.Select2(false, 0);//选中草图平面
                SwSketch.InsertSketch(true);//新建草图
                CutSketchFeat.Select2(false, 0);//选中切割草图
                SwSketch.SketchUseEdge2(false);//转化实体引用
                AssemDoc.ClearSelection2(true);//清除选择
                ((Feature)SwSketch.ActiveSketch).Name = SwComp.Name2 + "_Sketch";//给草图重命名

                bool Dir = false;//拉伸切除方向
                Feature SwCutFeat = SwFeatMrg.FeatureCut4(true, false, Dir, (int)swEndConditions_e.swEndCondThroughAll, 0, 0.01, 0.01, false, false, false, false, 0, 0, false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);
                if (SwCutFeat == null)//说明拉伸方向有问题,让Dir取反 再拉伸切除
                {
                    SwCutFeat = SwFeatMrg.FeatureCut4(true, false, !Dir, (int)swEndConditions_e.swEndCondThroughAll, 0, 0.01, 0.01, false, false, false, false, 0, 0, false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);
                }
                SwCutFeat.Name = SwComp.Name2 + "_Cut";
            }

            ((AssemblyDoc)AssemDoc).EditAssembly();//返回到编辑装配体，即退出编辑零件或部件
        }
        public Component2 FindComp(ModelDoc2 AssemDoc,string CompName)
        {
            Component2 SwComp = null;
            for (int i = 1; i < 20; i++)
            {
                SwComp = ((AssemblyDoc)AssemDoc).GetComponentByName(CompName + "-" + i.ToString().Trim());
                if (SwComp != null)
                {
                    break;
                }
            }
            return SwComp;
        }

        #region 装配公共函数
        public void swMateAXIS(ModelDoc2 ZongAssemM, string SubAixName1, string SubAixName2, int types, string MateName)//与总装配
        {
             /*ZongAssemM=总装配体对象
             * types=0,1正反，调试得到------
             * MateName=该装配特征的名字
             * */
            int temp = 0;
            ZongAssemM.ShowNamedView2("*前视", 1);
            ZongAssemM.Extension.SelectByID2(SubAixName1, "AXIS", 0, 0, 0, false, 1, null, 0);
            ZongAssemM.Extension.SelectByID2(SubAixName2, "AXIS", 0, 0, 0, true, 1, null, 0);
            Feature swMateFeature;
            swMateFeature = (Feature)((AssemblyDoc)ZongAssemM).AddMate3(0, types, false, 0, 0, 0, 0.001, 0.001, 0, 0, 0, false, out temp);

            if (swMateFeature != null)
            {
                swMateFeature.Name = MateName;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(MateName + ":中心轴装配出错");
            }

            AssemModleDoc.EditRebuild3();
            ZongAssemM.ViewZoomtofit2();
            ZongAssemM.ClearSelection2(true);
        }
        public void swMateAng(ModelDoc2 ZongAssemM, string SubPlaneName1, string SubPlaneName2, int a5, double jiaodu, bool zhengfan, string MateName)//与总装配的
        {
             /*ZongAssemM=总装配体对象
             * a5=0,1正反，调试得到-------
             * jiaodu=角度对应的弧度值
             * zhengfan=正反，调试得到--------
             * MateName=该装配特征的名字
             * */
            int temp = 0;
            ZongAssemM.ShowNamedView2("*前视", 1);
            ZongAssemM.Extension.SelectByID2(SubPlaneName2, "PLANE", 0, 0, 0, false, 1, null, 0);
            ZongAssemM.Extension.SelectByID2(SubPlaneName1, "PLANE", 0, 0, 0, true, 1, null, 0);

            Feature swMateFeature;
            string angstart = ((CustomPropertyManager)ZongAssemM.Extension.CustomPropertyManager[""]).Get("角度起始");
            if (angstart == "Right")
            {
                swMateFeature = (Feature)((AssemblyDoc)ZongAssemM).AddMate3(6, a5, zhengfan, 0, 0, 0, 0, 0, jiaodu, jiaodu, jiaodu, false, out temp); //把选中的装配起来
            }
            else
            {
                swMateFeature = (Feature)((AssemblyDoc)ZongAssemM).AddMate3(6, a5, !zhengfan, 0, 0, 0, 0, 0, jiaodu, jiaodu, jiaodu, false, out temp);//把选中的装配起来
            }

            if (swMateFeature != null)
            {
                swMateFeature.Name = MateName;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(MateName + ":方位装配出错");
            }
            AssemModleDoc.EditRebuild3();
        }
        public void swMateDis(ModelDoc2 ZongAssemM, string SubPlaneName1, string SubPlaneName2, int a5, double dis, string MateName, bool zhengfan)//与总装配的 基准配合面 配距离,即主体下焊缝线
        {
             /*ZongAssemM=总装配体对象
             * a5=0,1正反，调试得到-------一般为0
             * dis=配合距离
             * MateName=该装配特征的名字
             * */
            int temp = 0;
            //ZongAssemM.ShowNamedView2("*前视", 1);
            ZongAssemM.Extension.SelectByID2(SubPlaneName2, "PLANE", 0, 0, 0, false, 1, null, 0);
            ZongAssemM.Extension.SelectByID2(SubPlaneName1, "PLANE", 0, 0, 0, true, 1, null, 0);

            int MateType = 5;
            if (dis == 0)//重合
            {
                MateType =0;
            }

            Feature swMateFeature;
            if (dis < 0)
            {
                dis = Math.Abs(dis);
            }
            swMateFeature = (Feature)((AssemblyDoc)ZongAssemM).AddMate5(MateType, a5, zhengfan, dis, dis, dis, 0, 0, 0, 0, 0, false, false, 0, out temp);

            if (swMateFeature != null)
            {
                swMateFeature.Name = MateName;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(MateName + ":距离出错");
            }
            AssemModleDoc.EditRebuild3();
        }
        #endregion

        #endregion

        #region 出图部分方法
        public void OutPutDrawing()
        {
            string TemplateName = rootpath + @"\A1模板.DRWDOT";//工程图模板路径--包含了属性和文档设置信息
            string DwgFormatePath = rootpath + @"\A1图纸格式.slddrt";//图纸格式模板
            string DrawStdPath = rootpath + @"\绘图标准.sldstd";
            string AssemTitleBlock = rootpath + @"\BLOCK_TITLE_ASSEM.SLDBLK";
            string PartTitleBlock = rootpath + @"\BLOCK_TITLE_PART.SLDBLK";
            string PartHeadBlock = rootpath + @"\PartTitle.SLDBLK";
            string BomPath = rootpath + @"\BomTopRight.sldbomtbt";
            //***************
            DrawingModleDoc = swApp.NewDocument(TemplateName, 10, 0, 0);
            DrawingModleDoc.SaveAs(AssemModleDoc.GetPathName().Substring(0, AssemModleDoc.GetPathName().LastIndexOf("."))+".SLDDRW");
            DrawingDoc SwDraw = (DrawingDoc)DrawingModleDoc;
            SelectionMgr SwSelectionMgr=DrawingModleDoc.SelectionManager;
            SelectData SwSelectData= SwSelectionMgr.CreateSelectData();
            SketchManager SwSketchMrg = DrawingModleDoc.SketchManager;
            MathUtility SwMathUtility = swApp.GetMathUtility();
            MathPoint SwMathPoint = null;
            DrawingModleDoc.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swViewDisplayHideAllTypes, true);
            SwDraw.EditTemplate();
            SwMathPoint = SwMathUtility.CreatePoint(new double[] { 829 / 1000.0, 12 / 1000.0, 0 });
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, AssemTitleBlock, false, 1, 0);
            SwDraw.EditSheet();


            #region 绘制总图
            double[] viewpos = new double[] { 240 / 1000.0, 180 / 1000.0 };//定义SwView1插入点
            SolidWorks.Interop.sldworks.View SwView1 = SwDraw.CreateDrawViewFromModelView3(AssemModleDoc.GetPathName(), "*上视", viewpos[0], viewpos[1], 0);
            SwView1.InsertBomTable4(false, 829 / 1000.0, 72 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight, (int)swBomType_e.swBomType_TopLevelOnly, "默认", BomPath, false, (int)swNumberingType_e.swNumberingType_Flat, false);

            viewpos = new double[] { 240/ 1000.0, 400 / 1000.0 };//定义SwView1插入点
            SolidWorks.Interop.sldworks.View SwView2 = SwDraw.CreateDrawViewFromModelView3(AssemModleDoc.GetPathName(), "*前视", viewpos[0], viewpos[1], 0);

            viewpos = new double[] { 600 / 1000.0, 400 / 1000.0 };//定义SwView1插入点
            SolidWorks.Interop.sldworks.View SwView3 = SwDraw.CreateDrawViewFromModelView3(AssemModleDoc.GetPathName(), "*等轴测", viewpos[0], viewpos[1], 0);
            DrawingComponent DcTopBox = GetDrawingComp(SwView3.RootDrawingComponent, "PlugTopBox", SwView3);
            DcTopBox.Select(false, SwSelectData);
            DrawingModleDoc.HideComponent2();

            #region 自动拉件号
            SwDraw.ActivateView(SwView3.Name);
            DrawingModleDoc.Extension.SelectByID2(SwView3.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            AutoBalloonOptions SwAutoBalloonOptions = SwDraw.CreateAutoBalloonOptions();
            SwAutoBalloonOptions.Layout = 1;
            SwAutoBalloonOptions.ReverseDirection = false;
            SwAutoBalloonOptions.IgnoreMultiple = true;
            SwAutoBalloonOptions.InsertMagneticLine = true;
            SwAutoBalloonOptions.LeaderAttachmentToFaces = false;
            SwAutoBalloonOptions.Style = 1;
            SwAutoBalloonOptions.Size = 2;
            SwAutoBalloonOptions.EditBalloonOption = 1;
            SwAutoBalloonOptions.EditBalloons = true;
            SwAutoBalloonOptions.UpperTextContent = 1;
            SwAutoBalloonOptions.UpperText = "";
            SwAutoBalloonOptions.Layername = "Text";
            SwAutoBalloonOptions.ItemNumberStart = 1;
            SwAutoBalloonOptions.ItemNumberIncrement = 1;
            SwAutoBalloonOptions.ItemOrder = 0;
            SwDraw.AutoBalloon5(SwAutoBalloonOptions);
            #endregion

            #endregion

            #region 绘制零件图
            SwDraw.NewSheet3("零件图1", 12, 12, 1, 1, true, DwgFormatePath, 0.42, 0.297, "默认");
            SwDraw.EditTemplate();
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, PartTitleBlock, false, 1, 0);
            SwSketchMrg.CreateLine(24 / 1000.0, 294.5 / 1000.0, 0, 829 / 1000.0, 294.5 / 1000.0, 0);
            SwSketchMrg.CreateLine(426.5 / 1000.0, 12 / 1000.0, 0, 426.5 / 1000.0, 577 / 1000.0, 0);
            SwSketchMrg.CreateLine(627.75 / 1000.0, 12 / 1000.0, 0, 627.75 / 1000.0, 294.5 / 1000.0, 0);
            SwDraw.EditSheet();
            
            int x = -1;

            swApp.ActivateDoc3(AssemModleDoc.GetPathName(), true, 2, ref x);
            Component2 SwComp = ((AssemblyDoc)AssemModleDoc).GetComponentByName("PlugTopBox-1");
            swApp.ActivateDoc3(SwComp.GetPathName(), true, 2, ref x);
            swApp.ActivateDoc3(DrawingModleDoc.GetPathName(), true, 2, ref x);
            viewpos = new double[] {200 / 1000.0, 450/ 1000.0 };//定义SwView1插入点
            SolidWorks.Interop.sldworks.View SwView4 = SwDraw.CreateDrawViewFromModelView3(SwComp.GetPathName(), "*下视", viewpos[0], viewpos[1], 0);
            SwView4.ScaleRatio = new double[] { 1, 2 };
            SwView4.InsertBomTable4(false, 426.5 / 1000.0, 312.5 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight, (int)swBomType_e.swBomType_TopLevelOnly, "默认", BomPath, false, (int)swNumberingType_e.swNumberingType_Flat, false);
            swApp.CloseDoc(SwComp.GetPathName());
            ////////
            SwDraw.ActivateView(SwView4.Name);
            double[] SheetPos = new double[2] { 426.5 / 1000.0, 294.5 / 1000.0 };
            SwMathPoint = SwMathUtility.CreatePoint(CalPartHeadPosInView(SheetPos, viewpos, SwView4.ScaleDecimal));
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, PartHeadBlock, false, 1, 0);


            swApp.ActivateDoc3(AssemModleDoc.GetPathName(), true, 2, ref x);
            SwComp = ((AssemblyDoc)AssemModleDoc).GetComponentByName("PlugHead-1");
            swApp.ActivateDoc3(SwComp.GetPathName(), true, 2, ref x);
            swApp.ActivateDoc3(DrawingModleDoc.GetPathName(), true, 2, ref x);
            viewpos = new double[] { 520 / 1000.0, 480 / 1000.0 };//定义SwView1插入点
            SolidWorks.Interop.sldworks.View SwView5= SwDraw.CreateDrawViewFromModelView3(SwComp.GetPathName(), "*前视", viewpos[0], viewpos[1], 0);
            SwView5.ScaleRatio = new double[] { 1, 1 };
            
            viewpos = new double[] { 650 / 1000.0, 480 / 1000.0 };//定义SwView1插入点
            SolidWorks.Interop.sldworks.View SwView5Left = SwDraw.CreateDrawViewFromModelView3(SwComp.GetPathName(), "*左视", viewpos[0], viewpos[1], 0);
            SwView5Left.ScaleRatio = new double[] { 1, 1 };
            
            viewpos = new double[] { 520 / 1000.0, 380 / 1000.0 };//定义SwView1插入点
            SolidWorks.Interop.sldworks.View SwView5Front = SwDraw.CreateDrawViewFromModelView3(SwComp.GetPathName(), "*上视", viewpos[0], viewpos[1], 0);
            SwView5Front.ScaleRatio = new double[] { 1, 1 };
            SwView5Front.SetDisplayMode3(false, (int)swDisplayMode_e.swFACETED_HIDDEN_GREYED, true, true);

            DrawingModleDoc.Extension.SelectByID2(SwView5Front.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            int InsertCombine = 0;//插入内容是一个组合值
            InsertCombine = InsertCombine + (int)swInsertAnnotation_e.swInsertDimensions;//插入尺寸
            InsertCombine = InsertCombine + (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing;//插入为工程图标注的尺寸
            SwDraw.InsertModelAnnotations3((int)swImportModelItemsSource_e.swImportModelItemsFromSelectedComponent, InsertCombine, false, true, false, true);//


            BomTableAnnotation SwBomTableAnnotation = SwView5.InsertBomTable4(false, 829 / 1000.0, 312.5 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight, (int)swBomType_e.swBomType_TopLevelOnly, "默认", BomPath, false, (int)swNumberingType_e.swNumberingType_Flat, false);
            SwView5.SetKeepLinkedToBOM(true, SwBomTableAnnotation.BomFeature.Name);
            swApp.CloseDoc(SwComp.GetPathName());

            #region 自动拉件号
            SwDraw.ActivateView(SwView5.Name);
            DrawingModleDoc.Extension.SelectByID2(SwView5.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            SwAutoBalloonOptions = SwDraw.CreateAutoBalloonOptions();
            SwAutoBalloonOptions.Layout = 1;
            SwAutoBalloonOptions.ReverseDirection = false;
            SwAutoBalloonOptions.IgnoreMultiple = true;
            SwAutoBalloonOptions.InsertMagneticLine = true;
            SwAutoBalloonOptions.LeaderAttachmentToFaces = false;
            SwAutoBalloonOptions.Style = 1;
            SwAutoBalloonOptions.Size = 2;
            SwAutoBalloonOptions.EditBalloonOption = 1;
            SwAutoBalloonOptions.EditBalloons = true;
            SwAutoBalloonOptions.UpperTextContent = 1;
            SwAutoBalloonOptions.UpperText = "";
            SwAutoBalloonOptions.Layername = "Text";
            SwAutoBalloonOptions.ItemNumberStart = 1;
            SwAutoBalloonOptions.ItemNumberIncrement = 1;
            SwAutoBalloonOptions.ItemOrder = 0;
            SwDraw.AutoBalloon5(SwAutoBalloonOptions);
            #endregion
           
            ////////
            SwDraw.ActivateView(SwView5Front.Name);
            SheetPos = new double[2] { 829 / 1000.0, 294.5 / 1000.0 };
            SwMathPoint = SwMathUtility.CreatePoint(CalPartHeadPosInView(SheetPos, viewpos, SwView5Front.ScaleDecimal));
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, PartHeadBlock, false, 1, 0);


            SwComp = ((AssemblyDoc)AssemModleDoc).GetComponentByName("PlugSlotA-1");
            swApp.ActivateDoc3(SwComp.GetPathName(), true, 2, ref x);
            swApp.ActivateDoc3(DrawingModleDoc.GetPathName(), true, 2, ref x);
            viewpos = new double[] { 200 / 1000.0, 200 / 1000.0 };//定义SwView1插入点
            SolidWorks.Interop.sldworks.View SwView6 = SwDraw.CreateDrawViewFromModelView3(SwComp.GetPathName(), "*上视", viewpos[0], viewpos[1], 0);
            SwView6.ScaleRatio = new double[] { 4, 1 };

            SwView6.InsertBomTable4(false, 426.5 / 1000.0, 30 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight, (int)swBomType_e.swBomType_TopLevelOnly, "默认", BomPath, false, (int)swNumberingType_e.swNumberingType_Flat, false);
            swApp.CloseDoc(SwComp.GetPathName());
            ////////
            SwDraw.ActivateView(SwView6.Name);
            SheetPos = new double[2] { 426.5 / 1000.0, 12 / 1000.0 };
            SwMathPoint = SwMathUtility.CreatePoint(CalPartHeadPosInView(SheetPos, viewpos, SwView6.ScaleDecimal));
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, PartHeadBlock, false, 1, 0);

            DrawingModleDoc.Extension.SelectByID2(SwView6.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            InsertCombine = 0;//插入内容是一个组合值
            InsertCombine = InsertCombine + (int)swInsertAnnotation_e.swInsertDimensions;//插入尺寸
            InsertCombine = InsertCombine + (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing;//插入为工程图标注的尺寸
            SwDraw.InsertModelAnnotations3((int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel, InsertCombine, false, true, false, true);//

            SwComp = ((AssemblyDoc)AssemModleDoc).GetComponentByName("PlugSlotB-1");
            swApp.ActivateDoc3(SwComp.GetPathName(), true, 2, ref x);
            swApp.ActivateDoc3(DrawingModleDoc.GetPathName(), true, 2, ref x);
            viewpos = new double[] { 525 / 1000.0, 200 / 1000.0 };//定义SwView1插入点
            SolidWorks.Interop.sldworks.View SwView7 = SwDraw.CreateDrawViewFromModelView3(SwComp.GetPathName(), "*上视", viewpos[0], viewpos[1], 0);
            SwView7.ScaleRatio = new double[] { 4, 1 };//这里直接赋值了，可以通过box等计算视图比例

            SwView7.InsertBomTable4(false, 627.75 / 1000.0, 30 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight, (int)swBomType_e.swBomType_TopLevelOnly, "默认", BomPath, false, (int)swNumberingType_e.swNumberingType_Flat, false);
            swApp.CloseDoc(SwComp.GetPathName());
            ////////
            SwDraw.ActivateView(SwView7.Name);
            SheetPos = new double[2] { 627.75 / 1000.0, 12 / 1000.0 };
            SwMathPoint = SwMathUtility.CreatePoint(CalPartHeadPosInView(SheetPos, viewpos, SwView7.ScaleDecimal));
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, PartHeadBlock, false, 1, 0);

            DrawingModleDoc.Extension.SelectByID2(SwView7.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            InsertCombine = 0;//插入内容是一个组合值
            InsertCombine = InsertCombine + (int)swInsertAnnotation_e.swInsertDimensions;//插入尺寸
            InsertCombine = InsertCombine + (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing;//插入为工程图标注的尺寸
            SwDraw.InsertModelAnnotations3((int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel, InsertCombine, false, true, false, true);//

            SwComp = ((AssemblyDoc)AssemModleDoc).GetComponentByName("PlugSlotC-1");
            swApp.ActivateDoc3(SwComp.GetPathName(), true, 2, ref x);
            swApp.ActivateDoc3(DrawingModleDoc.GetPathName(), true, 2, ref x);
            viewpos = new double[] { 725 / 1000.0, 200 / 1000.0 };//定义SwView1插入点
            SolidWorks.Interop.sldworks.View SwView8 = SwDraw.CreateDrawViewFromModelView3(SwComp.GetPathName(), "*上视", viewpos[0], viewpos[1], 0);
            SwView8.ScaleRatio = new double[] { 4, 1 };

            SwView8.InsertBomTable4(false, 829 / 1000.0, 90 / 1000.0, (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight, (int)swBomType_e.swBomType_TopLevelOnly, "默认", BomPath, false, (int)swNumberingType_e.swNumberingType_Flat, false);
            swApp.CloseDoc(SwComp.GetPathName());
            ////////
            SwDraw.ActivateView(SwView8.Name);
            SheetPos = new double[2] { 829 / 1000.0, 72 / 1000.0 };
            SwMathPoint = SwMathUtility.CreatePoint(CalPartHeadPosInView(SheetPos, viewpos, SwView8.ScaleDecimal));
            SwSketchMrg.MakeSketchBlockFromFile(SwMathPoint, PartHeadBlock, false, 1, 0);

            DrawingModleDoc.Extension.SelectByID2(SwView8.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            InsertCombine = 0;//插入内容是一个组合值
            InsertCombine = InsertCombine + (int)swInsertAnnotation_e.swInsertDimensions;//插入尺寸
            InsertCombine = InsertCombine + (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing;//插入为工程图标注的尺寸
            SwDraw.InsertModelAnnotations3((int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel, InsertCombine, false, true, false, true);//

            swApp.ActivateDoc3(DrawingModleDoc.GetPathName(), true, 2, ref x);
            #endregion
        }
        public double[] CalPartHeadPosInView(double[] PosInSheet, double[] ViewPos, double ViewScale)
        {
            double[] Pos = new double[3];
            double[] PosRevToViewPos = new double[2] { PosInSheet[0] - ViewPos[0], PosInSheet[1] - ViewPos[1] };//插入点相对以视图位置为原点的坐标
            Pos[0] = PosRevToViewPos[0] / ViewScale;//视图比例
            Pos[1] = PosRevToViewPos[1] / ViewScale;//视图比例
            Pos[2] = 0;
            return Pos;
        }
        public DrawingComponent GetDrawingComp(DrawingComponent ViewRoot,string CompName, SolidWorks.Interop.sldworks.View SwView)
        {
            DrawingComponent Dc = null;
            Component2 SwRootComp = ViewRoot.Component;
            ModelDoc2 SwRootDoc = SwRootComp.GetModelDoc2();
            AssemblyDoc SwRootAssem = (AssemblyDoc)SwRootDoc;
            Component2 SwCompNeed = null;
            for (int i = 1; i < 21; i++)
            {
                SwCompNeed = SwRootAssem.GetComponentByName(CompName + "-" + i.ToString().Trim());
                if (SwCompNeed != null)
                {
                    break;
                }
            }
            if (SwCompNeed == null)
            {
                return null;
            }
            //
            Dc = SwCompNeed.GetDrawingComponent(SwView);
            return Dc;
        }
       

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            OutPutDrawing();
        }


    }

   


}
