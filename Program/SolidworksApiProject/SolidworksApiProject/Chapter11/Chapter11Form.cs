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
using System.Runtime.InteropServices;

namespace SolidworksApiProject.Chapter11
{
    public partial class Chapter11Form : Form
    {
        SldWorks swApp = null;
        ModelDoc2 SwModleDoc = null;
        string ModleRoot = @"D:\正式版机械工业出版社出书\SOLIDWORKS API 二次开发实例详解\ModleAsbuit";
        public Chapter11Form()
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
            GetBomInOrder();//11.2 实例分析：按特征树顺序提取零件信息
        }

        public void GetBomInOrder()
        {
            string OutputFilePath = Application.StartupPath + @"\FeatureOutput.txt";
            Dictionary<string, BomBean> BomDic = new Dictionary<string, BomBean>();//  Dictionary<件号, 零件信息>
            Dictionary<string, int> PartCountDic = new Dictionary<string, int>();//Dictionary<文件名, 数量>()
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");//判断进程是否有SolidWorks进程
            TraverseFeature(SwModleDoc, BomDic, PartCountDic);//遍历模型零件    
            #region 将阵列，镜像的内容统计到数量中
            foreach (string PartNo in BomDic.Keys)
            {
                if (PartCountDic.ContainsKey(BomDic[PartNo].FileNameNoExt))
                {
                    BomDic[PartNo].Count = BomDic[PartNo].Count + PartCountDic[BomDic[PartNo].FileNameNoExt];
                }
            }
            #endregion
            #region 输出到外部txt文本
            StreamWriter sw = File.AppendText(OutputFilePath);//新建文件
            foreach (string partno in BomDic.Keys)
            {
                StringBuilder sb = new StringBuilder("件号:"+partno+",");//件号：序号（循环次数）
                sb.Append(BomDic[partno].DwgNo+",");//名称序号
                sb.Append(BomDic[partno].PartName + ",");//部件名称
                sb.Append(BomDic[partno].Material + ",");//材料名称
                sb.Append(BomDic[partno].PreMass + ",");//质量值
                sb.Append(BomDic[partno].Count.ToString() + ",");
                sw.WriteLine(sb.ToString());
            }
            sw.WriteLine("******结束******");
            sw.Flush();
            sw.Close();
            Process.Start(OutputFilePath);
            #endregion
        }
        public void TraverseFeature(ModelDoc2 Modle, Dictionary<string, BomBean> BomDic, Dictionary<string, int> PartCountDic)//顶层部件
        {
            FeatureManager FM = Modle.FeatureManager;
            Feature swFeat = Modle.FirstFeature();
            CustomPropertyManager CPM;
            string FileNameNoExt = "";
            int PartNoInNowLevetl = 1;//当前层级下的件号
            while (swFeat != null)
            {
                object ObjComp = swFeat.GetSpecificFeature2();
                if (ObjComp is Component2)
                {
                    Component2 Comp = (Component2)ObjComp;
                    if (Comp.ExcludeFromBOM == true || Comp.GetSuppression() == 0)//压缩或排除在明细表外的部件
                    {
                        swFeat = swFeat.GetNextFeature();
                        continue;
                    }
                    else if (Comp.GetSuppression() == 1 || Comp.GetSuppression() == 3 || Comp.GetSuppression() == 4)//轻化
                    {
                        Comp.SetSuppression2(2);//全部还原
                    }
                    ModelDoc2 SubModle = Comp.GetModelDoc2();
                    CPM = SubModle.Extension.CustomPropertyManager[""];
                    FileNameNoExt = SubModle.GetTitle().Substring(0, SubModle.GetTitle().LastIndexOf("."));
                    if (Comp.IsPatternInstance())//是镜像阵列之类的
                    {
                        if (PartCountDic.ContainsKey(FileNameNoExt))
                        {
                            PartCountDic[FileNameNoExt] = PartCountDic[FileNameNoExt] + 1;
                        }
                        else
                        {
                            PartCountDic.Add(FileNameNoExt, 1);
                        }
                        swFeat = swFeat.GetNextFeature();
                        continue;
                    }
                    #region 登记零件信息
                    BomBean bb = new BomBean();
                    bb.DwgNo = GetBomInfo(CPM, "图号");
                    bb.PartName = GetBomInfo(CPM, "名称");
                    bb.Material = GetBomInfo(CPM, "材料");
                    bb.PreMass = double.Parse(GetBomInfo(CPM, "单重"));
                    bb.Remark = GetBomInfo(CPM, "备注");
                    bb.FileNameNoExt = FileNameNoExt;
                    BomDic.Add(PartNoInNowLevetl.ToString().Trim(), bb);
                    #endregion
                    if (SubModle.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)//说明是装配体//开分进程扫描子装配
                    {
                        TraverseComp(Comp, BomDic, PartNoInNowLevetl.ToString().Trim(), PartCountDic, CPM);
                    }
                    PartNoInNowLevetl = PartNoInNowLevetl + 1;
                }
                swFeat = swFeat.GetNextFeature();
            }
        }
        public void TraverseComp(Component2 Comp, Dictionary<string, BomBean> BomDic, string ParentPartNo, Dictionary<string, int> PartCountDic, CustomPropertyManager CPM)// 这个是扫描子部件,TopPbb是相对的顶层,StdDwg=true,是标准图
        {
            Feature swFeat = Comp.FirstFeature();
            int PartNoInNowLevetl = 1;//当前层级下的件号
            string FileNameNoExt = "";
            while (swFeat != null)
            {
                object ObjComp = swFeat.GetSpecificFeature2();
                if (ObjComp is Component2)
                {
                    Component2 Comp1 = (Component2)ObjComp;
                    if (Comp1.ExcludeFromBOM == true || Comp1.GetSuppression() == 0)//压缩或排除在明细表外的部件
                    {
                        swFeat = swFeat.GetNextFeature();
                        continue;
                    }
                    else if (Comp1.GetSuppression() == 1 || Comp1.GetSuppression() == 3 || Comp1.GetSuppression() == 4)//轻化
                    {
                        Comp1.SetSuppression2(2);//全部还原
                    }
                    ModelDoc2 SubModle = Comp1.GetModelDoc2();
                    CPM = SubModle.Extension.CustomPropertyManager[""];
                    FileNameNoExt = SubModle.GetTitle().Substring(0, SubModle.GetTitle().LastIndexOf("."));
                    if (Comp.IsPatternInstance())//是镜像阵列之类的
                    {
                        if (PartCountDic.ContainsKey(FileNameNoExt))
                        {
                            PartCountDic[FileNameNoExt] = PartCountDic[FileNameNoExt] + 1;
                        }
                        else
                        {
                            PartCountDic.Add(FileNameNoExt, 1);
                        }
                        swFeat = swFeat.GetNextFeature();
                        continue;
                    }
                    #region 登记零件信息
                    BomBean bb = new BomBean();
                    bb.DwgNo = GetBomInfo(CPM, "图号");
                    bb.PartName = GetBomInfo(CPM, "名称");
                    bb.Material = GetBomInfo(CPM, "材料");
                    bb.PreMass = double.Parse(GetBomInfo(CPM, "单重"));
                    bb.Remark = GetBomInfo(CPM, "备注");
                    bb.FileNameNoExt = FileNameNoExt;
                    BomDic.Add(ParentPartNo + "-" + PartNoInNowLevetl.ToString().Trim(), bb);
                    #endregion
                    if (SubModle.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)//说明是装配体//开分进程扫描子装配
                    {
                        TraverseComp(Comp1, BomDic, ParentPartNo + "-" + PartNoInNowLevetl.ToString().Trim(), PartCountDic, CPM);
                    }
                    PartNoInNowLevetl = PartNoInNowLevetl + 1;
                }
                swFeat = swFeat.GetNextFeature();
            }
        }
        public string GetBomInfo(CustomPropertyManager CPM, string FieldName)
        {
            string DisplayValue = "";
            string Value = "";
            bool x = false;
            CPM.Get5(FieldName, false, out Value, out DisplayValue, out x);
            return DisplayValue;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DoLinearPattern();//11.4 实例分析：线性阵列特征数据获取与修改
        }

        public void DoLinearPattern()
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            Feature SwFeat = ((AssemblyDoc)SwModleDoc).FeatureByName("SlotAs");
            if (SwFeat.GetTypeName2() == "LocalLPattern")//线型阵列
            {
                StringBuilder sb = new StringBuilder();
                LocalLinearPatternFeatureData SwLinearPattern = SwFeat.GetDefinition();//得到线型阵列特征数据对象
                sb.Append("阵列总是为(含源与跳过实例):" + SwLinearPattern.D1TotalInstances.ToString().Trim() + "\r\n");//包含了源与跳过的实例
                sb.Append("阵列间距:" + SwLinearPattern.D1Spacing + "\r\n");

                #region 获得阵列的源数据
                sb.Append("********************阵列源信息********************\r\n");
                object[] ObjSeedFeatures = SwLinearPattern.SeedComponentArray;//获得了阵列属性管理器中 要阵列的零部件 列表中的内容
                foreach (object ObjSeedFeature in ObjSeedFeatures)
                {
                    Feature SwSeedFeature = (Feature)ObjSeedFeature;
                    sb.Append("源特征名为:" + SwSeedFeature.Name + "\r\n");
                    object ObjComp = SwSeedFeature.GetSpecificFeature2();
                    if (ObjComp is Component2)
                    {
                        Component2 SwComp = (Component2)ObjComp;
                        sb.Append("源为部件，部件名为:" + SwComp.Name2 + "\r\n");
                    }
                }
                #endregion
                sb.Append("********************阵列跳过实例信息********************\r\n");
                #region 获得跳过的阵列实例
                sb.Append("跳过的实例为阵列方向(不含源)的第");
                int[] SkipIndexs = SwLinearPattern.SkippedItemArray;
                foreach (int Index in SkipIndexs)
                {
                    sb.Append(Index.ToString().Trim() + ",");
                }
                #endregion
                sb.Append("个\r\n");
                MessageBox.Show(sb.ToString(), "特征信息获取结果");

                #region 添加阵列源
                if (SwLinearPattern.AccessSelections(SwModleDoc, null))
                {
                    Array.Resize(ref ObjSeedFeatures, (ObjSeedFeatures.GetUpperBound(0) + 2));
                }
                ObjSeedFeatures[1] = ((AssemblyDoc)SwModleDoc).GetComponentByName("PlugSlotB-1");
                DispatchWrapper[] dispWrapArr = null;
                dispWrapArr = (DispatchWrapper[])ObjectArrayToDispatchWrapperArray(ObjSeedFeatures);
                SwLinearPattern.SeedComponentArray = dispWrapArr;
                SwFeat.ModifyDefinition(SwLinearPattern, SwModleDoc, null);
                SwLinearPattern.ReleaseSelectionAccess();
                #endregion
            }
        }

        public DispatchWrapper[] ObjectArrayToDispatchWrapperArray(object[] Objects)
        {
            int ArraySize = 0;
            ArraySize = Objects.GetUpperBound(0);
            DispatchWrapper[] d = new DispatchWrapper[ArraySize + 1];
            int ArrayIndex = 0;
            for (ArrayIndex = 0; ArrayIndex <= ArraySize; ArrayIndex++)
            {
                d[ArrayIndex] = new DispatchWrapper(Objects[ArrayIndex]);
            }
            return d;
        }
    }

    public class BomBean
    {
        public string DwgNo { get; set; }
        public string PartName { get; set; }
        public string Material { get; set; }
        public int Count { get; set; }
        public double PreMass { get; set; }//单个零件的重量
        public string Remark { get; set; }
        public string FileNameNoExt { get; set; }//文件名，用于统计数量
        public BomBean()
        {
            DwgNo = "";
            PartName = "";
            Material = "";
            Count = 1;
            PreMass = 0;
            Remark = "";
            FileNameNoExt = "";
        }
    }
}
