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

namespace SolidworksApiProject.Chapter6
{
    public partial class Chapter6Form : Form
    {
        SldWorks swApp = null;
        ModelDoc2 swAssemModleDoc = null;
        string ModleRoot = @"D:\正式版机械工业出版社出书\SOLIDWORKS API 二次开发实例详解\ModleAsbuit";
        public Chapter6Form()
        {
            InitializeComponent();

            ModleRoot = Path.GetDirectoryName(this.GetType().Assembly.Location) + @"\..\..\..\..\..\ModleAsbuit";
            ModleRoot = Path.GetFullPath(ModleRoot);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region 1先后打开两个文档
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            int IntError = -1;
            int IntWraning = -1;
            string filepath1 = ModleRoot + @"\RectanglePlug\PowerStrip.SLDASM";
            string filepath2 = ModleRoot + @"\RectanglePlug\PowerStrip.SLDDRW";
            ModelDoc2 SwAssemDoc= swApp.OpenDoc6(filepath1, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
            ModelDoc2 SwDrawDoc = swApp.OpenDoc6(filepath2, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
            //SwAssemDoc.Visible = true;
            //SwDrawDoc.Visible = true;
            #endregion
            #region 2通过工程图视图得到引用文档
            object[] SwViews = ((DrawingDoc)SwDrawDoc).Sheet["图纸1(2)"].GetViews();
            foreach (object ob in SwViews)
            {
                SolidWorks.Interop.sldworks.View SwView = (SolidWorks.Interop.sldworks.View)ob;
                if (SwView.Name == "工程图视图3")
                {
                    ModelDoc2 ModleForView = SwView.ReferencedDocument;
                    MessageBox.Show("工程图视图3的参考模型为:" + ModleForView.GetPathName());
                }
            }
            #endregion
            #region 3
            swApp.ActivateDoc3(filepath1, true, 2, IntError);
            MessageBox.Show("重新激活PowerStrip.SLDASM装配体成功!");
            #endregion
            #region 4得到装配体中部件对象
            Component2 Comp = ((AssemblyDoc)SwAssemDoc).GetComponentByName("PlugTopBox-1");
            ModelDoc2 CompToDoc = Comp.GetModelDoc2();
            MessageBox.Show("通过装配体中的部件获得:" + CompToDoc.GetPathName());
            #endregion
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
                swAssemModleDoc = (ModelDoc2)swApp.ActiveDoc;
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

        private void button2_Click(object sender, EventArgs e)
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            int IntError = -1;
            int IntWraning = -1;
            string filepath1 = ModleRoot + @"\RectanglePlug\SlotA\InnerPlugeA.SLDPRT";
            ModelDoc2 SwPartDoc = swApp.OpenDoc6(filepath1, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);

            #region 配置管理器
            SwPartDoc.AddConfiguration3("文档添加配置示例", "", "", (int)swConfigurationOptions2_e.swConfigOption_DontActivate);

            ConfigurationManager CfgMrg = SwPartDoc.ConfigurationManager;
            CfgMrg.AddConfiguration("配置管理器添加配置示例", "", "", (int)swConfigurationOptions2_e.swConfigOption_DontActivate, "默认", "");

            Configuration cfgnow = CfgMrg.ActiveConfiguration;
            MessageBox.Show("当前被激活的配置为:"+cfgnow.Name);

            SwPartDoc.DeleteConfiguration2("文档添加配置示例");
            MessageBox.Show("文档添加配置示例 被删除");
            #endregion

            #region 特征管理器
            FeatureManager FeatMrg = SwPartDoc.FeatureManager;
            object[] FeatObjs = FeatMrg.GetFeatures(true);
            foreach (object feobj in FeatObjs)
            {
                Feature SwFeat = (Feature)feobj;
                if (SwFeat!=null)
                {
                    MessageBox.Show(SwFeat.Name);
                    if (SwFeat.Name == "FFiveBaseTP1")
                    {
                        SwFeat.Select2(false, 0);
                    }
                }
            }
            #endregion

            #region 选择管理器
            SelectionMgr SelMrg = SwPartDoc.SelectionManager;
            MessageBox.Show("当前被选择的数量:"+SelMrg.GetSelectedObjectCount2(0).ToString());
            SwPartDoc.ClearSelection2(true);
            MessageBox.Show("清除选择集后,当前被选择的数量:" + SelMrg.GetSelectedObjectCount2(0).ToString());
            #endregion
        }

        private void button3_Click(object sender, EventArgs e)
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            int IntError = -1;
            int IntWraning = -1;
            string filepath1 = ModleRoot + @"\RectanglePlug\PlugButton.SLDPRT";
            ModelDoc2 SwPartDoc = swApp.OpenDoc6(filepath1, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);

            double L = SwPartDoc.Parameter("D2@SRectangle").SystemValue*1000;//长
            double K = SwPartDoc.Parameter("D1@SRectangle").SystemValue * 1000;//宽
            double H = SwPartDoc.Parameter("D1@FRectangle").SystemValue * 1000;//高
            double R = SwPartDoc.Parameter("D1@FRRectangle").SystemValue * 1000;//倒角
            MessageBox.Show("长=" + L.ToString() + "宽=" + K.ToString() + "高=" + H.ToString() + "导圆=" + R.ToString());

            SwPartDoc.Parameter("D2@SRectangle").SystemValue = 40 / 1000.0;//修改为长40
            SwPartDoc.Parameter("D1@SRectangle").SystemValue = 33 / 1000.0;//修改为宽33
            SwPartDoc.Parameter("D1@FRRectangle").SystemValue = 5 / 1000.0;//修改为导圆R5
            SwPartDoc.EditRebuild3();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            #region A. 打开PlugTopBox.SLDPRT零件。
            int IntError = -1;
            int IntWraning = -1;
            string filepath1 = ModleRoot + @"\RectanglePlug\PlugTopBox.SLDPRT";
            ModelDoc2 SwPartDoc = swApp.OpenDoc6(filepath1, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
            #endregion

            ModelDocExtension SwDocExten = SwPartDoc.Extension;
            string[] ConfigNames = SwPartDoc.GetConfigurationNames();
            bool x = false;

            #region B. 获得自定义标签中的已有属性“单重”，“类型”的属性值。
            CustomPropertyManager CuspMrg = SwDocExten.CustomPropertyManager[""];//自定义标签的属性管理器
            string MassValue = "";
            string ResolvedMass = "";
            CuspMrg.Get5("单重", true, out MassValue, out ResolvedMass, out x);
            string TypeValue = "";
            string ResolvedType = "";
            CuspMrg.Get5("类型", true, out TypeValue, out ResolvedType, out x);
            MessageBox.Show("单重=" + ResolvedMass + "\r\n" + "类型=" + ResolvedType, "自定义标签属性");
            #endregion
            
            #region C. 获得配置待定标签中，方壳配置下已有属性“名称”，“材料”的属性值。
            CuspMrg = SwDocExten.CustomPropertyManager["方壳"];//配置待定标签中方壳配置的属性管理器
            string NameValue = "";
            string ResolvedName = "";
            CuspMrg.Get5("名称", true, out NameValue, out ResolvedName, out x);
            string MaterialValue = "";
            string ResolvedMaterial = "";
            CuspMrg.Get5("材料", true, out MaterialValue, out ResolvedMaterial, out x);
            MessageBox.Show("名称=" + ResolvedName + "\r\n" + "材料=" + ResolvedMaterial, "配置待定标签中方壳配置");
            #endregion

            #region D. 在自定义标签中，添加新属性名“连接方式”，值为“螺钉”。
            CuspMrg = SwDocExten.CustomPropertyManager[""];
            CuspMrg.Add3("连接方式", 30, "螺钉", 2);
            #endregion

            //添加属性信息(自定义标签)
            CuspMrg = SwDocExten.CustomPropertyManager[""];
            CuspMrg.Add3("易烊千玺",30,"跳街舞最帅",2);

            #region E. 在自定义标签中，更新属性 “类型”，值修改为“TopBox”
            CuspMrg.Add3("类型", 30, "TopBox", 2);
            #endregion

            #region F. 在配置待定标签中，为圆壳配置添加属性“名称”，“材料”，其中“名称”属性的值带有尺寸联动。
            CuspMrg = SwDocExten.CustomPropertyManager["圆壳"];
            CuspMrg.Add3("名称", 30, "圆壳 \"D1@SketchCircle\"X\"D2@CircleBox\"", 2);
            CuspMrg.Add3("材料", 30, "ABC", 2);
            #endregion

            //添加属性--配置待定标签
            CuspMrg = SwDocExten.CustomPropertyManager["千玺"];
        }


        //文档的设置
        private void button5_Click(object sender, EventArgs e)
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
         
            #region 调用模板--新建文档
            ModelDoc2 SwPartDoc = swApp.NewDocument(@"C:\ProgramData\SolidWorks\SolidWorks 2012\templates\gb_part.prtdot", 0,0,0);
            MessageBox.Show("新建文档成功，当前文档名为:" + SwPartDoc.GetTitle());
            #endregion

            ModelDocExtension SwDocExten = SwPartDoc.Extension;//获得文档扩展对象

            #region 设置绘图标准
            SwDocExten.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDetailingDimensionStandard, 0, (int)swDetailingStandard_e.swDetailingStandardGB);
            #endregion

            #region 设置出详图显示过滤器
            SwDocExten.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayFeatureDimensions, 0, true);
            SwDocExten.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayReferenceDimensions, 0, true);
            SwDocExten.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayDimXpertDimensions, 0, true);
            #endregion

            #region 设置质量单位为公斤
            SwDocExten.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsMassPropMass, 0, (int)swUnitsMassPropMass_e.swUnitsMassPropMass_Kilograms);
            #endregion

            MessageBox.Show("文档设置完成!");
        }

    }
}
