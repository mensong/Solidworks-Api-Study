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

namespace SolidworksApiProject.Chapter8
{
    public partial class Chapter8Form : Form
    {
        SldWorks swApp = null;
        ModelDoc2 swAssemModleDoc = null;
        string ModleRoot = @"D:\正式版机械工业出版社出书\SOLIDWORKS API 二次开发实例详解\ModleAsbuit";
        public Chapter8Form()
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


        private void button1_Click(object sender, EventArgs e)
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");

            #region 打开PowerStrip.SLDASM。
            int IntError = -1;
            int IntWraning = -1;
            string filepath1 = ModleRoot + @"\第8章\RectanglePlug\PowerStrip.SLDASM";
            ModelDoc2 SwAssemDoc = swApp.OpenDoc6(filepath1, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);//打开部件
            #endregion

            #region 强制转化
            AssemblyDoc swAssem = null;
            if (SwAssemDoc.GetType() == 2)
            {
                swAssem = (AssemblyDoc)SwAssemDoc;
            }
            string assemname = SwAssemDoc.GetTitle().Substring(0, SwAssemDoc.GetTitle().LastIndexOf("."));//得到装配体名字
            #endregion

            #region 打开需要被插入的零件
            string parttoinsert = ModleRoot  +@"\第8章\RectanglePlug\PlugButton.SLDPRT";
            ModelDoc2 SwPartDoc = swApp.OpenDoc6(parttoinsert, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
            #endregion

            #region B.往装配体中插入零件PlugButton.SLDPRT。
            Component2 SwComp = swAssem.AddComponent5(parttoinsert,0,"",false,"",50,100,150);
            MessageBox.Show("插入零件成功!");
            swApp.CloseDoc(SwPartDoc.GetTitle());
            #endregion

            #region 选择操作
            SwAssemDoc.ClearSelection2(true);
            MessageBox.Show("已清除当前所有选择");
            SwAssemDoc.Extension.SelectByID2(SwComp.Name2 + "@" + assemname,"COMPONENT",0,0,0,false,0,null,0);
            #endregion

            #region C.将插入的零件固定。
            swAssem.FixComponent();
            #endregion

            #region 选中装配体中的部件PlugLED.SLDPRT
            SelectData SwData = ((SelectionMgr)SwAssemDoc.SelectionManager).CreateSelectData();//得到装配体的选择数据集
            SwComp = null;
            for (int i = 1; i < 20; i++)
            {
                SwComp = swAssem.GetComponentByName("PlugLED-" + i.ToString().Trim());
                if (SwComp != null)
                {
                    break;
                }
            }
            if (SwComp != null)
            {
                SwComp.Select4(false, SwData, false);
            }
            #endregion
 
            #region E.并打开选中的部件
            swAssem.OpenCompFile();
            #endregion
        }

        //在装配体中获取所有部件//装配树
        private void button2_Click(object sender, EventArgs e)
        {
            //这种方法是没有次序的，如果需要次序，需要使用特征去遍历，在特征中解
            string TopCompToFind = "PlugHead";
            string InnerCompToFind = "PlugPin2";
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            string filepath1 = ModleRoot + @"\第8章\RectanglePlug\5-8滑轮\装配体2.SLDASM";
            int IntError = -1;
            int IntWraning = -1;
            swAssemModleDoc = swApp.OpenDoc6(filepath1, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
          
            AssemblyDoc SwAssem = null;
            if (swAssemModleDoc.GetType() == 2)
            {
                SwAssem = (AssemblyDoc)swAssemModleDoc;
            }
            else
            {
                return;
            }

            #region 得到零件级部件--类似于明细表中的仅限顶层
            int n = SwAssem.GetComponentCount(false);//获得顶层部件的数量--包含阵列等部件
            MessageBox.Show("零件级部件数量:" + n.ToString().Trim());
            string compname = "";
            if (n > 0)
            {
                object[] ObjComps = SwAssem.GetComponents(false);//获得部件数组
                for (int i = 0; i < ObjComps.Length; i++)//循环得到每个部件
                {
                    Component2 swComp = (Component2)ObjComps[i]; //得到部件
                    compname = swComp.Name2; //记录部件名,获得的compname非顶层时候，含有路径
                    string compbase = compname;
                    if (compbase.Contains("/"))//名称存在路径的时候，截取获得最终所需的模型名
                    {
                        compbase = compbase.Substring(compbase.LastIndexOf("/") + 1, compbase.Length - compbase.LastIndexOf("/") - 1);
                    }

                    if (compbase.Substring(0, compbase.LastIndexOf("-")) == TopCompToFind)
                    {
                        compname = TopCompToFind + "搜索结果:\r\n部件名称:" + compname + "\r\n部件地址:" + swComp.GetPathName();
                    }
                    else if (compbase.Substring(0, compbase.LastIndexOf("-")) == InnerCompToFind)
                    {
                        compname = TopCompToFind + "搜索结果:\r\n部件名称:" + compname + "\r\n部件地址:" + swComp.GetPathName();
                    }
                    MessageBox.Show(compname);
                }
            }
            #endregion
            //遍历零件级别效率低，等不用就不用
            #region 得到顶层部件--类似于明细表中的仅限零件
            n = SwAssem.GetComponentCount(true);//获得顶层部件的数量--包含阵列等部件
            MessageBox.Show("顶层部件数量:" + n.ToString().Trim());
            #endregion

            //与GetComponentByName对比，已经名字推荐使用GetComponentByName，未知名称时候，若需要通过文件属性寻找，则使用GetComponents（但结果为无序）
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string CompToDo = "PlugPin1";
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            //string filepath1 = ModleRoot + @"\第8章\RectanglePlug\PowerStrip.SLDASM";
            string filepath1 = ModleRoot + @"\第8章\RectanglePlug\5-8滑轮\装配体2.SLDASM";
            int IntError = -1;
            int IntWraning = -1;
            swAssemModleDoc = swApp.OpenDoc6(filepath1, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
          
            AssemblyDoc SwAssem = null;
            if (swAssemModleDoc.GetType() == 2)
            {
                SwAssem = (AssemblyDoc)swAssemModleDoc;
            }
            else
            {
                return;
            }

            SelectData Seldate = ((SelectionMgr)swAssemModleDoc.SelectionManager).CreateSelectData();
            Component2 SwComp = null;
            for (int i = 1; i < 20; i++)
            {
                SwComp = SwAssem.GetComponentByName(CompToDo + "-" + i.ToString().Trim());
                if (SwComp != null)
                {
                    break;
                }
            }
       
            if (SwComp != null)
            {
                GetCompData(SwComp, CompToDo);
                #region 设置部件信息与状态
                SwComp.Select4(false, Seldate, false);
                swAssemModleDoc.HideComponent2();
                SwComp.ExcludeFromBOM = true;
                SwComp.SetSuppression2(1);//将部件设置为轻化       
                #endregion
                GetCompData(SwComp, CompToDo);//重新获得部件信息
            }
        }
        public void GetCompData(Component2 SwComp, string CompToDo)//获取部件信息
        {
            StringBuilder sb = new StringBuilder("部件状态:\r\n");//对话框显示信息
            bool isVirtual = SwComp.IsVirtual;
            sb.Append("是否是虚件:" + isVirtual.ToString() + "\r\n");
            string ParentCompname = ((Component2)SwComp.GetParent()).Name2;
            sb.Append("父部件名:" + ParentCompname + "\r\n");
            string SelectByIDString = SwComp.GetSelectByIDString();//没有零件中的元素时候，即没有A元素
            sb.Append("SelectByID2参数Name:" + SelectByIDString + "\r\n");
            bool NoBom = SwComp.ExcludeFromBOM;//是否排除在明细栏
            sb.Append("排除在明细栏:" + NoBom.ToString() + "\r\n");
            bool isfixed = SwComp.IsFixed();
            sb.Append("是否固定:" + isfixed.ToString() + "\r\n");
            bool isMirrored = SwComp.IsMirrored();//是否是镜像出来的实例
            sb.Append("是否为镜像实例:" + isMirrored.ToString() + "\r\n");
            bool isPatternInstance = SwComp.IsPatternInstance();//是否是阵列出来的实例
            sb.Append("是否为阵列实例:" + isPatternInstance.ToString() + "\r\n");
            bool isRoot = SwComp.IsRoot();//是否是顶层部件（即装配体特征树中顶层部件）
            sb.Append("是否为顶层部件:" + isRoot.ToString() + "\r\n");
            int CompState = SwComp.GetSuppression();//获得部件的状态
            sb.Append("零件状态:" + CompState.ToString() + "\r\n");
            MessageBox.Show(sb.ToString(), CompToDo + "部件信息");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string OldSub = "PlugLED";
            string NewSubPath = ModleRoot + @"\第8章\RectanglePlug\PlugLEDForReplace.SLDPRT";
            //string NewSubPath = ModleRoot + @"\第8章\RectanglePlug\5-8滑轮\装配体2.SLDASM";
            
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            AssemblyDoc SwAssem = null;
            if (swAssemModleDoc.GetType() == 2)
            {
                SwAssem = (AssemblyDoc)swAssemModleDoc;
            }
            else
            {
                return;
            }

            SelectData Seldate = ((SelectionMgr)swAssemModleDoc.SelectionManager).CreateSelectData();
            Component2 SwComp = null;
            for (int i = 1; i < 20; i++)
            {
                SwComp = SwAssem.GetComponentByName(OldSub + "-" + i.ToString().Trim());
                if (SwComp != null)
                {
                    break;
                }
            }

            if (SwComp != null)
            {
                SwComp.Select4(false, Seldate, false);
                SwAssem.ReplaceComponents(NewSubPath, "默认", true, true);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {     
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            AssemblyDoc SwAssem = null;
            if (swAssemModleDoc == null)
                MessageBox.Show("请打开装配对象");
                return;
            if (swAssemModleDoc.GetType() == 2)
            {
                SwAssem = (AssemblyDoc)swAssemModleDoc;//转换
            }
            else
            {
                return;
            }

            #region 打开需要被插入的零件
            string parttoinsert = ModleRoot + @"\第8章\RectanglePlug\PlugButton.SLDPRT";
            int IntError = -1;
            int IntWraning = -1;
            ModelDoc2 SwPartDoc = swApp.OpenDoc6(parttoinsert, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
            #endregion

            #region 往装配体中插入零件PlugButton.SLDPRT。
            Component2 SwComp = SwAssem.AddComponent5(parttoinsert, 0, "", false, "", 0, 0, 0);
            swApp.CloseDoc(SwPartDoc.GetTitle());
            #endregion


            #region 装配一对基准
            string MateBaseName1 = "CenterV@PlugButton-1@PowerStrip";
            string MateBaseName2 = "RectangleWireConnectFace@PlugBottomBox-1@PowerStrip";
            swAssemModleDoc.Extension.SelectByID2(MateBaseName1, "PLANE", 0, 0, 0, false, 1, null, 0);
            swAssemModleDoc.Extension.SelectByID2(MateBaseName2, "PLANE", 0, 0, 0, true, 1, null, 0);
            int x = -1;
            Mate2 SwMate = SwAssem.AddMate5(5, 1, true, 0.02, 0.02, 0.02, 0, 0, 0, 0, 0, false, true, 0, out x);
            Feature swMateFeature = (Feature)SwMate;//配合也是特征，将其转化为特征对象
            if (swMateFeature != null)
            {
                swMateFeature.Name = "测试距离配合";//对配合特征进行重命名，以后将来获取方便。
            }
            #endregion

            #region 装配第二对基准
            MateBaseName1 = "CenterH@PlugButton-1@PowerStrip";
            MateBaseName2 = "BoxCenterH@PlugBottomBox-1@PowerStrip";
            swAssemModleDoc.Extension.SelectByID2(MateBaseName1, "PLANE", 0, 0, 0, false, 1, null, 0);
            swAssemModleDoc.Extension.SelectByID2(MateBaseName2, "PLANE", 0, 0, 0, true, 1, null, 0);
            x = -1;
            SwMate = SwAssem.AddMate5(0, 1, true, 0.02, 0.02, 0.02, 0, 0, 0, 0, 0, false, true,0, out x);
            swMateFeature = (Feature)SwMate;//配合也是特征，将其转化为特征对象
            if (swMateFeature != null)
            {
                swMateFeature.Name = "测试重合配合";//对配合特征进行重命名，以后将来获取方便。
            }
            #endregion

            #region 装配第三对基准
            MateBaseName1 = "ButtonTop@PlugButton-1@PowerStrip";
            MateBaseName2 = "BoxInnerTop@PlugTopBox-1@PowerStrip";
            swAssemModleDoc.Extension.SelectByID2(MateBaseName1, "PLANE", 0, 0, 0, false, 1, null, 0);
            swAssemModleDoc.Extension.SelectByID2(MateBaseName2, "PLANE", 0, 0, 0, true, 1, null, 0);
            x = -1;
            SwMate = SwAssem.AddMate5(5, 0, true, 0.006, 0.006, 0.006, 0, 0, 0, 0, 0, false, true, 0, out x);
            swMateFeature = (Feature)SwMate;//配合也是特征，将其转化为特征对象
            if (swMateFeature != null)
            {
                swMateFeature.Name = "测试重合2配合";//对配合特征进行重命名，以后将来获取方便。
            }
            #endregion
          
            swAssemModleDoc.EditRebuild3();//需要刷新
        }

        private void butt(object sender,EventArgs e)
        {
            string compname = "PlugSlotA";
            open_swfile(" ",getProcesson("SLDWORKS"),"SldWorks..Application");
            AssemblyDoc SwAssem = null;
            if (swAssemModleDoc.GetType() == 2)
            {
                SwAssem = (AssemblyDoc)swAssemModleDoc;
            }
            else
            {
                return;
            }
            string[] Configs = swAssemModleDoc.GetConfigurationNames();//装配体集合
            Component2 SwComp = null;
            for (int i = 1;i < 20;i++)
            {
                SwComp = SwAssem.GetComponentByName(compname+"-"+i.ToString().Trim());
                if (SwComp != null)
                {
                    break;
                }
            }
            if (SwComp != null)
            {
                object[] ObjComMates = SwComp.GetMates();
                if (ObjComMates != null)
                {

                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string Compname = "PlugSlotA";//指定寻找配合的部件
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");

            string filepath1 = ModleRoot + @"\第8章\RectanglePlug\5-8滑轮\装配体2.SLDASM";
            int IntError = -1;
            int IntWraning = -1;
            //swAssemModleDoc = swApp.OpenDoc6(filepath1, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);

            AssemblyDoc SwAssem = null;
            if (swAssemModleDoc.GetType() == 2)
            {
                SwAssem = (AssemblyDoc)swAssemModleDoc;
            }
            else
            {
                return;
            }
            string[] Configs = swAssemModleDoc.GetConfigurationNames();//获得装配体配置名称集合

            Component2 SwComp = null;
            for (int i = 1; i < 20; i++)
            {
                SwComp = SwAssem.GetComponentByName(Compname + "-" + i.ToString().Trim());//获得部件
                if (SwComp != null)
                {
                    break;
                }
            }

            if (SwComp != null)
            {
                object[] ObjCompMates = SwComp.GetMates();//获得部件中的配合对象集合
                if (ObjCompMates != null)
                {
                    StringBuilder sb = new StringBuilder("部件" + SwComp.Name2 + "存在如下配合关系:\r\n");
                    foreach (object ObjCompMate in ObjCompMates)
                    {
                        if (ObjCompMate is Mate2)//判断集合元素是否为Mate2对象
                        {
                            Mate2 swmate = (Mate2)ObjCompMate;//强制转化为Mate2对象
                            sb.Append("配合" + ((Feature)swmate).Name + ":");
                            if (swmate.Type == 6)//是否为角度配合
                            {
                                sb.Append("类型为角度;");
                                DisplayDimension disdim = swmate.DisplayDimension2[0];//默认取第一个参数，第二个参数仅齿轮情况下才会用到
                                Dimension swDim = disdim.GetDimension2(0);//双尺寸情况下，取第一个尺寸，见Remark
                                double[] dimcollect = swDim.GetValue3((int)swInConfigurationOpts_e.swSpecifyConfiguration, "配置2");//指定配置
                                sb.Append("指定配置[" + "配置2" + "]尺寸为" + dimcollect[0].ToString() + ";");
                            }
                            else if (swmate.Type == 5)//是否为距离配合
                            {
                                sb.Append("类型为距离;");
                                DisplayDimension disdim = swmate.DisplayDimension2[0];//默认取缔一个参数，第二个参数仅齿轮情况下才会用到
                                Dimension swDim = disdim.GetDimension2(0);
                                double[] dimcollect = swDim.GetValue3((int)swInConfigurationOpts_e.swAllConfiguration, "");//所有配置
                                for (int i = 0; i < Configs.Length; i++)
                                {
                                    sb.Append("配置[" + Configs[i] + "]尺寸为" + dimcollect[i].ToString() + ";");
                                }
                            }
                            else if (swmate.Type == 0)//是否为重合配合
                            {
                                sb.Append("类型为重合");
                            }
                            sb.Append("\r\n");
                        }
                    }
                    MessageBox.Show(sb.ToString());
                }
            }
        }
    }
}
