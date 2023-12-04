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

namespace SolidworksApiProject.Chapter12
{
    public partial class Chapter12Form : Form
    {
        SldWorks swApp = null;
        ModelDoc2 SwModleDoc = null;
        string ModleRoot = @"D:\正式版机械工业出版社出书\SOLIDWORKS API 二次开发实例详解\ModleAsbuit";
        public Chapter12Form()
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
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            SelectionMgr SwSelMrg = SwModleDoc.SelectionManager;
            MessageBox.Show(((swSelectType_e)SwSelMrg.GetSelectedObjectType3(-1, 0)).ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DoEntity();//12.2.2 实例分析：实体的设置与获得
        }

        public void DoEntity()
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            SelectionMgr SwSelMrg = SwModleDoc.SelectionManager;
            SelectData SwSelData = SwSelMrg.CreateSelectData();
            #region 选择边线或面进行重命名
            SwModleDoc.Extension.SelectByID2("", "EDGE", 0.05, 0, 0.035, false, 0, null, 0);//坐标选中
            Edge SwEdge = SwSelMrg.GetSelectedObject6(1, -1);//通过选择管理器方法转化为实体对象
            MessageBox.Show("选中的边线实体名字:" + SwModleDoc.GetEntityName(SwEdge));//获得实体名称
            ((PartDoc)SwModleDoc).SetEntityName(SwEdge, "自定义边线A2");//设置实体名称
            MessageBox.Show("选中的边线实体被重命名为:" + SwModleDoc.GetEntityName(SwEdge));//反馈实体名称
            SwModleDoc.Extension.SelectByID2("", "FACE", 0.05, 2 / 1000.0, 0.025, false, 0, null, 0);
            Face2 SwFace = SwSelMrg.GetSelectedObject6(1, -1);
            MessageBox.Show("选中的面实体名字:" + SwModleDoc.GetEntityName(SwFace));
            ((PartDoc)SwModleDoc).SetEntityName(SwFace, "顶壳外表面");
            MessageBox.Show("选中的面实体被重命名为:" + SwModleDoc.GetEntityName(SwFace));
            #endregion

            #region 获得所有被命名的实体
            SwModleDoc.ClearSelection2(true);
            object[] ObjNamedEntities = ((PartDoc)SwModleDoc).GetNamedEntities();//获得零件中所有有命名的实体
            StringBuilder sb = new StringBuilder("所有被命名的实体数量为:" + ObjNamedEntities.Length.ToString() + "\r\n");
            foreach (object ObjNamedEntity in ObjNamedEntities)
            {
                Entity SwEntity = (Entity)ObjNamedEntity;
                sb.Append("实体名字：" + SwModleDoc.GetEntityName(SwEntity) + ",类型为:" + ((swSelectType_e)SwEntity.GetType()).ToString() + "\r\n");
            }
            MessageBox.Show(sb.ToString());//反馈所有有命名的实体信息
            #endregion

            #region 通过指定名称得到实体对象，并选中，再得到细分对象面，边线
            SwModleDoc.ClearSelection2(true);
            Entity SwEntityToSelect = ((PartDoc)SwModleDoc).GetEntityByName("自定义边线A2", (int)swSelectType_e.swSelEDGES);//根据指定名称得到边线对象
            SwEntityToSelect.Select4(false, SwSelData);//将边线选中
            Edge SwEdgeToSelect = SwSelMrg.GetSelectedObject6(1, -1);//通过选择管理器方法转化为细分的边线对象
            if (SwEdgeToSelect != null)
            {
                MessageBox.Show("成功通过名称\"" + "自定义边线A2" + "\"获得边线对象，并选中");//反馈选中及获得的对象结果
            }
            /////////////
            SwModleDoc.ClearSelection2(true);
            SwEntityToSelect = ((PartDoc)SwModleDoc).GetEntityByName("顶壳外表面", (int)swSelectType_e.swSelFACES);
            SwEntityToSelect.Select4(false, SwSelData);
            Face2 SwFaceToSelect = SwSelMrg.GetSelectedObject6(1, -1);
            if (SwFaceToSelect != null)
            {
                MessageBox.Show("成功通过名称\"" + "顶壳外表面" + "\"获得面对象，并选中");
            }
            #endregion       
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DoPersistentIDs();//12.2.3 实例分析：对象永久ID获取与使用
        }

        public void DoPersistentIDs()
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            SelectionMgr SwSelMrg = SwModleDoc.SelectionManager;
            SelectData SwSelData = SwSelMrg.CreateSelectData();

            #region 获得对象ID
            int SelectCount = SwSelMrg.GetSelectedObjectCount2(-1);//获得所有选择内容
            byte[] FaceID = null;
            byte[] EdgeID = null;
            byte[] FeatureID = null;
            Face2 SwFace = null;
            Edge SwEdge = null;
            Feature SwFeature = null;
            for (int i = 1; i <= SelectCount; i++)//循环选择集中的每个元素
            {
                string SelType = ((swSelectType_e)SwSelMrg.GetSelectedObjectType3(i, -1)).ToString().Trim();
                if (SelType == "swSelFACES")//面对象
                {
                    SwFace = SwSelMrg.GetSelectedObject6(i, -1);
                    FaceID = SwModleDoc.Extension.GetPersistReference3(SwFace);
                }
                else if (SelType == "swSelEDGES")//边线对象
                {
                    SwEdge = SwSelMrg.GetSelectedObject6(i, -1);
                    EdgeID = SwModleDoc.Extension.GetPersistReference3(SwEdge);
                }
                else if (SelType == "swSelBODYFEATURES")//特征对象
                {
                    SwFeature = SwSelMrg.GetSelectedObject6(i, -1);
                    FeatureID = SwModleDoc.Extension.GetPersistReference3(SwFeature);
                }
            }
            #endregion

            SwModleDoc.ClearSelection2(true);
            SwFace = null;
            SwEdge = null;
            SwFeature = null;
            MessageBox.Show("已清空当前所有选择");
            int a = 0;

            #region 通过ID获得对象
            SwFace = SwModleDoc.Extension.GetObjectByPersistReference3(FaceID, out a);
            ((Entity)SwFace).Select4(false, SwSelData);
            SwEdge = SwModleDoc.Extension.GetObjectByPersistReference3(EdgeID, out a);
            ((Entity)SwEdge).Select4(true, SwSelData);
            SwFeature = SwModleDoc.Extension.GetObjectByPersistReference3(FeatureID, out a);
            SwFeature.Select2(true, 0);
            #endregion   
        }

        public static byte[] intToBytes2(int n)
        {
            byte[] b = new byte[4];

            for (int i = 0; i < 4; i++)
            {
                b[i] = (byte)(n >> (24 - i * 8));

            }
            return b;
        }

        public static int byteToInt2(byte[] b)
        {
            int mask = 0xff;
            int temp = 0;
            int n = 0;
            for (int i = 0; i < b.Length; i++)
            {
                n <<= 8;
                temp = b[i] & mask;
                n |= temp;
            }
            return n;
        }
    }
}
