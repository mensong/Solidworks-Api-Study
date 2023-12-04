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


namespace SolidworksApiProject.Chapter14
{
    public partial class Chapter14Form : Form
    {
        SldWorks swApp = null;
        ModelDoc2 SwModleDoc = null;
        public Chapter14Form()
        {
            InitializeComponent();

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
            DoAttribute();//14.4 实例分析：属性的添加与访问
        }
        public void DoAttribute()
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            SelectionMgr SwSelMrg = SwModleDoc.SelectionManager;
            SelectData SwSeldt = SwSelMrg.CreateSelectData();

            AttributeDef SwAttributeDef = swApp.DefineAttribute("ManufacturingRecord");
            #region 向数据包定义中添加参数
            SwAttributeDef.AddParameter("EntityName", (int)swParamType_e.swParamTypeString, 1.0, 0);//新建实体名称参数
            SwAttributeDef.AddParameter("EntityFinish", (int)swParamType_e.swParamTypeDouble, 1.0, 0);//新建实体是否完成参数
            SwAttributeDef.AddParameter("EntityRequire", (int)swParamType_e.swParamTypeString, 1.0, 0);//对实体的要求
            #endregion
            SwAttributeDef.Register();//注册数据包

            #region 给边线添加数据包
            SwModleDoc.Extension.SelectByID2("", "EDGE", 0.05, 0, 0.035, false, 0, null, 0);//坐标选中
            Edge SwEdge = SwSelMrg.GetSelectedObject6(1, -1);
            SolidWorks.Interop.sldworks.Attribute swAttribute = SwAttributeDef.CreateInstance5(SwModleDoc, SwEdge, "EdgeRecord1", 0, (int)swInConfigurationOpts_e.swAllConfiguration);//给边线添加数据包
            #region 修改边线数据包中的参数
            Parameter swParameter = (Parameter)swAttribute.GetParameter("EntityName");
            swParameter.SetStringValue2("边线A1", (int)swInConfigurationOpts_e.swAllConfiguration, "");
            swParameter = (Parameter)swAttribute.GetParameter("EntityFinish");
            swParameter.SetDoubleValue2(-1, (int)swInConfigurationOpts_e.swAllConfiguration, "");
            swParameter = (Parameter)swAttribute.GetParameter("EntityRequire");
            swParameter.SetStringValue2("边线需要圆滑过渡", (int)swInConfigurationOpts_e.swAllConfiguration, "");
            #endregion
            #endregion

            #region 给面添加数据包
            SwModleDoc.Extension.SelectByID2("", "FACE", 0.05, 2 / 1000.0, 0.025, false, 0, null, 0);
            Face2 SwFace = SwSelMrg.GetSelectedObject6(1, -1);
            swAttribute = SwAttributeDef.CreateInstance5(SwModleDoc, SwFace, "FaceRecord1", 0, (int)swInConfigurationOpts_e.swAllConfiguration);//给面添加数据包
            #region 修改面数据包中的参数
            swParameter = (Parameter)swAttribute.GetParameter("EntityName");
            swParameter.SetStringValue2("面C1", (int)swInConfigurationOpts_e.swAllConfiguration, "");
            swParameter = (Parameter)swAttribute.GetParameter("EntityFinish");
            swParameter.SetDoubleValue2(1, (int)swInConfigurationOpts_e.swAllConfiguration, "");
            swParameter = (Parameter)swAttribute.GetParameter("EntityRequire");
            swParameter.SetStringValue2("表面粗糙度需要满足相关要求", (int)swInConfigurationOpts_e.swAllConfiguration, "");
            #endregion
            #endregion

            SwModleDoc.ClearSelection2(true);//清空所有选择
            MessageBox.Show("数据包添加完毕，并已清空所有选择");

            #region 获得添加的数据包信息
            StringBuilder sb = new StringBuilder("属性数据包信息：\r\n");
            GetAttributeInfo("EdgeRecord1", sb, SwSeldt);//获得边线数据包信息
            sb.Append("\r\n");
            GetAttributeInfo("FaceRecord1", sb, SwSeldt);//获得面数据包信息
            MessageBox.Show(sb.ToString(), "数据包信息获得并选中相应元素!");
            #endregion
        }

        public void GetAttributeInfo(string AttributeName, StringBuilder sb, SelectData SwSeldt)
        {
            Feature SwFeat = ((PartDoc)SwModleDoc).FeatureByName(AttributeName);
            if (SwFeat.GetTypeName2() == "Attribute")
            {
                sb.Append(AttributeName+"数据包:\r\n");
                SolidWorks.Interop.sldworks.Attribute swAttribute = SwFeat.GetSpecificFeature2();
                Parameter swParameter = (Parameter)swAttribute.GetParameter("EntityName");
                sb.Append(swParameter.GetName() + "=" + swParameter.GetStringValue().Trim() + "\r\n");
                swParameter = (Parameter)swAttribute.GetParameter("EntityFinish");
                sb.Append(swParameter.GetName() + "=" + swParameter.GetDoubleValue().ToString().Trim() + "\r\n");
                swParameter = (Parameter)swAttribute.GetParameter("EntityRequire");
                sb.Append(swParameter.GetName() + "=" + swParameter.GetStringValue().Trim() + "\r\n");

                Entity SwEntity = swAttribute.GetEntity();//得到数据包附加的元素
                SwEntity.Select4(true,SwSeldt);//选中该对象
            }
        }
    }
}
