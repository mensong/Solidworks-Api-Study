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


namespace SolidworksApiProject.Chapter13
{
    public partial class Chapter13Form : Form
    {
        SldWorks swApp = null;
        ModelDoc2 SwModleDoc = null;
        public Chapter13Form()
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
            DoEquationMgr();//13.2 实例分析：方程式的增删改
        }

        public void DoEquationMgr()
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");
            EquationMgr SwEquationMgr = SwModleDoc.GetEquationMgr();

            StringBuilder sb = new StringBuilder("方程式存储列表：\r\n");
            GetAllEquationDetail(SwEquationMgr, sb);////获得当前所有存在的方程式

            #region 修改方程
            string EquationStr = "\"D2@RectangleBox\"=2*\"H\"";
            int EqIndex = AddOrReviseEquation(SwEquationMgr, EquationStr, SwModleDoc);
            sb = new StringBuilder("");
            if (EqIndex >= 0)
            {
                sb.Append("修改的方程索引为:" + EqIndex.ToString().Trim() + "\r\n");
            }
            SwEquationMgr.EvaluateAll();
            SwModleDoc.EditRebuild3();
            sb.Append("\r\n");
            sb.Append("方程式存储列表：\r\n");
            GetAllEquationDetail(SwEquationMgr, sb);//获得当前模型所有存在的方程式
            #endregion

            #region 删除方程
            string DimName = "D2@RectangleBox";
            List<int> Indexs = RemoveEquation(SwEquationMgr, DimName);//移除指定名称的方程
            if (Indexs.Count > 0)
            {
                sb = new StringBuilder("被删除的索引为:\r\n");
                foreach (int aa in Indexs)
                {
                    sb.Append(aa.ToString() + ",");
                }
                SwEquationMgr.EvaluateAll();
                sb.Append("\r\n");
                sb.Append("方程式存储列表：\r\n");
                GetAllEquationDetail(SwEquationMgr, sb);
            }
            SwModleDoc.EditRebuild3();
            #endregion

            #region 添加方程
            EquationStr = "\"D2@RectangleBox\"=\"H\"";
            EqIndex = AddOrReviseEquation(SwEquationMgr, EquationStr, SwModleDoc);
            sb = new StringBuilder("");
            if (EqIndex>=0)
            {
                sb.Append("添加的方程索引为:" + EqIndex.ToString().Trim() + "\r\n");
            }
            SwEquationMgr.EvaluateAll();
            SwModleDoc.EditRebuild3();
            sb.Append("\r\n");
            sb.Append("方程式存储列表：\r\n");
            GetAllEquationDetail(SwEquationMgr, sb);
            #endregion
        }

        public void GetAllEquationDetail(EquationMgr SwEquationMgr,StringBuilder sb)//获得所有存在的方程式
        {          
            for (int i = 0; i < SwEquationMgr.GetCount(); i++)
            {
                sb.Append("索引号" + i.ToString().Trim() + "-->表的式:" + SwEquationMgr.Equation[i] + "-->数值:" + SwEquationMgr.Value[i].ToString() + "\r\n");
            }
            MessageBox.Show(sb.ToString().Trim(), "方程式存储情况");
        }

        public List<int> RemoveEquation(EquationMgr SwEquationMgr,string EquationLeftName)//移除指定名称的方程
        {
            List<int> RemoveIndex = new List<int>() ;
            string Left = "";//记录等号左边部分
            string Right = "";//记录等号右边部分
            for (int i = 0; i < SwEquationMgr.GetCount(); i++)//遍历全部方程式删除指定名称方程式
            {
                Left = SwEquationMgr.Equation[i].Substring(0, SwEquationMgr.Equation[i].IndexOf("="));
                Right = SwEquationMgr.Equation[i].Substring(SwEquationMgr.Equation[i].IndexOf("=") + 1, SwEquationMgr.Equation[i].Length - SwEquationMgr.Equation[i].IndexOf("=") - 1);
                if (EquationLeftName == Left.Substring(1, Left.Length - 2))//将左边头尾双引号去掉与需要的名称对比
                {
                    SwEquationMgr.Delete(i);//删除
                    RemoveIndex.Add(i);//记录删除的索引号
                    continue;
                }
            }
            return RemoveIndex;
        }

        public int AddOrReviseEquation(EquationMgr SwEquationMgr, string EquationStr, ModelDoc2 Doc)//添加方程式
        {
            int EqIndex = -1;
            bool NewEquation = true;//判断是否新建
            string Left = "";//记录等号左边部分
            for (int i = 0; i < SwEquationMgr.GetCount(); i++)
            {
                Left = SwEquationMgr.Equation[i].Substring(0, SwEquationMgr.Equation[i].IndexOf("="));
                if (Left == EquationStr.Substring(0, EquationStr.IndexOf("=")))//说明存在相同方程，则为修改
                {
                    NewEquation = false;
                    EqIndex = i;//记录索引号
                    break;
                }
            }
            if (NewEquation)//新建
            {
                EqIndex = SwEquationMgr.Add3(EqIndex, EquationStr, true, (int)swInConfigurationOpts_e.swAllConfiguration, Doc.GetConfigurationNames());
            }
            else//修改
            {
                SwEquationMgr.Equation[EqIndex] = EquationStr;
            }
            return EqIndex;
        }
    }
}
