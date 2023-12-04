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

namespace SolidworksApiProject.Chapter7
{
    public partial class Chapter7Form : Form
    {
        SldWorks swApp = null;
        ModelDoc2 swAssemModleDoc = null;
        string ModleRoot = @"D:\正式版机械工业出版社出书\SOLIDWORKS API 二次开发实例详解\ModleAsbuit";
        public Chapter7Form()
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

            #region 打开PlugTopBox.SLDPRT零件。
            int IntError = -1;
            int IntWraning = -1;
            string filepath1 = ModleRoot + @"\RectanglePlug\PlugTopBox.SLDPRT";
            ModelDoc2 SwPartDoc = swApp.OpenDoc6(filepath1, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
            #endregion

            #region 强制转化
            PartDoc swPart = null;
            if (SwPartDoc.GetType() == 1)
            {
               swPart = (PartDoc)SwPartDoc;
            }
            #endregion

            #region A.得到零件的材料
            string MtDateBaseName="";
            string mt = swPart.GetMaterialPropertyName2("", out MtDateBaseName);
            MessageBox.Show("零件材料为：" + mt + "-->所在材料数据库为:" + MtDateBaseName);
            #endregion

            #region B.赋予零件的材料
            swPart.SetMaterialPropertyName2("", MtDateBaseName, "PVC 僵硬");
            #endregion

            #region C.选中特征
            Feature swFeat1 = swPart.FeatureByName("BoxInnerTop");
            swFeat1.Select2(false, 0);//清空之前
            Feature swFeat2 = swPart.FeatureByName("RectangleR");
            swFeat2.Select2(true,0);//保留之前选择
            #endregion
        }
    }
}
