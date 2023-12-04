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

namespace SolidworksApiProject.Chapter5
{
    public partial class Chapter5Form : Form
    {
        SldWorks swApp = null;
        ModelDoc2 swAssemModleDoc = null;
        string ModleRoot = @"D:\正式版机械工业出版社出书\SOLIDWORKS API 二次开发实例详解\ModleAsbuit";
        public Chapter5Form()
        {
            InitializeComponent();

            ModleRoot = Path.GetDirectoryName(this.GetType().Assembly.Location) + @"\..\..\..\..\..\ModleAsbuit";
            ModleRoot = Path.GetFullPath(ModleRoot);
        }

        private void btn_newapp_Click(object sender, EventArgs e)
        {
            KILLSW();
            swApp = new SldWorks();
            swApp.Visible = true;
        }

        private void btn_getapp_Click(object sender, EventArgs e)
        {
            open_swfile("", getProcesson("SLDWORKS"), "SldWorks.Application");//
            if (swAssemModleDoc == null)//没有文件打开
            {
                MessageBox.Show("当前无打开的Solidworks文件");
            }
            else//有文件打开
            {
                MessageBox.Show("当前激活的Solidworks文件为:\r\n" + swAssemModleDoc.GetTitle());
            }
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
                //swApp.Visible = true;
      
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
            int IntError = -1;
            int IntWraning = -1;
            string filepath1 = ModleRoot + @"\RectanglePlug\PlugTopBox.SLDPRT";
            string filepath2 = ModleRoot + @"\RectanglePlug\PlugWire.SLDPRT";
            swApp.OpenDoc6(filepath1, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
            swApp.OpenDoc6(filepath2, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref IntError, ref IntWraning);
            //swApp.ActivateDoc3(Application.StartupPath + @"\Modle\RectanglePlug\PlugTopBox.SLDPRT", true, 2, IntError);
            object[] ObjModles = swApp.GetDocuments();
            int i = 1;
            foreach (object objmodle in ObjModles)
            {
                ModelDoc2 mc = (ModelDoc2)objmodle;
                //mc.Visible = true;//使得文档能被看到，有时文档不显示但已经打开了
                MessageBox.Show("SW进程打开的所有文档-"+i.ToString()+":"+mc.GetTitle());
                i = i + 1;
            }
            MessageBox.Show("当前激活的文档为:"+((ModelDoc2)swApp.ActiveDoc).GetTitle());
            swApp.ActivateDoc3(filepath1, true, 2, IntError);
            MessageBox.Show("文档:" + ((ModelDoc2)swApp.ActiveDoc).GetTitle()+"被激活!");


            #region 系统设置
            swApp.SetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsBOMTemplates, Application.StartupPath + @"\Modle");
            #endregion
            swApp.CloseDoc(filepath1);

        }
    }
}
