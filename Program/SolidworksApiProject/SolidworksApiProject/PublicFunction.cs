using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidworksApiProject
{
    class PublicFunction
    {
        public void open_swfile(string filepath, int x, string pgid, ModelDoc2 swAssemModleDoc)
        {
            //SldWorks swApp;//启动程序接口//放到公共变量去
            if (x == 0)//无进程-->新建
            {
                System.Windows.Forms.MessageBox.Show("当前无启动中的Solidworks应用");
            }
            else if (x == 1)//有进程,得到进程的应用对象
            {
                System.Type swtype = System.Type.GetTypeFromProgID(pgid);
                MainForm .swApp = (SldWorks)System.Activator.CreateInstance(swtype);
                swAssemModleDoc = (ModelDoc2)MainForm.swApp.ActiveDoc;
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
    }
}
