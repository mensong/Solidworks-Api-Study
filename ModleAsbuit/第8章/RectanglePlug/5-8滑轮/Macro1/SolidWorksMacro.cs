using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace Macro1
{
    public partial class SolidWorksMacro
    {
        public void Main()
        {
            
            ModelDoc2 swDoc = null;
            PartDoc swPart = null;
            DrawingDoc swDrawing = null;
            AssemblyDoc swAssembly = null;
            bool boolstatus = false;
            int longstatus = 0;
            int longwarnings = 0;
            swDoc = ((ModelDoc2)(swApp.ActiveDoc));
            ModelView myModelView = null;
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0, 0.015324842212633171);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.017180584824319219, 0.061299368850532686);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.017180584824319219, 0.091949053275799039);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0068722339297276882, 0.02298726331894976);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.027488935718910753, 0.16857326433896491);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0068722339297276882, 0.04597452663789952);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.024052818754046908, 0.13792357991369855);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.024052818754046908, 0.14558600102001512);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.020616701789183067, 0.12259873770106537);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0068722339297276882, 0.03831210553158293);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.024052818754046908, 0.16857326433896491);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.017180584824319219, 0.091949053275799039);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0068722339297276882, 0.03831210553158293);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.027488935718910753, 0.15324842212633172);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.013744467859455376, 0.061299368850532686);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0068722339297276882, 0.015324842212633171);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0034361169648638441, 0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0034361169648638441, 0);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0034361169648638441, 0);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0034361169648638441, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0034361169648638441, -0.015324842212633171);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0034361169648638441, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0, -0.0076624211063165857);
            // 
            // Roll View
            ModelView swModelView = null;
            swModelView = ((ModelView)(swDoc.ActiveView));
            swModelView.RollBy(0);
            // 
            // Pan
            swModelView = ((ModelView)(swDoc.ActiveView));
            swModelView.TranslateBy(0.05215702671312427, 0.012286991869918698);
            boolstatus = swDoc.Extension.SelectByRay(-0.062384321143667876, 0.001999999999981128, -0.00971342200770664, 0.94997471370288211, -0.29572291604661705, -0.1004788547407514, 0.0011107395454230619, 2, false, 0, 0);
            swDoc.FeatEdit();
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("L@SketchRec@PlugTopBox.SLDPRT", "DIMENSION", 0.0018732682004611728, 0.0025962605992594923, -0.056320403898723947, false, 0, null, 0);
            Dimension myDimension = null;
            myDimension = ((Dimension)(swDoc.Parameter("L@SketchRec")));
            myDimension.SystemValue = 0.1;
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);


            return;
        }

        // The SldWorks swApp variable is pre-assigned for you.
        public SldWorks swApp;

    }
}

