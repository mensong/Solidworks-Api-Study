using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace Macro11
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
            myModelView.RotateAboutCenter(0, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0068722339297276882, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0034361169648638441, 0);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.054977871437821506, 0.10727389548843222);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.065286222332413033, 0.17623568544528148);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.024052818754046908, 0.061299368850532686);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.065286222332413033, 0.19922294876423124);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.024052818754046908, 0.076624211063165859);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.089339041086459944, 0.2758471598273971);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.068722339297276877, 0.19156052765791465);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.020616701789183067, 0.04597452663789952);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.044669520543229972, 0.091949053275799039);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.013744467859455376, 0.030649684425266343);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0068722339297276882, 0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.013744467859455376, 0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0034361169648638441, 0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.017180584824319219, 0);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.010308350894591534, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.017180584824319219, -0.04597452663789952);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.013744467859455376, -0.030649684425266343);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.0034361169648638441, 0);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.0034361169648638441, -0.015324842212633171);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.0034361169648638441, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.0068722339297276882, -0.02298726331894976);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.0068722339297276882, -0.015324842212633171);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.0034361169648638441, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.017180584824319219, -0.03831210553158293);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.017180584824319219, -0.03831210553158293);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.0068722339297276882, -0.0076624211063165857);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(0.034361169648638439, -0.030649684425266343);
            // 
            // Roll View
            ModelView swModelView = null;
            swModelView = ((ModelView)(swDoc.ActiveView));
            swModelView.RollBy(0);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.013744467859455376, 0.061299368850532686);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0068722339297276882, 0.030649684425266343);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.048105637508093817, 0.19156052765791465);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.024052818754046908, 0.068961789956849276);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0859029241215961, 0.23753505429581417);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.061850105367549188, 0.14558600102001512);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.017180584824319219, 0.04597452663789952);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.061850105367549188, 0.14558600102001512);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.017180584824319219, 0.05363694774421611);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.082466807156732269, 0.18389810655159808);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.054977871437821506, 0.20688536987054781);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.024052818754046908, 0.12259873770106537);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0034361169648638441, 0.02298726331894976);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.017180584824319219, 0.068961789956849276);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0034361169648638441, 0.015324842212633171);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.020616701789183067, 0.061299368850532686);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0068722339297276882, 0.0076624211063165857);
            // 
            // Roll View
            swModelView = ((ModelView)(swDoc.ActiveView));
            swModelView.RollBy(0);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.2817615911188352, 0.3831210553158293);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.034361169648638439, 0.04597452663789952);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.12026409377023455, 0.16857326433896491);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.0859029241215961, 0.068961789956849276);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.024052818754046908, 0.015324842212633171);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.024052818754046908, 0.015324842212633171);
            myModelView = ((ModelView)(swDoc.ActiveView));
            myModelView.RotateAboutCenter(-0.037797286613502283, 0.030649684425266343);
            // 
            // Roll View
            swModelView = ((ModelView)(swDoc.ActiveView));
            swModelView.RollBy(0);
            boolstatus = swDoc.Extension.SelectByRay(0.042449999999999988, 0, 0, -0.56159996523349964, 0.788641309378255, 0.25030054772589627, 0.0027638754256671129, 3, false, 0, 0);


            return;
        }

        // The SldWorks swApp variable is pre-assigned for you.
        public SldWorks swApp;

    }
}

