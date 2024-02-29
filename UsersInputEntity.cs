using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace Dll类库_输出多段线坐标
{
    /// <summary>
    /// 得到用户输入的PolyLine和Pint3D
    /// </summary>
    public class UsersInputEntity
    {
        public Polyline Polyline = new Polyline();
        public Point3d BasePonint = new Point3d();
        public bool isSelected = false;

        public UsersInputEntity(Polyline _Polyline, Point3d _BasePonint, bool _isSelected)
        {
            Polyline = _Polyline;
            BasePonint = _BasePonint;
            isSelected = _isSelected;
        }

        public UsersInputEntity()
        {

        }

    }
}
