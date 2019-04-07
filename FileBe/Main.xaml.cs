using Corel.Interop.VGCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FileBe
{
    public class Size
    {
        public double x { get; set; }
        public double y { get; set; }
        public Size()
        {
            x = 0; y = 0;
        }
        public Size(double _x, double _y)
        {
            x = _x;
            y = _y;
        }
    }
    public partial class Main : UserControl
    {
        Corel.Interop.CorelDRAW.Application app = null;
        public Main()
        {
            InitializeComponent();
        }
        public Main(object _app)
        {
            InitializeComponent();
            app = (Corel.Interop.CorelDRAW.Application)_app;
            //app.Application.SelectionChange += new DIVGApplicationEvents_SelectionChangeEventHandler(SelectionChanged);
        }

        private void SelectionChanged()
        {
            //MessageBox.Show("a");
        }

        private void btnReset_Click(object sender, RoutedEventArgs e)
        {
            numRow.Value = 1;
            numCol.Value = 1;
            numSpace.Value = 1;
            chkUnSpace.IsChecked = false;
            lblHei.Content = "0";
            lblWid.Content = "0";
            lblTotalSize.Content = "0";
            lblTotal.Content = "0";
            numInsert.Value = 1;
        }
        private void calSize()
        {
            if (app?.ActiveDocument == null) return;
            app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
            if (app.ActiveSelectionRange.Count < 1) return;
            try
            {
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                double wid = ((s.x + spaceCal()) * (numCol.Value ?? 1)) - spaceCal();
                double hei = ((s.y + spaceCal()) * (numRow.Value ?? 1)) - spaceCal();
                lblWid.Content = string.Format("{0:#,##0.###}", wid);
                lblHei.Content = string.Format("{0:#,##0.###}", hei);
                lblTotalSize.Content = string.Format("{0:#,##0.###}", ((wid * hei) / 10000));
                lblTotal.Content = (numCol.Value ?? 1) * (numRow.Value ?? 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source);
            }
        }
        private Size getTotalSize()
        {
            Size s = new Size();
            try
            {
                s.x = app.ActiveSelection.SizeWidth * numCol.Value ?? 1;
                s.y = app.ActiveSelection.SizeHeight * numRow.Value ?? 1;
                return s;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
                return null;
            }
        }
        private float spaceCal()
        {
            if (chkUnSpace.IsChecked.Value) return 0;
            return (float)(numSpace.Value ?? 0) / 10;
        }

        private void cal_KeyUp(object sender, KeyEventArgs e)
        {
            calSize();
        }

        private void chkUnSpace_Checked(object sender, RoutedEventArgs e)
        {
            calSize();
        }

        private void btnReCal_Click(object sender, RoutedEventArgs e)
        {
            calSize();
        }

        private void btnCreate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                ShapeRange orSh = app.ActiveSelectionRange;
                double space = 0;
                for (int i = 1; i < numRow.Value; i++)
                {
                    space += s.y + spaceCal();
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(0, -space));
                }
                orSh.CreateSelection();
                space = 0;
                for (int j = 1; j < numCol.Value; j++)
                {
                    space += s.x + spaceCal();
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(space, 0));
                }
                orSh.Group();
                app.ActiveLayer.CreateRectangle(orSh.LeftX, orSh.TopY, orSh.RightX, orSh.BottomY).CreateSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        private void btnCrLine_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                ShapeRange orSh = null;
                Size position = new Size(app.ActiveSelectionRange.PositionX, app.ActiveSelectionRange.PositionY);
                Size size = getTotalSize();
                app.ActiveSelection.Delete();
                app.ActiveLayer.CreateLineSegment(position.x,position.y, position.x, position.y - size.y).CreateSelection();
                orSh = app.ActiveSelectionRange;
                double space = 0;
                for (int i = 0; i < (numCol.Value ?? 1); i++)
                {
                    space += s.x;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(space, 0));
                }
                space = 0;
                app.ActiveLayer.CreateLineSegment(position.x, position.y, position.x + size.x, position.y).CreateSelection();
                orSh.AddRange(app.ActiveSelectionRange);
                for (int j = 0; j < (numRow.Value ?? 1); j++)
                {
                    space += s.y;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(0, -space));
                }
                orSh.Group();
                app.ActiveLayer.CreateRectangle(orSh.LeftX, orSh.TopY, orSh.RightX, orSh.BottomY).CreateSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        private void cal_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            calSize();
        }

        private void btnCrSemiLine_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                ShapeRange orSh = app.ActiveSelectionRange;
                Corel.Interop.VGCore.Shape sh = app.ActiveLayer.CreateRectangle(orSh.LeftX, orSh.TopY, orSh.RightX, orSh.BottomY);
                sh.ConvertToCurves();
                Curve c = app.ActiveDocument.CreateCurve();
                SubPath ss = c.CreateSubPath(orSh.RightX, orSh.BottomY);
                ss.AppendCurveSegment(orSh.LeftX, orSh.BottomY);
                ss.AppendCurveSegment(orSh.LeftX, orSh.TopY);
                sh.Curve.CopyAssign(c);
                orSh.Delete();
                sh.CreateSelection();
                orSh = app.ActiveSelectionRange;
                double space = 0;
                for (int i = 1; i < numRow.Value; i++)
                {
                    space += s.y;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(0, -space));
                }
                orSh.CreateSelection();
                space = 0;
                for (int j = 1; j < numCol.Value; j++)
                {
                    space += s.x;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(space, 0));
                }
                orSh.Add(app.ActiveLayer.CreateLineSegment(orSh.PositionX, orSh.PositionY, orSh.PositionX + orSh.SizeWidth, orSh.PositionY));
                orSh.Add(app.ActiveLayer.CreateLineSegment(orSh.PositionX + orSh.SizeWidth, orSh.PositionY, orSh.PositionX + orSh.SizeWidth, orSh.PositionY - orSh.SizeHeight ));
                orSh.Group();
                app.ActiveLayer.CreateRectangle(orSh.LeftX, orSh.TopY, orSh.RightX, orSh.BottomY).CreateSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        private void btnDelImg_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ShapeRange orSh = app.ActiveSelectionRange;
                if (orSh.Count == 1)
                    orSh.UngroupAll();
                foreach(Corel.Interop.VGCore.Shape sh in orSh)
                {
                    if (sh.Type == cdrShapeType.cdrBitmapShape)
                        sh.Delete();
                    //MessageBox.Show(sh.Type.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        private void btnCalSizeInsert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                numInsert.Value = calSizeInsert();
                numSpace.Value = 3;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }
        private int calSizeInsert()
        {
            try
            {
                if (app?.ActiveSelectionRange == null) return 0;
                app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
                ShapeRange orSh = app.ActiveSelectionRange;
                if (orSh.Count < 1) return 0;
                double sRec = orSh.SizeHeight * orSh.SizeWidth;
                double sEll = Math.Pow(orSh.SizeWidth / 2, 2) * Math.PI;
                double d = Math.Sqrt(((sRec - sEll) / 1.5775796) / Math.PI) * 2;
                int round = (int)Math.Round(d * 10, 0);
                return round;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
                return 0;
            }
        }
        private void numInsertChange()
        {
            try
            {
                int size = calSizeInsert();
                if (size < 1) return;
                if (numInsert.Value == 1) return;
                if (numInsert.Value > size)
                {
                    int round = (int)(((numInsert.Value - size) ?? 0) / 1.2);
                    numSpace.Value = round + 3;
                }
                else numSpace.Value = 3;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }
        private void numInsert_KeyUp(object sender, KeyEventArgs e)
        {
            numInsertChange();
        }
        private void numInsert_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            numInsertChange();
        }

        private void btnCreInsert_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
