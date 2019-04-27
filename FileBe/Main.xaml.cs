using Corel.Interop.VGCore;
using NumericUpDownLib;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FileBe
{
    public partial class Main : UserControl
    {
        Corel.Interop.CorelDRAW.Application app = null;
        string err = "Chưa chọn đối tượng!";
        public Main()
        {
            InitializeComponent();
        }
        public Main(object _app)
        {
            InitializeComponent();
            app = (Corel.Interop.CorelDRAW.Application)_app;
            readTheme();
            //app.Application.SelectionChange += new DIVGApplicationEvents_SelectionChangeEventHandler(SelectionChanged);
        }

        private void SelectionChanged()
        {
            //MessageBox.Show("a");
        }
        private void ChangeTheme(string theme)
        {
            string uri = "/FlatTheme;component/ColorStyle/" + theme + ".xaml";
            ResourceDictionary resourceDict = System.Windows.Application.LoadComponent(new Uri(uri, UriKind.Relative)) as ResourceDictionary;
            Resources.MergedDictionaries[0].Clear();
            Resources.MergedDictionaries[0].MergedDictionaries.Add(resourceDict);
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
            numInsert.Value = 0;
        }
        private void calSize()
        {
            if (app?.ActiveDocument == null) return;
            if (app.ActiveSelectionRange.Count < 1) return;
            app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
            try
            {
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                float space = spaceCal();
                double wid = ((s.x + space) * numCol.Value) - space;
                double hei = ((s.y + space) * numRow.Value) - space;
                lblWid.Content = string.Format("{0:#,##0.###}", wid);
                lblHei.Content = string.Format("{0:#,##0.###}", hei);
                lblTotalSize.Content = string.Format("{0:#,##0.###}", ((wid * hei) / 10000));
                lblTotal.Content = numCol.Value * numRow.Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source,"Lỗi");
            }
        }
        private Size getTotalSize()
        {
            Size s = new Size();
            try
            {
                s.x = app.ActiveSelection.SizeWidth * numCol.Value;
                s.y = app.ActiveSelection.SizeHeight * numRow.Value;
                return s;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
                return null;
            }
        }
        private float spaceCal()
        {
            if (chkUnSpace.IsChecked.Value) return 0;
            return (float)numSpace.Value / 10;
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
            if (app?.ActiveDocument == null) return;
            if (app.ActiveSelectionRange.Count < 1)
            {
                MessageBox.Show(err, "Lỗi");
                return;
            }
            try
            {
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                ShapeRange orSh = app.ActiveSelectionRange;
                double space = 0;
                float sp = spaceCal();
                for (int i = 1; i < numRow.Value; i++)
                {
                    space += s.y + sp;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(0, -space));
                }
                orSh.CreateSelection();
                space = 0;
                for (int j = 1; j < numCol.Value; j++)
                {
                    space += s.x + sp;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(space, 0));
                }
                orSh.Group();
                app.ActiveLayer.CreateRectangle(orSh.LeftX, orSh.TopY, orSh.RightX, orSh.BottomY).CreateSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }

        private void btnCrLine_Click(object sender, RoutedEventArgs e)
        {
            if (app?.ActiveDocument == null) return;
            if (app.ActiveSelectionRange.Count < 1)
            {
                MessageBox.Show(err, "Lỗi");
                return;
            }
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
                for (int i = 0; i < numCol.Value; i++)
                {
                    space += s.x;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(space, 0));
                }
                space = 0;
                app.ActiveLayer.CreateLineSegment(position.x, position.y, position.x + size.x, position.y).CreateSelection();
                orSh.AddRange(app.ActiveSelectionRange);
                for (int j = 0; j < numRow.Value; j++)
                {
                    space += s.y;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(0, -space));
                }
                orSh.Group();
                app.ActiveLayer.CreateRectangle(orSh.LeftX, orSh.TopY, orSh.RightX, orSh.BottomY).CreateSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
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
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }

        private void btnDelImg_Click(object sender, RoutedEventArgs e)
        {
            if (app?.ActiveDocument == null) return;
            if (app.ActiveSelectionRange.Count < 1)
            {
                MessageBox.Show(err, "Lỗi");
                return;
            }
            try
            {
                ShapeRange orSh = app.ActiveSelectionRange;
                ShapeRange img = new ShapeRange();
                orSh.UngroupAll();
                foreach(Shape sh in orSh)
                {
                    if (sh.Type == cdrShapeType.cdrBitmapShape)
                        img.Add(sh);
                }
                //MessageBox.Show(img.Count.ToString());
                img.Delete();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
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
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
        private float calSizeInsert()
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
                return (float)Math.Round(d, 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
                return 0;
            }
        }
        private void numInsertChange()
        {
            try
            {
                float size = calSizeInsert();
                if (size <= 0) return;
                if (numInsert.Value == 0) return;
                if (numInsert.Value > size)
                {
                    int round = (int)((numInsert.Value*10 - size*10) / 1.2);
                    numSpace.Value = round + 3;
                }
                else numSpace.Value = 3;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }

        private void btnCreInsert_Click(object sender, RoutedEventArgs e)
        {
            if (app?.ActiveDocument == null) return;
            if (app.ActiveSelectionRange.Count < 1)
            {
                MessageBox.Show(err, "Lỗi");
                return;
            }
            try
            {
                app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                ShapeRange orSh = app.ActiveSelectionRange;
                if (orSh.Shapes[1].Type == cdrShapeType.cdrBitmapShape && orSh.Count == 1)
                {
                    Shape ell = app.ActiveLayer.CreateEllipse2(orSh.CenterX, orSh.CenterY, orSh.SizeWidth / 2, orSh.SizeHeight / 2);
                    Shape newell = ell.Intersect(orSh.Shapes[1], true, true);
                    orSh.Delete();
                    orSh.Add(newell);
                    orSh.Add(ell);
                }
                orSh.CreateSelection();            
                double size = numInsert.Value;
                if (size == 0) return;
                double space = 0;
                ShapeRange insert = app.ActiveSelectionRange.Duplicate(0, 0);
                app.ActiveDocument.ReferencePoint = cdrReferencePoint.cdrBottomRight;
                insert.SizeHeight = size;
                insert.SizeWidth = size;
                app.ActiveDocument.ReferencePoint = cdrReferencePoint.cdrTopLeft;
                double move = (size / 2) + (spaceCal() / 2);
                insert.Move(move, -move);
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
                space = 0;
                insert.CreateSelection();
                ShapeRange insertRange = insert;
                for (int ii = 1; ii < numRow.Value - 1; ii++)
                {
                    space += s.y + spaceCal();
                    insertRange.AddRange(app.ActiveSelectionRange.Duplicate(0, -space));
                }
                space = 0;
                insertRange.CreateSelection();
                for (int jj = 1; jj < numCol.Value - 1; jj++)
                {
                    space += s.x + spaceCal();
                    insertRange.AddRange(app.ActiveSelectionRange.Duplicate(space, 0));
                }
                orSh.AddRange(insertRange);
                orSh.Group();
                app.ActiveLayer.CreateRectangle(orSh.LeftX, orSh.TopY, orSh.RightX, orSh.BottomY).CreateSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }

        private void btnChangeTheme_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Button b = sender as Button;
                switch(b.Tag.ToString())
                {
                    case "lightblue" : ChangeTheme("LightBlue"); writeTheme("LightBlue");  break;
                    case "bluegrey": ChangeTheme("BlueGrey"); writeTheme("BlueGrey"); break;
                    case "green": ChangeTheme("GreenLight"); writeTheme("GreenLight"); break;
                    case "mater": ChangeTheme("MaterialLight"); writeTheme("MaterialLight"); break;
                    case "orange": ChangeTheme("OrangeLight"); writeTheme("OrangeLight"); break;
                    case "pink": ChangeTheme("PinkLight"); writeTheme("PinkLight"); break;
                    case "purple": ChangeTheme("PurpleLight"); writeTheme("PurpleLight"); break;
                    case "red": ChangeTheme("RedLight"); writeTheme("RedLight"); break;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
        private void readTheme()
        {
            try
            {
                string theme = File.ReadAllText(@"Addons\FileBe\color.ini");
                if(theme != null && theme.Length > 0)
                    ChangeTheme(theme);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
        private void writeTheme(string theme)
        {
            try
            {
                File.WriteAllText(@"Addons\FileBe\color.ini", theme);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
        // Calculate size when NumericUpDown value changed
        private void cal_ValueChanged(object sender, RoutedPropertyChangedEventArgs<int> e)
        {  
            calSize();
        }
        // Change value NumericUpDown when Key Up
        private void cal_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            NumericUpDown num = sender as NumericUpDown;
            try
            {
                int i = int.Parse(num.txt.Text);
                    if(i == 1)
                    {
                        num.Value = i;
                        calSize();
                    }
                    else
                    {                       
                        num.Value = i;
                    }    
            }
            catch
            {
                num.Value = 1;
            }
        }

        private void numInsert_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            FloatUpDown num = sender as FloatUpDown;
            try
            {
                float i = float.Parse(num.txt.Text);
                num.Value = i;
            }
            catch
            {
                num.Value = 0;
            }
        }

        private void numInsert_ValueChanged(object sender, RoutedPropertyChangedEventArgs<float> e)
        {
            numInsertChange();
        }

        private void btnCreaRec_Click(object sender, RoutedEventArgs e)
        {
            if (app?.ActiveDocument == null) return;
            if (app.ActiveSelectionRange.Count < 1)
            {
                MessageBox.Show(err, "Lỗi");
                return;
            }
            try
            {
                ShapeRange orSh = app.ActiveSelectionRange;
                app.ActiveLayer.CreateRectangle(orSh.LeftX, orSh.TopY, orSh.RightX, orSh.BottomY);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }

        private void btnCreaEll_Click(object sender, RoutedEventArgs e)
        {
            if (app?.ActiveDocument == null) return;
            if (app.ActiveSelectionRange.Count < 1)
            {
                MessageBox.Show(err, "Lỗi");
                return;
            }
            try
            {
                ShapeRange orSh = app.ActiveSelectionRange;
                app.ActiveLayer.CreateEllipse2(orSh.CenterX, orSh.CenterY, orSh.SizeWidth/2, orSh.SizeHeight/2);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }

        private void btnCreaSize_Click(object sender, RoutedEventArgs e)
        {
            if (app?.ActiveDocument == null) return;
            if (app.ActiveSelectionRange.Count < 1)
            {
                MessageBox.Show(err, "Lỗi");
                return;
            }
            try
            {
                app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
                ShapeRange orSh = app.ActiveSelectionRange;
                Size s = new Size(orSh.SizeWidth, orSh.SizeHeight);
                int point = 0;
                int space = (int)(Math.Sqrt(s.x) / 2);
                if (s.x < 200)
                {
                    point = (int)s.x * 2;
                }
                else
                {
                    point = (int)s.x;
                }
                Shape sizeHeight = app.ActiveLayer.CreateArtisticText(0, 0, Math.Round(s.x / 100, 2).ToString() + "m");
                sizeHeight.Text.FontProperties.Size = point;
                Shape sizeWidth = app.ActiveLayer.CreateArtisticText(0, 0, Math.Round(s.y / 100, 2).ToString() + "m");
                sizeWidth.Text.FontProperties.Size = point;
                orSh.Add(sizeHeight);
                orSh.AlignAndDistribute(cdrAlignDistributeH.cdrAlignDistributeHAlignCenter, cdrAlignDistributeV.cdrAlignDistributeVAlignTop,
                     cdrAlignShapesTo.cdrAlignShapesToLastSelected, cdrDistributeArea.cdrDistributeToSelection, false);
                sizeHeight.Move(0, space + sizeHeight.SizeHeight);
                orSh.Remove(2);
                orSh.Add(sizeWidth);
                orSh.AlignAndDistribute(cdrAlignDistributeH.cdrAlignDistributeHAlignRight, cdrAlignDistributeV.cdrAlignDistributeVAlignCenter,
                     cdrAlignShapesTo.cdrAlignShapesToLastSelected, cdrDistributeArea.cdrDistributeToSelection, false);
                sizeWidth.Move(space + sizeWidth.SizeWidth, 0);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
    }
}
