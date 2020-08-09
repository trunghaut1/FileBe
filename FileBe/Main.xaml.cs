using Corel.Interop.VGCore;
using NumericUpDownLib;
using System;
using System.Collections.Generic;
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
        double _space = 1;
        Size _size = new Size(0, 0);
        public Main()
        {
            InitializeComponent();
        }
        public Main(object _app)
        {
            InitializeComponent();
            app = (Corel.Interop.CorelDRAW.Application)_app;
            numCol.ValueChanged += new RoutedPropertyChangedEventHandler<int>(cal_ValueChanged);
            numRow.ValueChanged += new RoutedPropertyChangedEventHandler<int>(cal_ValueChanged);
            numSpace.ValueChanged += new RoutedPropertyChangedEventHandler<int>(cal_ValueChanged);
            numInsert.ValueChanged += new RoutedPropertyChangedEventHandler<float>(numInsert_ValueChanged);
            numFirstNum.ValueChanged += new RoutedPropertyChangedEventHandler<int>(calFistNum_ValueChanged);
            numLastNum.ValueChanged += new RoutedPropertyChangedEventHandler<int>(calNum_ValueChanged);
            numColNum.ValueChanged += new RoutedPropertyChangedEventHandler<int>(calNum_ValueChanged);
            numSpaceNum.ValueChanged += new RoutedPropertyChangedEventHandler<int>(calNum_ValueChanged);
            readTheme();
            //app.Application.SelectionChange += new DIVGApplicationEvents_SelectionChangeEventHandler(SelectionChanged);
        }
        private void ChangeTheme(string theme)
        {
            string uri = "/FlatTheme;component/ColorStyle/" + theme + ".xaml";
            ResourceDictionary resourceDict = System.Windows.Application.LoadComponent(new Uri(uri, UriKind.Relative)) as ResourceDictionary;
            Resources.MergedDictionaries[0].Clear();
            Resources.MergedDictionaries[0].MergedDictionaries.Add(resourceDict);
        }
        // Khôi phục về giá trị mặc định
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
            _space = 1;
            _size.x = 0;
            _size.y = 0;
        }
        private bool checkActive()
        {
            if (app?.ActiveDocument == null || app.ActiveSelectionRange.Count < 1) return true;
            return false;
        }
        // Tính và hiển thị thông tin file bế
        private void calSize()
        {
            if (checkActive()) return;
            app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
            autoRound();
            Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
            try
            {
                spaceCal();
                double wid = ((s.x + _space) * numCol.Value) - _space;
                double hei = ((s.y + _space) * numRow.Value) - _space;
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
        private void spaceCal()
        {
            if (chkUnSpace.IsChecked.Value)
                _space = 0;
            else
                _space = (double)numSpace.Value / 10;
        }

        private void chkUnSpace_Checked(object sender, RoutedEventArgs e)
        {
            calSize();
        }

        private void btnCreate_Click(object sender, RoutedEventArgs e)
        {
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi",MessageBoxButton.OK,MessageBoxImage.Error);
                return;
            }
            try
            {
                autoRound();
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                ShapeRange orSh = app.ActiveSelectionRange;
                double space = 0;
                for (int i = 1; i < numRow.Value; i++)
                {
                    space += s.y + _space;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(0, -space));
                }
                orSh.CreateSelection();
                space = 0;
                for (int j = 1; j < numCol.Value; j++)
                {
                    space += s.x + _space;
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
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                autoRound();
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                ShapeRange orSh = null;
                Size position = new Size(app.ActiveSelectionRange.PositionX, app.ActiveSelectionRange.PositionY);
                Size size = new Size(s.x * numCol.Value, s.y * numRow.Value);
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
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                autoRound();
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                ShapeRange orSh = app.ActiveSelectionRange;
                Shape sh = app.ActiveLayer.CreateRectangle(orSh.LeftX, orSh.TopY, orSh.RightX, orSh.BottomY);
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
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                ShapeRange orSh = app.ActiveSelectionRange;
                ShapeRange img = new ShapeRange();
                ShapeRange unColor = new ShapeRange();
                orSh.UngroupAll();
                foreach(Shape sh in orSh)
                {
                    if (sh.Type == cdrShapeType.cdrBitmapShape)
                        img.Add(sh);
                    else
                        if (sh.Fill.Type == cdrFillType.cdrNoFill && sh.Outline.Type == cdrOutlineType.cdrNoOutline)
                            unColor.Add(sh);
                }
                //MessageBox.Show(img.Count.ToString());
                img.Delete();
                if(unColor.Count > 0)
                {
                    MessageBoxResult result = MessageBox.Show("Phát hiện viền bế ẩn, bạn có muốn tự động xóa?", "Cảnh báo!", MessageBoxButton.YesNo,MessageBoxImage.Warning);
                    if (result == MessageBoxResult.Yes)
                        unColor.Delete();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }

        private void btnCalSizeInsert_Click(object sender, RoutedEventArgs e)
        {
                numInsert.Value = calSizeInsert();
                numSpace.Value = 3;
        }
        private float calSizeInsert()
        {
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return 0;
            }
            try
            {
                app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
                autoRound();
                ShapeRange orSh = app.ActiveSelectionRange;
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
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
                autoRound();
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                ShapeRange orSh = app.ActiveSelectionRange;
                if (orSh.Shapes[1].Type == cdrShapeType.cdrBitmapShape && orSh.Count == 1)
                {
                    Shape ell = app.ActiveLayer.CreateEllipse2(orSh.CenterX, orSh.CenterY, orSh.SizeWidth / 2, orSh.SizeHeight / 2);
                    Shape newell = ell.Intersect(orSh.Shapes[1], true, true);
                    orSh.Delete();
                    orSh.Add(newell);
                    orSh.Add(ell);
                    orSh.CreateSelection();
                }          
                double size = numInsert.Value;
                if (size == 0) return;
                double space = 0;
                spaceCal();
                ShapeRange insert = app.ActiveSelectionRange.Duplicate(0, 0);
                app.ActiveDocument.ReferencePoint = cdrReferencePoint.cdrBottomRight;
                insert.SizeHeight = size;
                insert.SizeWidth = size;
                app.ActiveDocument.ReferencePoint = cdrReferencePoint.cdrTopLeft;
                double move = (size / 2) + (_space / 2);
                insert.Move(move, -move);
                for (int i = 1; i < numRow.Value; i++)
                {
                    space += s.y + _space;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(0, -space));
                }
                orSh.CreateSelection();
                space = 0;
                for (int j = 1; j < numCol.Value; j++)
                {
                    space += s.x + _space;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(space, 0));
                }
                space = 0;
                insert.CreateSelection();
                ShapeRange insertRange = insert;
                for (int ii = 1; ii < numRow.Value - 1; ii++)
                {
                    space += s.y + _space;
                    insertRange.AddRange(app.ActiveSelectionRange.Duplicate(0, -space));
                }
                space = 0;
                insertRange.CreateSelection();
                for (int jj = 1; jj < numCol.Value - 1; jj++)
                {
                    space += s.x + _space;
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
            string path = Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData) + "\\FileBe\\color.ini";
            try
            {
                if (File.Exists(path))
                {
                    string theme = File.ReadAllText(path);
                    if (theme != null && theme.Length > 0)
                        ChangeTheme(theme);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
        private void writeTheme(string theme)
        {
            string path = Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData) + "\\FileBe\\color.ini";
            FileInfo file = new FileInfo(path);
            try
            {
                file.Directory.Create();
                File.WriteAllText(path, theme);
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
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
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
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
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
        private void CreateSize(bool isMeter)
        {
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
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
                string width, height;
                if(isMeter)
                {
                    width = Math.Round(s.x / 100, 2).ToString() + "m";
                    height = Math.Round(s.y / 100, 2).ToString() + "m";
                }
                else
                {
                    width = Math.Round(s.x, 1).ToString() + "cm";
                    height = Math.Round(s.y, 1).ToString() + "cm";
                }
                Shape sizeHeight = app.ActiveLayer.CreateArtisticText(0, 0, width);
                sizeHeight.Text.FontProperties.Size = point;
                Shape sizeWidth = app.ActiveLayer.CreateArtisticText(0, 0, height);
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

        private void btnCreaSize_Click(object sender, RoutedEventArgs e)
        {
            CreateSize(true);
        }

        private void btnCreaSize_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            CreateSize(false);
        }

        private void modeSelect_Checked(object sender, RoutedEventArgs e)
        {
            taobeGrid.Visibility = Visibility.Collapsed;
            sobanGrid.Visibility = Visibility.Visible;
        }

        private void modeSelect_Unchecked(object sender, RoutedEventArgs e)
        {
            taobeGrid.Visibility = Visibility.Visible;
            sobanGrid.Visibility = Visibility.Collapsed;
        }

        private void btn1mSpace_Click(object sender, RoutedEventArgs e)
        {
            chkUnSpace.IsChecked = false;
            autoCal();
        }

        private void btn1m_Click(object sender, RoutedEventArgs e)
        {
            chkUnSpace.IsChecked = true;
            autoCal();
        }
        private void autoCal()
        {
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                int col = 0, row = 0;
                double w = 0, h = 0;
                if (chkUnSpace.IsChecked.Value)
                {
                    col = (int)(120.5 / s.x);
                    w = s.x * col;
                    row = (int)((10000 / w) / s.y);
                    h = row * s.y;
                    if (w * (row * s.y) < 9400)
                        row++;
                    else
                        if ((w * h) > 9400 && (w * h) < 9750)
                        if (w * (h + s.y) < 10100)
                            row++;
                }
                else
                {
                    col = (int)(120.5 / (s.x + 0.15));
                    w = (s.x + 0.1) * col - 0.1;
                    row = (int)((10000 / w) / (s.y + 0.15));
                    h = row * (s.y + 0.1) - 0.1;
                    if (w * h < 9400)
                        row++;
                    else
                        if ((w * h) > 9400 && (w * h) < 9700)
                        if (w * (h + s.y + 0.1) < 10100)
                            row++;

                }
                numCol.Value = col;
                numRow.Value = row;
                calSize();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
        private void autoRound()
        {
            Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
            if(s.x != _size.x || s.y != _size.y)
            {
                _size.x = Math.Round(s.x, 1);
                _size.y = Math.Round(s.y, 1);
                app.ActiveSelection.SizeWidth = _size.x;
                app.ActiveSelection.SizeHeight = _size.y;                
            }
        }

        private void btnCrLineColor_Click(object sender, RoutedEventArgs e)
        {
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                autoRound();
                Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
                ShapeRange rowRange = app.ActiveSelectionRange;
                ShapeRange colRev = app.ActiveSelectionRange;
                ShapeRange rowRev = app.ActiveSelectionRange;
                colRev.RemoveAll();
                rowRev.RemoveAll();
                rowRange.RemoveAll();
                Size position = new Size(app.ActiveSelectionRange.PositionX, app.ActiveSelectionRange.PositionY);
                Size size = new Size(s.x * numCol.Value, s.y * numRow.Value);
                app.ActiveSelection.Delete();
                Shape shapeCol = app.ActiveLayer.CreateLineSegment(position.x, position.y, position.x, position.y - size.y);
                shapeCol.Outline.Color.CMYKAssign(100, 0, 100, 0);
                double space = 0;
                for (int i = 0; i < numCol.Value; i++)
                {
                    space += s.x;
                    if (i % 2 == 0 || i == numCol.Value - 1)
                        colRev.Add(shapeCol.Duplicate(space, 0));
                    else
                        shapeCol.Duplicate(space, 0);
                }
                shapeCol.OrderToFront();
                colRev.Flip(cdrFlipAxes.cdrFlipVertical);
                space = 0;
                Shape shapeRow = app.ActiveLayer.CreateLineSegment(position.x, position.y, position.x + size.x, position.y);
                shapeRow.Outline.Color.CMYKAssign(0, 100, 100, 0);
                rowRev.Add(shapeRow);
                for (int j = 0; j < numRow.Value; j++)
                {
                    space += s.y;
                    if (j % 2 != 0 || j == numRow.Value - 1)
                        rowRev.Add(shapeRow.Duplicate(0, -space));
                    else
                        rowRange.Add(shapeRow.Duplicate(0, -space));
                }
                shapeRow.OrderToFront();
                rowRange.AddRange(rowRev);
                rowRange.OrderReverse();
                rowRev.Flip(cdrFlipAxes.cdrFlipHorizontal);
                app.ActiveLayer.CreateRectangle2(position.x, position.y - size.y, size.x, size.y).CreateSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }

        private void btnLogo_Click(object sender, RoutedEventArgs e)
        {
            var random = new Random();
            var list = new List<string> { "LightBlue", "BlueGrey", "GreenLight", "MaterialLight", "OrangeLight", "PinkLight", "PurpleLight", "RedLight" };
            int index = random.Next(list.Count);
            ChangeTheme(list[index]);
            writeTheme(list[index]);
        }

        
    }
}
