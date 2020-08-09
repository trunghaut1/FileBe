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
        private void btnResetNum_Click(object sender, RoutedEventArgs e)
        {
            numFirstNum.Value = 1;
            numLastNum.MinValue = 2;
            numColNum.MaxValue = 2;
            numLastNum.Value = 2;
            numSpaceNum.Value = 1;
            chkUnSpaceNum.IsChecked = false;
            numColNum.Value = 1;
            numSpaceNum.Value = 1;
            lblRowNum.Content = "0";
            lblHeiNum.Content = "0";
            lblTotalNum.Content = "0";
            lblTotalSizeNum.Content = "0";
            lblWidNum.Content = "0";
            txtTextName.Text = "";
            txtTextFirst.Text = "";
            txtTextLast.Text = "";
            chk00.IsChecked = false;
            chkAuto.IsChecked = true;
            num0.Value = 1;
            try
            {
                txtTextName.SetResourceReference(System.Windows.Controls.Control.BorderBrushProperty, "MainColor");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
        private void txtTextName_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if(string.IsNullOrWhiteSpace(txtTextName.Text))
                    txtTextName.SetResourceReference(System.Windows.Controls.Control.BorderBrushProperty, "ErrorColor");
                else
                    txtTextName.SetResourceReference(System.Windows.Controls.Control.BorderBrushProperty, "MainColor");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
            
        }
        private void btnAutoText_Click(object sender, RoutedEventArgs e)
        {
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                ShapeRange orSh = app.ActiveSelectionRange;
                double x = 0, y = 0;
                bool flag = false;
                if (orSh.Shapes.Count == 1 && orSh.Shapes.First.Type == cdrShapeType.cdrGroupShape)
                    orSh.Ungroup();
                foreach (Shape s in orSh.Shapes)
                {
                    if(s.Type == cdrShapeType.cdrTextShape)
                    {
                        txtTextName.Text = s.Text.Contents;
                        x = s.PositionX;
                        y = s.PositionY;
                        s.Text.AlignProperties.Alignment = cdrAlignment.cdrCenterAlignment;
                        s.PositionX = x;
                        s.PositionY = y;
                        flag = true;
                        break;
                    }
                }
                if(!flag)
                    MessageBox.Show("Không tìm thấy đối tượng dạng Text!", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
        private void calNum()
        {
            if (checkActive()) return;
            app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
            Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
            try
            {
                double space = spaceCalNum();
                double row = (numLastNum.Value - numFirstNum.Value + 1.0) / numColNum.Value;
                int realRow = (int)Math.Ceiling(row);
                double wid = ((s.x + space) * numColNum.Value) - space;
                double hei = ((s.y + space) * realRow) - space;
                lblRowNum.Content = realRow.ToString();
                lblWidNum.Content = string.Format("{0:#,##0.###}", wid);
                lblHeiNum.Content = string.Format("{0:#,##0.###}", hei);
                lblTotalSizeNum.Content = string.Format("{0:#,##0.###}", ((wid * hei) / 10000));
                lblTotalNum.Content = (numLastNum.Value - numFirstNum.Value + 1).ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
        private double spaceCalNum()
        {
            if (chkUnSpaceNum.IsChecked.Value) return 0;
            return (double)numSpaceNum.Value / 10;
        }
        private void calMaxCol()
        {
            int max = numLastNum.Value - numFirstNum.Value;
            numColNum.MaxValue = max + 1;
        }
        private void calFirstNum_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            NumericUpDown num = sender as NumericUpDown;
            try
            {
                int i = int.Parse(num.txt.Text);
                if (i == 1)
                {
                    num.Value = i;
                    calMaxCol();
                    calNum();
                }
                else
                {
                    num.Value = i;
                }
                if (numLastNum.Value <= i)
                {
                    numLastNum.Value = i + 1;
                }
                numLastNum.MinValue = i + 1;
            }
            catch
            {
                num.Value = 1;
            }
        }
        private void calNum_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            NumericUpDown num = sender as NumericUpDown;
            try
            {
                int i = int.Parse(num.txt.Text);
                if (i == 1)
                {
                    num.Value = i;
                    calMaxCol();
                    calNum();
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
        private void calLastNum_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            NumericUpDown num = sender as NumericUpDown;
            try
            {
                int i = int.Parse(num.txt.Text);
                if (i == 2)
                {
                    num.Value = i;
                    calMaxCol();
                    calNum();
                }
                else
                {
                    num.Value = i;
                }
            }
            catch
            {
                num.Value = 2;
            }
        }
        private void numLastNum_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            numLastNum.MinValue = 1;
        }
        private void numLastNum_MouseLeave(object sender, MouseEventArgs e)
        {
            if (numLastNum.Value <= numFirstNum.Value)
            {
                numLastNum.Value = numFirstNum.Value + 1;
            }
            numLastNum.MinValue = numFirstNum.Value + 1;
        }
        private void calFistNum_ValueChanged(object sender, RoutedPropertyChangedEventArgs<int> e)
        {
            if (numLastNum.Value <= numFirstNum.Value)
            {
                numLastNum.Value = numFirstNum.Value + 1;
            }
            numLastNum.MinValue = numFirstNum.Value + 1;
            calMaxCol();
            calNum();
        }
        private void calNum_ValueChanged(object sender, RoutedPropertyChangedEventArgs<int> e)
        {
            calMaxCol();
            calNum();
        }
        private void chkUnSpaceNum_Checked(object sender, RoutedEventArgs e)
        {
            //calMaxCol();
            calNum();
        }
        private void btnCreaNum_Click(object sender, RoutedEventArgs e)
        {
            if (checkActive())
            {
                MessageBox.Show(err, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                if(string.IsNullOrWhiteSpace(txtTextName.Text))
                {
                    txtTextName.SetResourceReference(System.Windows.Controls.Control.BorderBrushProperty, "ErrorColor");
                    MessageBox.Show("Không được để trống phần số gốc đang chọn!", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                ShapeRange orSh = app.ActiveSelectionRange;
                bool flag = false;
                double x = 0, y = 0;
                if (orSh.Shapes.Count == 1 && orSh.Shapes.First.Type == cdrShapeType.cdrGroupShape)
                    orSh.Ungroup();
                foreach (Shape ss in orSh.Shapes)
                {
                    if (ss.Type == cdrShapeType.cdrTextShape)
                    {
                        if(ss.Text.Contents == txtTextName.Text)
                        {
                            x = ss.PositionX;
                            y = ss.PositionY;
                            ss.Text.AlignProperties.Alignment = cdrAlignment.cdrCenterAlignment;
                            ss.PositionX = x;
                            ss.PositionY = y;
                            flag = true;
                            ss.Name = "txtName";
                            break;
                        }   
                    }
                }
                if (!flag)
                {
                    MessageBox.Show("Không tìm thấy Text có nội dung '"+txtTextName.Text+"'!", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                orSh.Group().CreateSelection();
                Size s = new Size(orSh.SizeWidth, orSh.SizeHeight);
                double space = 0;
                double p = spaceCalNum();
                orSh = app.ActiveSelectionRange;
                for (int j = 1; j < numColNum.Value; j++)
                {
                    space += s.x + p;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(space, 0));
                }
                orSh.CreateSelection();
                space = 0;
                for (int i = 1; i < Convert.ToInt32(lblRowNum.Content); i++)
                {
                    space += s.y + p;
                    orSh.AddRange(app.ActiveSelectionRange.Duplicate(0, -space));
                }
                if (Convert.ToInt32(lblTotalNum.Content) < Convert.ToInt32(lblRowNum.Content) * numColNum.Value)
                {
                    ShapeRange remove = new ShapeRange();
                    for (int i = Convert.ToInt32(lblTotalNum.Content) + 1; i <= Convert.ToInt32(lblRowNum.Content) * numColNum.Value; i++)
                    {
                        remove.Add(orSh.Shapes[i]);
                    }
                    remove.Delete();
                }
                orSh.CreateSelection();
                orSh = app.ActiveSelectionRange;
                int count = numFirstNum.Value;
                int lenght = numLastNum.Value.ToString().Length;
                foreach (Shape sp in orSh)
                {
                    if (chk00.IsChecked.Value)
                    {
                        if(chkAuto.IsChecked.Value)
                            sp.Shapes["txtName"].Text.Contents = txtTextFirst.Text + count.ToString().PadLeft(lenght, '0') + txtTextLast.Text;
                        else
                        {
                            string num0String = new string('0', num0.Value);
                            sp.Shapes["txtName"].Text.Contents = txtTextFirst.Text + num0String + count.ToString() + txtTextLast.Text;
                        }
                    }  
                    else
                        sp.Shapes["txtName"].Text.Contents = txtTextFirst.Text + count.ToString() + txtTextLast.Text;
                    count++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
        private void btnGetOutline_Click(object sender, RoutedEventArgs e)
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
                orSh.Ungroup();
                Color black = new Color();
                black.CMYKAssign(0, 0, 0, 100);
                foreach (Shape sh in orSh)
                {
                    if (sh.Type == cdrShapeType.cdrCurveShape || sh.Type == cdrShapeType.cdrEllipseShape || sh.Type == cdrShapeType.cdrPolygonShape
                        || sh.Type == cdrShapeType.cdrRectangleShape || sh.Type == cdrShapeType.cdrPerfectShape || sh.Type == cdrShapeType.cdrCustomShape)
                    {
                        if(sh.Fill.Type != cdrFillType.cdrNoFill)
                            img.Add(sh);
                        else
                            if(sh.Outline.Type == cdrOutlineType.cdrNoOutline)
                                img.Add(sh);
                            else
                                if(!sh.Outline.Color.IsSame(black))
                                    img.Add(sh);
                    }
                    else
                        img.Add(sh);
                }
                img.Delete();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }

        private void btnDelOutline_Click(object sender, RoutedEventArgs e)
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
                orSh.Ungroup();
                Color black = new Color();
                black.CMYKAssign(0, 0, 0, 100);
                foreach (Shape sh in orSh)
                {
                    if (sh.Type == cdrShapeType.cdrCurveShape || sh.Type == cdrShapeType.cdrEllipseShape || sh.Type == cdrShapeType.cdrPolygonShape
                        || sh.Type == cdrShapeType.cdrRectangleShape || sh.Type == cdrShapeType.cdrPerfectShape || sh.Type == cdrShapeType.cdrCustomShape)
                    {
                        if (sh.Outline.Type != cdrOutlineType.cdrNoOutline && sh.Fill.Type == cdrFillType.cdrNoFill)
                            if(sh.Outline.Color.IsSame(black))
                                img.Add(sh);
                    }
                }
                img.Delete();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }

    }
}
