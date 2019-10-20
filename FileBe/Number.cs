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
            chk0.IsChecked = false;
            numColNum.Value = 1;
            numSpaceNum.Value = 1;
            lblRowNum.Content = "0";
            lblHeiNum.Content = "0";
            lblTotalNum.Content = "0";
            lblTotalSizeNum.Content = "0";
            lblWidNum.Content = "0";
        }
        private void btnReCalNum_Click(object sender, RoutedEventArgs e)
        {
            calNum();
        }
        private void calNum()
        {
            if (app?.ActiveDocument == null) return;
            if (app.ActiveSelectionRange.Count < 1) return;
            app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
            Size s = new Size(app.ActiveSelection.SizeWidth, app.ActiveSelection.SizeHeight);
            try
            {
                float space = spaceCalNum();
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
        private float spaceCalNum()
        {
            if (chkUnSpaceNum.IsChecked.Value) return 0;
            return (float)numSpaceNum.Value / 10;
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
            int bitmap = 0, text = 0, rec = 0, temp = 0;
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
                float p = spaceCalNum();
                orSh.Group().CreateSelection();
                orSh = app.ActiveSelectionRange;
                foreach (Shape sp in orSh.Shapes[1].Shapes)
                {
                    temp++;
                    if (sp.Type == cdrShapeType.cdrBitmapShape)
                    {
                        bitmap = temp;
                    }
                    else if (sp.Type == cdrShapeType.cdrTextShape)
                    {
                        text = temp;
                    }
                    else rec = temp;
                }
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
                if(Convert.ToInt32(lblTotalNum.Content) < Convert.ToInt32(lblRowNum.Content) * numColNum.Value)
                {
                    ShapeRange remove = new ShapeRange();
                    for(int i = Convert.ToInt32(lblTotalNum.Content) + 1; i <= Convert.ToInt32(lblRowNum.Content) * numColNum.Value; i++)
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
                    if((bool)chk0.IsChecked)
                        sp.Shapes[text].Text.Contents = count.ToString().PadLeft(lenght, '0');
                    else
                        sp.Shapes[text].Text.Contents = count.ToString();
                    count++;
                }
                //MessageBox.Show("bit: " + bitmap.ToString() + "- text: " + text.ToString() + "- rec: " + rec.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.Source, "Lỗi");
            }
        }
        private void btnGetOutline_Click(object sender, RoutedEventArgs e)
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
                foreach (Shape sh in orSh)
                {
                    if (sh.Type == cdrShapeType.cdrBitmapShape || sh.Type == cdrShapeType.cdrTextShape)
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

        private void btnDelOutline_Click(object sender, RoutedEventArgs e)
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
                foreach (Shape sh in orSh)
                {
                    if (sh.Type != cdrShapeType.cdrBitmapShape && sh.Type != cdrShapeType.cdrTextShape)
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

    }
}
