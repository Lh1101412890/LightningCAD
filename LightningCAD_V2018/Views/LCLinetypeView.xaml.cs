using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Xml;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using Lightning.Extension;
using Lightning.Information;

using LightningCAD.Commands;
using LightningCAD.Extension;
using LightningCAD.LightningExtension;
using LightningCAD.Models;

using Button = System.Windows.Controls.Button;
using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MessageBox = System.Windows.MessageBox;

namespace LightningCAD.Views
{
    /// <summary>
    /// LCLinetypeView.xaml 的交互逻辑
    /// </summary>
    public partial class LCLinetypeView : ViewBase
    {
        private const string group = "Linetype";
        private const string nameKey = "Name";
        private const string descriptionKey = "Description";
        private const string countKey = "Count";
        private const string unitKey = "Unit";
        private const string lengthKey = "Length";
        private const string textKey = "Text";
        private const string isUcsKey = "IsUcs";
        private const string rotateKey = "Rotate";
        private const string scaleKey = "Scale";
        private const string xKey = "X";
        private const string yKey = "Y";

        private List<LineTypeUnit> Models;
        public LCLinetypeView()
        {
            InitializeComponent();
            up.Source = Information.GetFileInfo("Commands\\向上箭头.png").ToBitmapImage();
            down.Source = Information.GetFileInfo("Commands\\向下箭头.png").ToBitmapImage();
            Models = new List<LineTypeUnit>();
            Read();
            Closed += LCLinetypeView_Closed;
        }

        private void LCLinetypeView_Closed(object sender, EventArgs e) => Write();

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            LineTypeUnit model = new LineTypeUnit()
            {
                Number = Models.Count + 1,
                LineUnit = LineUnit.直线,
                Length = 10,
                Text = "",
                Scale = 1,
                Rotate = 0,
                IsUcs = false,
                X = 0,
                Y = 0,
            };
            Models.Add(model);
            dataGrid.DataContext = null;
            dataGrid.DataContext = Models;
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (dataGrid.SelectedItems.Count == 0)
            {
                return;
            }
            List<LineTypeUnit> models = dataGrid.SelectedItems.Cast<LineTypeUnit>().ToList();
            Models = Models.Except(models).ToList();
            for (int i = 0; i < Models.Count; i++)
            {
                Models[i].Number = i + 1;
            }
            dataGrid.DataContext = null;
            dataGrid.DataContext = Models;
        }

        private void MoveNumberButton_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            int selected = dataGrid.SelectedItems.Count;
            if (selected != 1) return;//没选择返回
            LineTypeUnit model = (dataGrid.SelectedItem as LineTypeUnit);
            int number = model.Number;

            if ((sender as Button).Name == "uper")
            {
                if (number == 1) return;//选择的是第一个无法上移
                Models[number - 2].Number = number;
                Models[number - 1].Number = number - 1;
            }
            else
            {
                if (number == Models.Count) return;//选择的是第一个无法下移
                Models[number - 1].Number = number + 1;
                Models[number].Number = number;
            }
            Models = Models.OrderBy(p => p.Number).Cast<LineTypeUnit>().ToList();//按顺序重新排序
            dataGrid.DataContext = null;
            dataGrid.DataContext = Models;
        }

        private void Complete_Click(object sender, RoutedEventArgs e)
        {
            LSetting.CreatStyles(document);
            using (DocumentLock @lock = document.LockDocument())
            {
                using (Transaction transaction = database.NewTransaction())
                {
                    LinetypeTable linetypeTable = (LinetypeTable)transaction.GetObject(database.LinetypeTableId, OpenMode.ForWrite);
                    if (string.IsNullOrWhiteSpace(name.Text))
                    {
                        return;
                    }
                    //自定义线型
                    if (!linetypeTable.Has(name.Text))
                    {
                        LinetypeTableRecord linetype = new LinetypeTableRecord()
                        {
                            Name = name.Text,
                            AsciiDescription = description.Text,
                            PatternLength = (double)Models.Sum(m => m.Length),
                            NumDashes = Models.Count,
                        };
                        for (int i = 0; i < Models.Count; i++)
                        {
                            double length = Models[i].LineUnit == LineUnit.直线 ? Models[i].Length : -1 * Models[i].Length;
                            linetype.SetDashLengthAt(i, length);//笔画长度
                            if (!string.IsNullOrWhiteSpace(Models[i].Text))
                            {
                                linetype.SetTextAt(i, Models[i].Text);
                                linetype.SetShapeStyleAt(i, document.GetLineTextObjectId());//设置文字的样式
                                linetype.SetShapeNumberAt(i, 0);//设置空格处包含的图案图形
                                linetype.SetShapeScaleAt(i, Models[i].Scale);//图形的缩放比例
                                linetype.SetShapeRotationAt(i, Models[i].Rotate * Math.PI / 180);//图形的旋转弧度
                                linetype.SetShapeIsUcsOrientedAt(i, Models[i].IsUcs);//图形是否面向ucs方向
                                linetype.SetShapeOffsetAt(i, new Vector2d(Models[i].X, Models[i].Y));//图形在X轴Y轴方向上的偏移单位
                            }
                        }
                        linetypeTable.Add(linetype);
                        transaction.AddNewlyCreatedDBObject(linetype, true);
                        editor.WriteMessage($"已添加线型 {name.Text}\n");
                        Dispatcher.Invoke(Close);
                    }
                    else
                    {
                        MessageBox.Show("该名称已存在！");
                    }
                    transaction.Commit();
                }
            }
        }

        private void Read()
        {
            Information.God.GetValue(group, nameKey, out object nameValue);
            if (nameValue == null)
            {
                name.Text = "Lightning";
                description.Text = "----L----L----L----";
                Models.AddRange(new LineTypeUnit[]
                {
                    new LineTypeUnit()
                    {
                        Number = 1,
                        LineUnit = LineUnit.直线,
                        Length = 15,
                        Text = "",
                        IsUcs = false,
                        Rotate = 0,
                        Scale = 1,
                        X = 0,
                        Y = 0,
                    },
                    new LineTypeUnit()
                    {
                        Number = 2,
                        LineUnit = LineUnit.空格,
                        Length = 2.5,
                        Text = "L",
                        IsUcs = false,
                        Rotate = 0,
                        Scale = 1,
                        X = -1.225,
                        Y = -1.75,
                    },
                    new LineTypeUnit() {
                        Number = 3,
                        LineUnit = LineUnit.空格,
                        Length = 2.5,
                        Text = "",
                        IsUcs = false,
                        Rotate = 0,
                        Scale = 1,
                        X = 0,
                        Y = 0,
                    },
                });
            }
            else
            {
                Information.God.GetValue(group, descriptionKey, out object descriptionValue);
                name.Text = nameValue.ToString();
                description.Text = descriptionValue.ToString();
                Information.God.GetValue(group, countKey, out object countValue);
                int count = (int)countValue;
                for (int i = 0; i < count; i++)
                {
                    Information.God.GetValue($"{group}\\Unit{i + 1}", unitKey, out object unitValue);
                    Information.God.GetValue($"{group}\\Unit{i + 1}", lengthKey, out object lengthValue);
                    Information.God.GetValue($"{group}\\Unit{i + 1}", textKey, out object textValue);
                    Information.God.GetValue($"{group}\\Unit{i + 1}", isUcsKey, out object isUcsValue);
                    Information.God.GetValue($"{group}\\Unit{i + 1}", rotateKey, out object rotateValue);
                    Information.God.GetValue($"{group}\\Unit{i + 1}", scaleKey, out object scaleValue);
                    Information.God.GetValue($"{group}\\Unit{i + 1}", xKey, out object xValue);
                    Information.God.GetValue($"{group}\\Unit{i + 1}", yKey, out object yValue);
                    Models.Add(new LineTypeUnit()
                    {
                        Number = i + 1,
                        LineUnit = (LineUnit)Enum.Parse(typeof(LineUnit), unitValue.ToString()),
                        Length = double.Parse((string)lengthValue),
                        Text = textValue.ToString(),
                        IsUcs = bool.Parse((string)isUcsValue),
                        Rotate = double.Parse((string)rotateValue),
                        Scale = double.Parse((string)scaleValue),
                        X = double.Parse((string)xValue),
                        Y = double.Parse((string)yValue),
                    });
                }
            }
            dataGrid.DataContext = Models;
        }

        private void Write()
        {
            Information.God.SetValue(group, nameKey, name.Text);
            Information.God.SetValue(group, descriptionKey, description.Text);
            Information.God.SetValue(group, countKey, Models.Count);
            for (int i = 0; i < Models.Count; i++)
            {
                Information.God.SetValue($"{group}\\Unit{i + 1}", unitKey, Models[i].LineUnit.ToString());
                Information.God.SetValue($"{group}\\Unit{i + 1}", lengthKey, Models[i].Length.ToString());
                Information.God.SetValue($"{group}\\Unit{i + 1}", textKey, Models[i].Text.ToString());
                Information.God.SetValue($"{group}\\Unit{i + 1}", isUcsKey, Models[i].IsUcs.ToString());
                Information.God.SetValue($"{group}\\Unit{i + 1}", rotateKey, Models[i].Rotate.ToString());
                Information.God.SetValue($"{group}\\Unit{i + 1}", scaleKey, Models[i].Scale.ToString());
                Information.God.SetValue($"{group}\\Unit{i + 1}", xKey, Models[i].X.ToString());
                Information.God.SetValue($"{group}\\Unit{i + 1}", yKey, Models[i].Y.ToString());
            }
        }
    }
}