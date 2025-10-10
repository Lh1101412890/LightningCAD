using System;
using System.Collections.Generic;
using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 提取文字
    /// </summary>
    public class LTExtract : CommandBase
    {
        [CommandMethod(nameof(LTExtract), CommandFlags.UsePickSet)]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Editor editor = document.Editor;

                //提取所有的单行文字和多行文字
                PromptSelectionOptions options = new PromptSelectionOptions()
                {
                    MessageForAdding = "请选择需提取的文字:",
                };
                TypedValue[] typedValues = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<or"),
                    new TypedValue((int)DxfCode.Start, "text"),
                    new TypedValue((int)DxfCode.Start, "mtext"),
                    new TypedValue((int)DxfCode.Operator, "or>"),
                };
                SelectionFilter filter = new SelectionFilter(typedValues);
                PromptSelectionResult result = editor.GetSelection(options, filter);
                if (result.Status != PromptStatus.OK)
                {
                    return;
                }
                List<DBObject> objects = document.GetObjects(result.Value.GetObjectIds());
                List<DBText> dBTexts = objects.OfType<DBText>().ToList();
                List<MText> mTexts = objects.OfType<MText>().ToList();

                //将文字添加至entities
                List<Entity> entities = new List<Entity>();
                entities.AddRange(dBTexts);
                entities.AddRange(mTexts);

                if (dBTexts.Count + mTexts.Count == 0)
                {
                    editor.WriteMessage("选择对象不包含文字信息!\n");
                    return;
                }

                List<Entity> newEntities = Sort(entities);

                string str = "";
                foreach (var item in newEntities)
                {
                    if (item is DBText text)
                    {
                        str += text.TextString;
                    }
                    if (item is MText text1)
                    {
                        str += text1.Text;
                    }
                }
                //拷贝到剪贴板
                System.Windows.Clipboard.SetText(str);
                editor.WriteMessage("已拷贝至剪贴板!\n");
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }

        /// <summary>
        /// 将entities中文字逐行排序并存放至newEntities
        /// </summary>
        /// <param name="entities">选择的文字</param>
        /// <param name="newEntities">排序后的文字</param>
        private static List<Entity> Sort(List<Entity> entities)
        {
            if (entities.Count == 0)
            {
                return new List<Entity>();
            }

            //找到最顶部的文字
            double maxY = entities.Max(it => it.Bounds.Value.GetCenter().Y);

            List<Entity> textsLine = new List<Entity>();
            foreach (Entity text in entities)
            {
                //文字高度的1/3
                double height = entities.First().Bounds.Value.GetHeight() / 3;
                //找到同一行的文字
                if (Math.Abs(text.Bounds.Value.GetCenter().Y - maxY) < height)
                {
                    textsLine.Add(text);
                }
            }
            //将已提取的文字移除
            foreach (Entity text in textsLine)
            {
                entities.Remove(text);
            }
            //将一行的文字从左至右排序
            List<Entity> temp = textsLine.OrderBy(it => it.Bounds.Value.GetCenter().X).ToList();
            //如果文字与文字间有间距则添加一个tab符号
            for (int i = 0; i < temp.Count - 1; i++)
            {
                //下一个文字的左侧x数值
                double nextLeft = temp[i + 1].Bounds.Value.MinPoint.X;
                //当前文字的右侧x数值
                double nowRight = temp[i].Bounds.Value.MaxPoint.X;
                //当前文字高度
                double height = (temp[i] as DBText).Height;
                //间隔一个文字高度则添加一个tab
                if (nextLeft - nowRight > height)
                {
                    temp.Insert(i + 1, new DBText() { TextString = "\t" });
                    i++;
                }
            }
            List<Entity> newEntities = new List<Entity>();
            newEntities.AddRange(temp);
            newEntities.Add(new DBText() { TextString = "\n" });//末尾换行
            newEntities.AddRange(Sort(entities));
            return newEntities;
        }
    }
}
