using System.Collections.Generic;
using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using Lightning.Extension;

using LightningCAD.LightningExtension;

namespace LightningCAD.Extension
{
    public static class DocumentExtension
    {
        private const string DefaultDescription = "LightningCAD";

        /// <summary>
        /// 发送取消命令
        /// </summary>
        /// <param name="document"></param>
        public static void CancelCommand(this Document document)
        {
            try
            {
                //关闭窗口后取消命令,x03代表取消命令
                document.SendStringToExecute("\x03", true, false, false);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }

        /// <summary>
        /// 在document的指定图层中绘制entities
        /// </summary>
        /// <param name="entities"></param>
        /// <param name="document"></param>
        /// <param name="layer"></param>
        /// <returns>绘制失败返回-1，绘制成功返回绘制entities的个数</returns>
        public static void Drawing(this Document document, IEnumerable<Entity> entities, string layer = LayerName.Lightning, ColorEnum color = ColorEnum.Red, string description = DefaultDescription)
        {
            if (!entities.Any()) return;
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                //查询是否存在Lightning图层，如果没有则创建
                using (Transaction transaction = database.NewTransaction())
                {
                    document.CreatLayer(layer, color, description);
                    BlockTable blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord blockRecord = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    foreach (Entity entity in entities)
                    {
                        entity.Layer = layer;
                        blockRecord.AppendEntity(entity);
                        transaction.AddNewlyCreatedDBObject(entity, true);
                    }
                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// 在document的指定图层中绘制entity
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="document"></param>
        /// <param name="layer"></param>
        /// <returns>绘制失败返回-1，成功返回1</returns>
        public static void Drawing(this Document document, Entity entity, string layer = LayerName.Lightning, ColorEnum color = ColorEnum.Red, string description = DefaultDescription)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                using (Transaction transaction = database.NewTransaction())
                {
                    document.CreatLayer(layer, color, description);
                    BlockTable blockTable = transaction.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord blockRecord = transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    entity.Layer = layer;
                    blockRecord.AppendEntity(entity);
                    transaction.AddNewlyCreatedDBObject(entity, true);
                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// 删除对象
        /// </summary>
        /// <param name="object"></param>
        public static void Delete(this Document document, DBObject @object)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                using (Transaction transaction = document.Database.NewTransaction())
                {
                    @object.ObjectId.GetObject(OpenMode.ForWrite);
                    @object.Erase(true);
                    @object.Dispose();
                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// 删除对象
        /// </summary>
        /// <param name="object"></param>
        public static void Delete(this Document document, ObjectId objectId)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                using (Transaction transaction = document.Database.NewTransaction())
                {
                    DBObject dBObject = objectId.GetObject(OpenMode.ForWrite);
                    dBObject.Erase(true);
                    dBObject.Dispose();
                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// 解锁图层
        /// </summary>
        /// <param name="document"></param>
        /// <param name="layer"></param>
        public static void UnLockLayer(this Document document, string layer)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                using (Transaction transaction = database.NewTransaction())
                {
                    LayerTable layerTable = transaction.GetObject(database.LayerTableId, OpenMode.ForRead) as LayerTable;
                    if (layerTable.Has(layer))
                    {
                        // 获取图层对象
                        LayerTableRecord layerTableRecord = (LayerTableRecord)transaction.GetObject(layerTable[layer], OpenMode.ForWrite);
                        layerTableRecord.IsLocked = false;
                    }
                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// 锁定图层
        /// </summary>
        /// <param name="document"></param>
        /// <param name="layer"></param>
        public static void LockLayer(this Document document, string layer)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                using (Transaction transaction = database.NewTransaction())
                {
                    LayerTable layerTable = (LayerTable)transaction.GetObject(database.LayerTableId, OpenMode.ForRead);
                    if (layerTable.Has(layer))
                    {
                        // 获取图层对象
                        LayerTableRecord layerTableRecord = (LayerTableRecord)transaction.GetObject(layerTable[layer], OpenMode.ForWrite);
                        layerTableRecord.IsLocked = true;
                    }
                    transaction.Commit();
                }
            }
        }

        public static void CreatLayer(this Document document, string layer, ColorEnum color = ColorEnum.White, string description = DefaultDescription)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                using (Transaction transaction = database.NewTransaction())
                {
                    LayerTable layerTable = transaction.GetObject(database.LayerTableId, OpenMode.ForRead) as LayerTable;
                    if (!layerTable.Has(layer))
                    {
                        LayerTableRecord layerTableRecord = new LayerTableRecord()
                        {
                            Name = layer,
                            Color = color.ToColor(),
                        };
                        layerTable.UpgradeOpen();
                        layerTable.Add(layerTableRecord);
                        transaction.AddNewlyCreatedDBObject(layerTableRecord, true);
                        layerTableRecord.Id.GetObject(OpenMode.ForWrite);
                        layerTableRecord.Description = description;
                    }
                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// 删除图层
        /// </summary>
        /// <param name="document"></param>
        /// <param name="layer"></param>
        public static void DeleteLayer(this Document document, string layer)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                Editor editor = document.Editor;
                using (Transaction transaction = database.NewTransaction())
                {
                    LayerTable layerTable = (LayerTable)transaction.GetObject(database.LayerTableId, OpenMode.ForRead);
                    if (layerTable.Has(layer))
                    {
                        TypedValue[] typedValues = new TypedValue[]
                        {
                            new TypedValue((int)DxfCode.LayerName, layer)
                        };
                        PromptSelectionResult result = editor.SelectAll(new SelectionFilter(typedValues));
                        if (result.Status == PromptStatus.OK)
                        {
                            //删除图层上的对象
                            foreach (var item in result.Value.GetObjectIds())
                            {
                                item.GetObject(OpenMode.ForWrite).Erase(true);
                            }
                        }
                        // 获取图层对象
                        LayerTableRecord layerTableRecord = (LayerTableRecord)transaction.GetObject(layerTable[layer], OpenMode.ForWrite);
                        if (database.Clayer == layerTableRecord.Id)
                        {
                            database.Clayer = ((LayerTableRecord)transaction.GetObject(layerTable["0"], OpenMode.ForRead)).Id;
                        }
                        // 删除图层
                        layerTableRecord.Erase();
                    }
                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// 获取LText文字样式的objectId
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static ObjectId GetLTextObjectId(this Document document)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                ObjectId objectId = ObjectId.Null;
                using (Transaction transaction = database.NewTransaction())
                {
                    //打开字体样式表
                    TextStyleTable textStyleTable = (TextStyleTable)transaction.GetObject(database.TextStyleTableId, OpenMode.ForRead);
                    if (!textStyleTable.Has("LText"))
                    {
                        TextStyleTableRecord text = new TextStyleTableRecord()
                        {
                            Name = "LText",
                            FileName = "gbenor.shx",
                            BigFontFileName = "gbcbig.shx",
                            TextSize = 0,
                        };
                        textStyleTable.UpgradeOpen();
                        textStyleTable.Add(text);
                        transaction.AddNewlyCreatedDBObject(text, true);
                    }
                    objectId = textStyleTable["LText"];
                    transaction.Commit();
                }
                return objectId;
            }
        }

        /// <summary>
        /// 获取LineText文字样式的objectId
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static ObjectId GetLineTextObjectId(this Document document)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                ObjectId objectId = ObjectId.Null;
                using (Transaction transaction = database.NewTransaction())
                {
                    //打开字体样式表
                    TextStyleTable textStyleTable = (TextStyleTable)transaction.GetObject(database.TextStyleTableId, OpenMode.ForRead);
                    if (!textStyleTable.Has("LineText"))
                    {
                        TextStyleTableRecord text = new TextStyleTableRecord()
                        {
                            Name = "LineText",
                            FileName = "gbenor.shx",
                            BigFontFileName = "gbcbig.shx",
                            TextSize = 3.5,
                        };
                        textStyleTable.UpgradeOpen();
                        textStyleTable.Add(text);
                        transaction.AddNewlyCreatedDBObject(text, true);
                    }
                    objectId = textStyleTable["LineText"];
                    transaction.Commit();
                }
                return objectId;
            }
        }

        public static List<DBObject> GetObjects(this Document document, IEnumerable<ObjectId> objectIds, bool openErased = false, bool forceOpenOnLockedLayer = true)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                List<DBObject> objects = new List<DBObject>();
                using (Transaction transaction = database.NewTransaction())
                {
                    foreach (ObjectId objectId in objectIds)
                    {
                        DBObject dBObject = objectId.GetObject(OpenMode.ForRead, openErased, forceOpenOnLockedLayer);
                        objects.Add(dBObject);
                    }
                    transaction.Commit();
                    return objects;
                }
            }
        }

        public static DBObject GetObject(this Document document, ObjectId objectId, bool openErased = false, bool forceOpenOnLockedLayer = true)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                using (Transaction transaction = database.NewTransaction())
                {
                    DBObject dBObject = objectId.GetObject(OpenMode.ForRead, openErased, forceOpenOnLockedLayer);
                    transaction.Commit();
                    return dBObject;
                }
            }
        }

        /// <summary>
        /// 注册Lightning
        /// </summary>
        /// <param name="document"></param>
        public static void RegLightning(this Document document)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                using (Transaction transaction = database.NewTransaction())
                {
                    RegAppTable regAppTable = (RegAppTable)transaction.GetObject(database.RegAppTableId, OpenMode.ForRead);
                    if (!regAppTable.Has("Lightning"))
                    {
                        RegAppTableRecord regApp = new RegAppTableRecord
                        {
                            Name = "Lightning",
                        };
                        regAppTable.UpgradeOpen();
                        regAppTable.Add(regApp);
                        transaction.AddNewlyCreatedDBObject(regApp, true);
                    }
                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// 设置对象的XData数据
        /// </summary>
        /// <param name="document"></param>
        /// <param name="dbobject"></param>
        public static void SetXData(this Document document, DBObject dbobject, ResultBuffer values)
        {
            document.RegLightning();
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                using (Transaction transaction = database.NewTransaction())
                {
                    dbobject.ObjectId.GetObject(OpenMode.ForWrite).XData = values;
                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// 读取XData数据
        /// </summary>
        /// <param name="dbobject"></param>
        /// <returns></returns>
        public static ResultBuffer GetXData(this Document document, DBObject dbobject)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                using (Transaction transaction = document.Database.NewTransaction())
                {
                    ResultBuffer resultBuffer = dbobject.GetXDataForApplication(Information.Brand);
                    transaction.Commit();
                    return resultBuffer;
                }
            }
        }

        public static void Zoom(this Document document, ObjectId objectId, double scale)
        {
            DBObject dBObject = document.GetObject(objectId);
            Editor editor = document.Editor;
            using (ViewTableRecord view = editor.GetCurrentView())
            {
                view.Height = dBObject.Bounds.Value.GetHeight() * scale;
                view.Width = dBObject.Bounds.Value.GetWidth() * scale;
                view.CenterPoint = dBObject.Bounds.Value.GetCenter().ToPoint2d();
                editor.SetCurrentView(view);
            }
            dBObject.Dispose();
        }
    }
}