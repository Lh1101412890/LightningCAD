using System.ComponentModel;

using Autodesk.AutoCAD.DatabaseServices;

namespace LightningCAD.Models
{
    /// <summary>
    /// 构件
    /// </summary>
    public class DRComponent : INotifyPropertyChanged
    {
        private string name;

        /// <summary>
        /// 编号
        /// </summary>
        public string Name
        {
            get => name;
            set
            {
                name = value;
                RaisePropertyChanged(nameof(Name));
            }
        }

        /// <summary>
        /// 文字对象ID
        /// </summary>
        public ObjectId DBTextObjectId { get; set; }

        /// <summary>
        /// 是否有效
        /// </summary>
        public bool IsValid { get; set; }

        private bool isDifferent;

        public bool IsDifferent
        {
            get => isDifferent;
            set
            {
                isDifferent = value;
                RaisePropertyChanged(nameof(IsDifferent));
            }
        }


        /// <summary>
        /// 构件对象ID
        /// </summary>
        public ObjectId ComponentObjectId { get; set; }
        /// <summary>
        /// 连接线对象ID
        /// </summary>
        public ObjectId LinklineObjectId { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void RaisePropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}