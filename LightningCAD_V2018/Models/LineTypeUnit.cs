using System.ComponentModel;

namespace LightningCAD.Models
{
    public class LineTypeUnit : INotifyPropertyChanged
    {
        private int number;
        private LineUnit lineUnit;
        private double length;
        private string text;
        private double scale;
        private double rotate;
        private bool isUcs;
        private double x;
        private double y;

        public int Number
        {
            get => number;
            set
            {
                number = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Number)));
            }
        }
        public LineUnit LineUnit
        {
            get => lineUnit;
            set
            {
                lineUnit = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LineUnit)));
            }
        }
        public double Length
        {
            get => length;
            set
            {
                length = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Length)));
            }
        }
        public string Text
        {
            get => text;
            set
            {
                text = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Text)));
            }
        }
        public double Scale
        {
            get => scale;
            set
            {
                scale = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Scale)));
            }
        }
        public double Rotate
        {
            get => rotate;
            set
            {
                rotate = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Rotate)));
            }
        }
        public bool IsUcs
        {
            get => isUcs;
            set
            {
                isUcs = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsUcs)));
            }
        }
        public double X
        {
            get => x;
            set
            {
                x = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(X)));
            }
        }
        public double Y
        {
            get => y;
            set
            {
                y = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Y)));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    public enum LineUnit
    {
        直线,
        空格
    }
}
