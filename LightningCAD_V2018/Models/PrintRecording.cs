using System.Collections.Generic;

namespace LightningCAD.Models
{
    /// <summary>
    /// 打印记录
    /// </summary>
    public class PrintRecording
    {
        public string Path { get; set; }
        public List<PrintInfo> PrintInfos { get; set; }
    }
}