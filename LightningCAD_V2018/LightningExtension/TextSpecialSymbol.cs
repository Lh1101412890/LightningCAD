namespace LightningCAD.LightningExtension
{
    /// <summary>
    /// 文字符号
    /// </summary>
    public readonly struct TextSpecialSymbol
    {
        /// <summary>
        /// ₂（下标2）
        /// </summary>
        public static readonly string Subscript2 = @"\U+2082";
        /// <summary>
        /// ²（平方）
        /// </summary>
        public static readonly string Square = @"\U+00B2";
        /// <summary>
        /// ³（立方）
        /// </summary>
        public static readonly string Cube = @"\U+00B3";
        /// <summary>
        /// α
        /// </summary>
        public static readonly string Alpha = @"\U+03B1";
        /// <summary>
        /// β
        /// </summary>
        public static readonly string Belta = @"\U+03B2";
        /// <summary>
        /// γ
        /// </summary>
        public static readonly string Gamma = @"\U+03B3";
        /// <summary>
        /// θ
        /// </summary>
        public static readonly string Theta = @"\U+03B8";
        /// <summary>
        /// 一级钢（不能显示）
        /// </summary>
        public static readonly string SteelBar1 = @"\U+0082";
        /// <summary>
        /// 二级钢（不能显示）
        /// </summary>
        public static readonly string SteelBar2 = @"\U+0083";
        /// <summary>
        /// 三级钢（不能显示）
        /// </summary>
        public static readonly string SteelBar3 = @"\U+0084";
        /// <summary>
        /// 四级钢（不能显示）
        /// </summary>
        public static readonly string SteelBar4 = @"\U+0085";
        /// <summary>
        /// °
        /// </summary>
        public static readonly string Degree = @"\U+00B0";
        /// <summary>
        /// ±
        /// </summary>
        public static readonly string Tolerance = @"\U+00B1";
        /// <summary>
        /// φ
        /// </summary>
        public static readonly string Diameter = @"\U+00D8";
        /// <summary>
        /// ∠
        /// </summary>
        public static readonly string Angle = @"\U+2220";
        /// <summary>
        /// ≈
        /// </summary>
        public static readonly string AlmostEqual = @"\U+2248";
        /// <summary>
        /// ≡
        /// </summary>
        public static readonly string AllEqual = @"\U+2261";
        /// <summary>
        /// Ω
        /// </summary>
        public static readonly string Omega = @"\U+03A9";
        /// <summary>
        /// Δ
        /// </summary>
        public static readonly string Delta = @"\U+0394";
        /// <summary>
        /// Φ
        /// </summary>
        public static readonly string ElectricalPhase = @"\U+0278";
        /// <summary>
        /// ≠
        /// </summary>
        public static readonly string NotEqual = @"\U+2260";

        /// <summary>
        /// 加上划线
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string AddOverline(string str)
        {
            return "%%o" + str + "%%o";
        }

        /// <summary>
        /// 加下划线
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string AddUnderline(string str)
        {
            return "%%u" + str + "%%u";
        }
    }
}
