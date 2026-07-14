using PowerPointConverter.Extension;
using PowerPointConverter.Model;
using System.Text.RegularExpressions;

namespace PowerPointConverter.Shapes
{
    /// <summary>
    /// Preset shape SVG path generators for OOXML preset geometry types.
    /// Each generator takes width, height, and optional adjustment values,
    /// returning an SVG path d-attribute string.
    /// Adjustment values follow OOXML convention: values are in 100000ths
    /// (so 50000 = 50%).
    /// </summary>
    public class PresetShape
    {
        // ---------------------------------------------------------------------------
        // Preset shape registry
        // ---------------------------------------------------------------------------
        public static Dictionary<string, Func<double, double, Dictionary<string, int>, string>> PresetShapes = new Dictionary<string, Func<double, double, Dictionary<string, int>, string>>();

        // Fallback rectangle for action buttons without multiPathPresets entry yet
        // actionButtonSound fallback removed — uses multiPathPresets entry below
        // Multi-path action button presets are registered after the multiPathPresets Map
        // declaration (see below in the multiPathPresets section).
        // ---------------------------------------------------------------------------
        // Action button icon paths (rendered as a second <path> with contrasting fill)
        // ---------------------------------------------------------------------------
        public static Dictionary<string, Func<double, double, string>> ActionButtonIcons = new Dictionary<string, Func<double, double, string>>();

        public static Dictionary<string, Func<double, double, Dictionary<string, int>, List<PresetOverlayInfo>>> PresetOverlays = new Dictionary<string, Func<double, double, Dictionary<string, int>, List<PresetOverlayInfo>>>();

        public static Dictionary<string, Func<double, double, Dictionary<string, int>, List<ArrowPathInfo>>> MultiPathPresets = new Dictionary<string, Func<double, double, Dictionary<string, int>, List<ArrowPathInfo>>>();

        static PresetShape()
        {
            InitPresetShapes();
            InitactionButtonIcons();
            InitPresetOverlays();
            InitMultiPathPresets();
        }

        public static double Adjust(Dictionary<string, int> adjustments, string name, int defaultVal)
        {
            var raw = adjustments != null && adjustments.ContainsKey(name) ? adjustments[name] : defaultVal;
            return raw / 100000.0d;
        }

        public static int AdjustRaw(Dictionary<string, int> adjustments, string name, int defaultVal)
        {
            return adjustments != null && adjustments.ContainsKey(name) ? adjustments[name] : defaultVal;
        }

        /// <summary>
        /// Helper: generate a star polygon.
        /// </summary>
        /// <param name="w"></param>
        /// <param name="h"></param>
        /// <param name="points"></param>
        /// <param name="innerRatio"></param>
        /// <returns></returns>
        public static string StarShape(double w, double h, double points, double innerRatio = 0.4d)
        {
            var cx = w / 2;
            var cy = h / 2;
            var outerRx = w / 2;
            var outerRy = h / 2;
            var innerRx = outerRx * innerRatio;
            var innerRy = outerRy * innerRatio;
            var totalPoints = points * 2;

            var parts = new List<string>();

            for (var i = 0; i < totalPoints; i++)
            {
                var angle = (2 * Math.PI * i) / totalPoints - Math.PI / 2;
                var isOuter = i % 2 == 0;
                var rx = isOuter ? outerRx : innerRx;
                var ry = isOuter ? outerRy : innerRy;
                var x = cx + rx * Math.Cos(angle);
                var y = cy + ry * Math.Sin(angle);

                parts.Add(i == 0 ? $"M{x},{y}" : $"L{x},{y}");
            }

            parts.Add("Z");

            return string.Join(" ", parts);
        }

        /// <summary>
        /// Mirror an absolute SVG path horizontally across the given width.
        /// Supports the command subset used by preset arrow shapes: M, L, A, Z.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="width"></param>
        /// <returns></returns>
        public static string MirrorAbsolutePathHorizontally(string path, double width)
        {
            var tokens = Regex.Match(path, @"[MLAZ]|-?\d*\.?\d+(?:e[-+]?\d+)?", RegexOptions.IgnoreCase);
            if (tokens.Success == false)
                return path;

            var _out = new List<string>();
            var i = 0;

            while (i < tokens.Length)
            {
                var cmd = tokens.Groups[i++];
                if (cmd == null || cmd.Value.Length == 0)
                    break;

                _out.Add(cmd.Value);

                if (cmd.Value == "Z")
                    continue;

                if (cmd.Value == "M" || cmd.Value == "L")
                {
                    var x = double.Parse(tokens.Groups[i++].Value);
                    var y = double.Parse(tokens.Groups[i++].Value);

                    _out.AddRange((width - x).ToString(), y.ToString());
                    continue;
                }

                if (cmd.Value == "A")
                {
                    var rx = tokens.Groups[i++].Value;
                    var ry = tokens.Groups[i++].Value;
                    var rot = tokens.Groups[i++].Value;
                    var largeArc = tokens.Groups[i++].Value;
                    var sweep = double.Parse(tokens.Groups[i++].Value);
                    var x = double.Parse(tokens.Groups[i++].Value);
                    var y = double.Parse(tokens.Groups[i++].Value);
                   
                    _out.AddRange(rx, ry, rot, largeArc, (sweep>0 ? 0 : 1).ToString(), (width - x).ToString(), y.ToString());

                    continue;
                }

                return path;
            }

            return string.Join(" ", _out);
        }

        public static string MirrorAbsolutePathVertically(string path, double height)
        {
            var tokens = Regex.Match(path, @"[MLAZ]|-?\d*\.?\d+(?:e[-+]?\d+)?", RegexOptions.IgnoreCase);
            if (tokens.Success == false)
                return path;

            var _out = new List<string>();

            var i = 0;
            while (i < tokens.Length)
            {
                var cmd = tokens.Groups[i++];

                if (cmd == null)
                    break;

                _out.Add(cmd.Value);

                if (cmd.Value == "Z")
                    continue;

                if (cmd.Value == "M" || cmd.Value == "L")
                {
                    var x = double.Parse(tokens.Groups[i++].Value);
                    var y = double.Parse(tokens.Groups[i++].Value);

                    _out.Add(x.ToString());
                    _out.Add((height - y).ToString());

                    continue;
                }

                if (cmd.Value == "A")
                {
                    var rx = tokens.Groups[i++].Value;
                    var ry = tokens.Groups[i++].Value;
                    var rot = tokens.Groups[i++].Value;
                    var largeArc = tokens.Groups[i++].Value;
                    var sweep = double.Parse(tokens.Groups[i++].Value);
                    var x = double.Parse(tokens.Groups[i++].Value);
                    var y = double.Parse(tokens.Groups[i++].Value);

                    _out.AddRange(rx, ry, rot, largeArc, (sweep > 0 ? 0 : 1).ToString(), x.ToString(), (height - y).ToString());
                }
            }

            return string.Join(" ", _out);
        }

        public static ContourInfo SplitFirstClosedContour(string path)
        {
            var closeIdx = path.IndexOf("Z");
            if (closeIdx == -1)
            {
                return new ContourInfo() { Outer = path, Remainder = "" };
            }

            var outer = path.Substring(0, closeIdx + 1).Trim();
            var remainder = path.Substring(closeIdx + 1).Trim();

            return new ContourInfo() { Outer = outer, Remainder = remainder };
        }

        public static List<ArrowPathInfo> BuildCurvedArrowMultiPath(string shapeName, double w, double h, Dictionary<string, int> adjustments)
        {
            var fullPath = PresetShapes[shapeName](w, h, adjustments);
            var outerInfo = SplitFirstClosedContour(fullPath);
            string outer = outerInfo.Outer;
            string remainder = outerInfo.Remainder;

            if (remainder == null)
            {
                return new List<ArrowPathInfo>() { new ArrowPathInfo() { D = fullPath, Fill = "norm", Stroke = true } };
            }

            if (shapeName == "curvedRightArrow")
            {
                return new List<ArrowPathInfo>() {
                       new ArrowPathInfo() { D = remainder, Fill= "norm", Stroke = true },
                       new ArrowPathInfo() { D= outer, Fill= "norm", Stroke= true }
                };
            }

            return new List<ArrowPathInfo>() {
                       new ArrowPathInfo() { D = outer, Fill= "norm", Stroke = true },
                       new ArrowPathInfo() { D= remainder, Fill= "norm", Stroke= true }
                };
        }

        public static List<ArrowPathInfo> buildCurvedVerticalArrowMultiPath(string shapeName, double w, double h, Dictionary<string, int> adjustments)
        {
            var downFullPath = PresetShapes["curvedDownArrow"](w, h, adjustments);
            var arrowPathInfo = SplitFirstClosedContour(downFullPath);
            string remainder = arrowPathInfo?.Remainder;
            string outer = arrowPathInfo?.Outer;

            var ordered = remainder != null
                ? new List<ArrowPathInfo>()
                {
                    new ArrowPathInfo()  { D= remainder, Fill= "norm", Stroke= true },
                    new ArrowPathInfo() {  D= outer, Fill= "norm", Stroke= true }
                }
                : new List<ArrowPathInfo>() { new ArrowPathInfo() { D = downFullPath, Fill = "norm", Stroke = true } };

            if (shapeName == "curvedDownArrow")
            {
                return ordered;
            }

            var mirrored = ordered.Select(path => new ArrowPathInfo()
            {
                D = MirrorAbsolutePathVertically(path.D, h),
                Fill = path.Fill,
                Stroke = path.Stroke
            }).Reverse().ToList();

            return mirrored;
        }

        public static string BuildCircularArrowPath(double w, double h, Dictionary<string, int> adjustments, bool _mirrorX = false, string variant = "circularArrow")
        {
            // OOXML circularArrow / leftCircularArrow: same guide formulas, different default adjustments.
            var hc = w / 2;
            var vc = h / 2;
            var wd2 = w / 2;
            var hd2 = h / 2;
            var ss = Math.Min(w, h);
            var cd2 = 10800000; // 180° in 60000ths
            Func<double, double> toRad60k = (a) => ((a / 60000) * Math.PI) / 180;
            // OOXML formula helpers
            Func<double, double, double> ooxSin = (val, ang) => val * Math.Sin(toRad60k(ang));
            Func<double, double, double> ooxCos = (val, ang) => val * Math.Cos(toRad60k(ang));
            Func<double, double, double, double> cat2 = (r, ht, wt) => r * Math.Cos(Math.Atan2(wt, ht));
            Func<double, double, double, double> sat2 = (r, ht, wt) => r * Math.Sin(Math.Atan2(wt, ht));
            // OOXML: at2(x, y) = atan2(y, x) — first arg is x, second is y
            Func<double, double, double> at2 = (x, y) => ((Math.Atan2(y, x) * 180) / Math.PI) * 60000;
            Func<double, double, double, double> modF = (x, y, z) => Math.Sqrt(x * x + y * y + z * z);

            // Adjustments — leftCircularArrow has different OOXML defaults
            var isLeft = variant == "leftCircularArrow";
            var adj1 = adjustments?["adj1"] ?? 12500;
            var adj2 = adjustments?["adj2"] ?? (isLeft ? -1142319 : 1142319);
            var adj3 = adjustments?["adj3"] ?? (isLeft ? 1142319 : 20457681);
            var adj4 = adjustments?["adj4"] ?? 10800000;
            var adj5v = adjustments?["adj5"] ?? 12500;
            var a5 = Math.Max(0, Math.Min(adj5v, 25000));
            var maxAdj1 = a5 * 2;
            var a1 = Math.Max(0, Math.Min(adj1, maxAdj1));
            var enAng = Math.Max(1, Math.Min(adj3, 21599999));
            var stAng = Math.Max(0, Math.Min(adj4, 21599999));
            var th = (ss * a1) / 100000.0;
            var thh = (ss * a5) / 100000.0;
            var th2 = th / 2;
            var rw1 = wd2 + th2 - thh;
            var rh1 = hd2 + th2 - thh;
            var rw2 = rw1 - th;
            var rh2 = rh1 - th;
            var rw3 = rw2 + th2;
            var rh3 = rh2 + th2;

            // Point H (mid-radius at end angle)
            var wtH = ooxSin(rw3, enAng);
            var htH = ooxCos(rh3, enAng);
            var dxH = cat2(rw3, htH, wtH);
            var dyH = sat2(rh3, htH, wtH);
            var xH = hc + dxH;
            var yH = vc + dyH;

            // Compute max arrowhead angle
            var rI = Math.Min(rw2, rh2);
            var u1 = dxH * dxH;
            var u2 = dyH * dyH;
            var u3 = rI * rI;
            var u4 = u1 - u3;
            var u5 = u2 - u3;
            var u6 = u2 != 0 ? (u4 * u5) / u1 : 0;
            var u7 = u2 != 0 ? u6 / u2 : 0;
            var u8 = 1 - u7;
            var u9 = Math.Sqrt(Math.Max(0, u8));
            var u10 = dxH != 0 ? u4 / dxH : 0;
            var u11 = dyH != 0 ? u10 / dyH : 0;
            var u12 = u11 != 0 ? (1 + u9) / u11 : 0;
            var u13 = at2(1, u12);
            var u14 = u13 + 21600000;
            var u15 = u13 >= 0 ? u13 : u14;
            var u16 = u15 - enAng;
            var u17 = u16 + 21600000;
            var u18 = u16 >= 0 ? u16 : u17;
            var u19 = u18 - cd2;
            var u20 = u18 - 21600000;
            var u21 = u19 >= 0 ? u20 : u18;
            var maxAng = Math.Abs(u21);
            double aAng;

            if (isLeft)
            {
                // leftCircularArrow: minAng = -abs(u21), a2 = -abs(adj2), aAng = pin(minAng, a2, 0)
                var minAng = -maxAng;
                var a2 = -Math.Abs(adj2);
                aAng = Math.Max(minAng, Math.Min(a2, 0));
            }
            else
            {
                aAng = Math.Max(0, Math.Min(adj2, maxAng));
            }

            var ptAng = enAng + aAng;
            // Point A (arrowhead tip)
            var wtA = ooxSin(rw3, ptAng);
            var htA = ooxCos(rh3, ptAng);
            var dxA = cat2(rw3, htA, wtA);
            var dyA = sat2(rh3, htA, wtA);
            var xA = hc + dxA;
            var yA = vc + dyA;

            // Point E (outer arc start)
            var wtE = ooxSin(rw1, stAng);
            var htE = ooxCos(rh1, stAng);
            var dxE = cat2(rw1, htE, wtE);
            var dyE = sat2(rh1, htE, wtE);
            var xE = hc + dxE;
            var yE = vc + dyE;

            // Points G and B (arrowhead base, offset from H by thh at angle ptAng)
            var dxG = ooxCos(thh, ptAng);
            var dyG = ooxSin(thh, ptAng);
            var xG = xH + dxG;
            var yG = yH + dyG;
            var xB = xH - dxG;
            var yB = yH - dyG;

            // Scale to normalized circle for line-circle intersection
            var sx1 = xB - hc;
            var sy1 = yB - vc;
            var sx2 = xG - hc;
            var sy2 = yG - vc;

            // Outer circle intersection
            var rO = Math.Min(rw1, rh1);
            var x1O = rw1 != 0 ? (sx1 * rO) / rw1 : 0;
            var y1O = rh1 != 0 ? (sy1 * rO) / rh1 : 0;
            var x2O = rw1 != 0 ? (sx2 * rO) / rw1 : 0;
            var y2O = rh1 != 0 ? (sy2 * rO) / rh1 : 0;
            var dxO = x2O - x1O;
            var dyO = y2O - y1O;
            var dOval = modF(dxO, dyO, 0);
            var q1 = x1O * y2O;
            var q2 = x2O * y1O;
            var DO = q1 - q2;
            var q3 = rO * rO;
            var q4 = dOval * dOval;
            var q5 = q3 * q4;
            var q6 = DO * DO;
            var q7 = q5 - q6;
            var q8 = Math.Max(q7, 0);
            var sdelO = Math.Sqrt(q8);
            var ndyO = dyO * -1;
            var sdyO = ndyO >= 0 ? -1 : 1;
            var q9 = sdyO * dxO;
            var q10 = q9 * sdelO;
            var q11 = DO * dyO;
            var dxF1 = q4 != 0 ? (q11 + q10) / q4 : 0;
            var q12 = q11 - q10;
            var dxF2 = q4 != 0 ? q12 / q4 : 0;
            var adyO = Math.Abs(dyO);
            var q13 = adyO * sdelO;
            var q14 = DO * dxO * -1;
            var dyF1 = q4 != 0 ? (q14 + q13) / q4 : 0;
            var q15 = q14 - q13;
            var dyF2 = q4 != 0 ? q15 / q4 : 0;

            // Pick intersection closest to G side
            var q16 = x2O - dxF1;
            var q17 = x2O - dxF2;
            var q18 = y2O - dyF1;
            var q19 = y2O - dyF2;
            var q20 = modF(q16, q18, 0);
            var q21 = modF(q17, q19, 0);
            var q22 = q21 - q20;
            var dxF = q22 >= 0 ? dxF1 : dxF2;
            var dyF = q22 >= 0 ? dyF1 : dyF2;
            var sdxF = rO != 0 ? (dxF * rw1) / rO : 0;
            var sdyF = rO != 0 ? (dyF * rh1) / rO : 0;
            var xF = hc + sdxF;
            var yF = vc + sdyF;

            // Inner circle intersection
            var x1I = rw2 != 0 ? (sx1 * rI) / rw2 : 0;
            var y1I = rh2 != 0 ? (sy1 * rI) / rh2 : 0;
            var x2I = rw2 != 0 ? (sx2 * rI) / rw2 : 0;
            var y2I = rh2 != 0 ? (sy2 * rI) / rh2 : 0;
            var dxI = x2I - x1I;
            var dyI = y2I - y1I;
            var dI = modF(dxI, dyI, 0);
            var v1 = x1I * y2I;
            var v2 = x2I * y1I;
            var DI = v1 - v2;
            var v3 = rI * rI;
            var v4 = dI * dI;
            var v5 = v3 * v4;
            var v6 = DI * DI;
            var v7 = v5 - v6;
            var v8 = Math.Max(v7, 0);
            var sdelI = Math.Sqrt(v8);
            var v9 = sdyO * dxI;
            var v10 = v9 * sdelI;
            var v11 = DI * dyI;
            var dxC1 = v4 != 0 ? (v11 + v10) / v4 : 0;
            var v12 = v11 - v10;
            var dxC2 = v4 != 0 ? v12 / v4 : 0;
            var adyI = Math.Abs(dyI);
            var v13 = adyI * sdelI;
            var v14 = DI * dxI * -1;
            var dyC1 = v4 != 0 ? (v14 + v13) / v4 : 0;
            var v15 = v14 - v13;
            var dyC2 = v4 != 0 ? v15 / v4 : 0;

            // Pick intersection closest to B side (x1I)
            var v16 = x1I - dxC1;
            var v17 = x1I - dxC2;
            var v18 = y1I - dyC1;
            var v19 = y1I - dyC2;
            var v20 = modF(v16, v18, 0);
            var v21 = modF(v17, v19, 0);
            var v22 = v21 - v20;
            var dxC = v22 >= 0 ? dxC1 : dxC2;
            var dyC = v22 >= 0 ? dyC1 : dyC2;
            var sdxC = rI != 0 ? (dxC * rw2) / rI : 0;
            var sdyC = rI != 0 ? (dyC * rh2) / rI : 0;
            var xC = hc + sdxC;
            var yC = vc + sdyC;

            // Inner arc angles — leftCircularArrow uses intermediate istAng0/iswAng0
            var ist0 = at2(sdxC, sdyC);
            var ist1 = ist0 + 21600000;
            var istAng0 = ist0 >= 0 ? ist0 : ist1;
            var isw1 = stAng - istAng0;
            double istAng;
            double iswAng;

            if (isLeft)
            {
                // leftCircularArrow: iswAng0 always ≥ 0, then istAng shifted, iswAng negated
                var iswAng0 = isw1 >= 0 ? isw1 : isw1 + 21600000;
                istAng = istAng0 + iswAng0;
                iswAng = -iswAng0;
            }
            else
            {
                // circularArrow: iswAng always ≤ 0 (clockwise inner arc)
                istAng = istAng0;
                iswAng = isw1 >= 0 ? isw1 - 21600000 : isw1;
            }

            // Adjusted arrowhead points (clamp when too close)
            var p1 = xF - xC;
            var p2 = yF - yC;
            var p3 = modF(p1, p2, 0);
            var p4 = p3 / 2;
            var p5 = p4 - thh;
            var xGp = p5 >= 0 ? xF : xG;
            var yGp = p5 >= 0 ? yF : yG;
            var xBp = p5 >= 0 ? xC : xB;
            var yBp = p5 >= 0 ? yC : yB;
            // Outer arc sweep angle
            var en0 = at2(sdxF, sdyF);
            var en1 = en0 + 21600000;
            var en2 = en0 >= 0 ? en0 : en1;
            var sw0 = en2 - stAng;
            double outerArcStAng;
            double outerArcSwAng;

            if (isLeft)
            {
                // leftCircularArrow: swAng ≤ 0, then stAng0 = stAng + swAng, swAng0 = -swAng
                var swAngRaw = sw0 >= 0 ? sw0 - 21600000 : sw0;
                outerArcStAng = stAng + swAngRaw; // stAng0
                outerArcSwAng = -swAngRaw; // swAng0 (positive)
            }
            else
            {
                var swAng = sw0 >= 0 ? sw0 : sw0 + 21600000;
                outerArcStAng = stAng;
                outerArcSwAng = swAng;
            }

            // Compute end points for SVG arcs using OOXML arcTo semantics
            // Outer arc: from outerArcStAng sweeping outerArcSwAng
            var outerEndAng = outerArcStAng + outerArcSwAng;
            var wtOE = ooxSin(rw1, outerEndAng);
            var htOE = ooxCos(rh1, outerEndAng);
            var xOE = hc + cat2(rw1, htOE, wtOE);
            var yOE = vc + sat2(rh1, htOE, wtOE);

            // Inner arc: from istAng sweeping iswAng
            var innerEndAng = istAng + iswAng;
            var wtIE = ooxSin(rw2, innerEndAng);
            var htIE = ooxCos(rh2, innerEndAng);
            var xIE = hc + cat2(rw2, htIE, wtIE);
            var yIE = vc + sat2(rh2, htIE, wtIE);

            // SVG arc flags
            var outerSweepDeg = Math.Abs(outerArcSwAng / 60000);
            var outerLargeArc = outerSweepDeg > 180 ? 1 : 0;
            var outerSweepFlag = outerArcSwAng > 0 ? 1 : 0;
            var innerSweepDeg = Math.Abs(iswAng / 60000);
            var innerLargeArc = innerSweepDeg > 180 ? 1 : 0;
            var innerSweepFlag = iswAng > 0 ? 1 : 0;

            if (isLeft)
            {
                // leftCircularArrow path: M(xE) → L(xD) → inner arc → arrowhead → L(xF) → outer arc → Z
                // Point D: inner arc start at stAng on rw2/rh2
                var wtD = ooxSin(rw2, stAng);
                var htD = ooxCos(rh2, stAng);
                var xD = hc + cat2(rw2, htD, wtD);
                var yD = vc + sat2(rh2, htD, wtD);
                return string.Join(" ", [
                    $"M{xE},{yE}",
                    $"L{xD},{yD}",
                    $"A{rw2},{rh2} 0 {innerLargeArc},{innerSweepFlag} {xIE},{yIE}",
                    $"L{xBp},{yBp}",
                    $"L{xA},{yA}",
                    $"L{xGp},{yGp}",
                    $"L{xF},{yF}",
                    $"A{rw1},{rh1} 0 {outerLargeArc},{outerSweepFlag} {xOE},{yOE}",
                    "Z"]);
            }

            return string.Join(" ", [
                $"M{xE},{yE}",
                $"A{rw1},{rh1} 0 {outerLargeArc},{outerSweepFlag} {xOE},{yOE}",
                $"L{xGp},{yGp}",
                $"L{xA},{yA}",
                $"L{xBp},{yBp}",
                $"L{xC},{yC}",
                $"A{rw2},{rh2} 0 {innerLargeArc},{innerSweepFlag} {xIE},{yIE}",
                "Z"]);
        }

        private static void InitPresetShapes()
        {
            // ==== Basic Shapes ====
            PresetShapes.Add("rect", (w, h, a) => $"M0,0 L{w},0 L{w},{h} L0,{h} Z");

            PresetShapes.Add("roundRect", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 16667);
                var r = Math.Min(w, h) * a;

                return string.Join(" ", [
                    $"M{r},0",
                    $"L{w - r},0",
                    $"A{r},{r} 0 0,1 {w},{r}",
                    $"L{w},{h - r}",
                    $"A{r},{r} 0 0,1 {w - r},{h}",
                    $"L{r},{h}",
                    $"A{r},{r} 0 0,1 0,{h - r}",
                    $"L0,{r}",
                    $"A{r},{r} 0 0,1 {r},0",
                    "Z",
                        ]);
            });

            PresetShapes.Add("plaque", (w, h, adjustments) =>
            {
                // OOXML: adj default 16667, concave (inward) arc corners via negative sweep arcTo
                var a = Math.Min(Math.Max(AdjustRaw(adjustments, "adj", 16667), 0), 50000);
                var x1 = (Math.Min(w, h) * a) / 100000.0;
                var x2 = w - x1;
                var y2 = h - x1;

                // Start at (0, x1), arcTo with negative sweep creates concave corner
                var a1 = OoArcTo(0, x1, x1, x1, 90, -90); // top-left: ends at (x1, 0)
                var a2 = OoArcTo(x2, 0, x1, x1, 180, -90); // top-right: ends at (w, x1)
                var a3 = OoArcTo(w, y2, x1, x1, 270, -90); // bottom-right: ends at (x2, h)
                var a4 = OoArcTo(x1, h, x1, x1, 0, -90); // bottom-left: ends at (0, y2) -> close to (0, x1)

                return string.Join(" ", [
                    $"M0,{x1}",
                    a1.SVG,
                    $"L{x2},0",
                    a2.SVG,
                    $"L{w},{y2}",
                    a3.SVG,
                    $"L{x1},{h}",
                    a4.SVG,
                    "Z",
                ]);
            });

            // Tab family: OOXML uses dx = sqrt(w²+h²)/20 (diagonal/20)
            PresetShapes.Add("cornerTabs", (w, h, a) =>
            {
                var dx = Math.Sqrt(w * w + h * h) / 20.0d;
                return string.Join(" ", [
                    $"M0,0 L{dx},0 L0,{dx} Z",
                    $"M{w},0 L{w - dx},0 L{w},{dx} Z",
                    $"M{w},{h} L{w - dx},{h} L{w},{h - dx} Z",
                    $"M0,{h} L{dx},{h} L0,{h - dx} Z",
                ]);
            });

            PresetShapes.Add("squareTabs", (w, h, a) =>
            {
                var dx = Math.Sqrt(w * w + h * h) / 20.0d;
                return string.Join(" ", [
                    $"M0,0 L{dx},0 L{dx},{dx} L0,{dx} Z",
                    $"M{w - dx},0 L{w},0 L{w},{dx} L{w - dx},{dx} Z",
                    $"M0,{h - dx} L{dx},{h - dx} L{dx},{h} L0,{h} Z",
                    $"M{w - dx},{h - dx} L{w},{h - dx} L{w},{h} L{w - dx},{h} Z",
                ]);
            });
            PresetShapes.Add("plaqueTabs", (w, h, a) =>
            {
                var dx = Math.Sqrt(w * w + h * h) / 20.0d;
                return string.Join(" ", [
                    $"M0,0 L{dx},0 A{dx},{dx} 0 0,1 0,{dx} Z",
                    $"M{w},0 L{w - dx},0 A{dx},{dx} 0 0,0 {w},{dx} Z",
                    $"M0,{h} L0,{h - dx} A{dx},{dx} 0 0,1 {dx},{h} Z",
                    $"M{w},{h} L{w - dx},{h} A{dx},{dx} 0 0,1 {w},{h - dx} Z",
                ]);
            });

            PresetShapes.Add("ellipse", (w, h, a) =>
            {
                var rx = w / 2.0d;
                var ry = h / 2.0d;

                return string.Join(" ", [$"M{w},{ry}", $"A{rx},{ry} 0 1,1 0,{ry}", $"A{rx},{ry} 0 1,1 {w},{ry}", "Z"]);
            });

            PresetShapes.Add("triangle", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 50000);
                var topX = w * a;

                return $"M{topX},0 L{w},{h} L0,{h} Z";
            });

            PresetShapes.Add("isosTriangle", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 50000);
                var topX = w * a;

                return $"M{topX},0 L{w},{h} L0,{h} Z";
            });

            PresetShapes.Add("rtTriangle", (w, h, a) => $"M0,0 L{w},{h} L0,{h} Z");
            PresetShapes.Add("diamond", (w, h, a) =>
            {
                var cx = w / 2;
                var cy = h / 2;
                return $"M{cx},0 L{w},{cy} L{cx},{h} L0,{cy} Z";
            });

            PresetShapes.Add("pentagon", (w, h, a) =>
            {
                // OOXML pentagon: hf=105146, vf=110557 with center shifted to svc so top vertex = y=0.
                var hc = w / 2.0d;
                var swd2 = (hc * 105146) / 100000.0d;
                var shd2 = ((h / 2) * 110557) / 100000.0d;
                var svc = shd2; // svc = vc * vf/100000 = shd2, so top vertex at svc - shd2 = 0
                var dx1 = swd2 * Math.Cos((18 * Math.PI) / 180); // cos 1080000
                var dx2 = swd2 * Math.Cos((54 * Math.PI) / 180); // cos 18360000
                var dy1 = shd2 * Math.Sin((18 * Math.PI) / 180); // sin 1080000
                var dy2 = shd2 * Math.Sin((54 * Math.PI) / 180); // |sin 18360000|

                return string.Join(" ", [
                    $"M{hc - dx1},{svc - dy1}", // x1, y1 (upper-left)
                    $"L{hc},0", // hc, t (top)
                    $"L{hc + dx1},{svc - dy1}", // x4, y1 (upper-right)
                    $"L{hc + dx2},{svc + dy2}", // x3, y2 (lower-right)
                    $"L{hc - dx2},{svc + dy2}", // x2, y2 (lower-left)
                    $"Z"
                ]);
            });

            PresetShapes.Add("hexagon", (w, h, adjustments) =>
            {
                // OOXML hexagon: adj=25000, vf=115470 (2/√3 scale factor for regular hex).
                var ss = Math.Min(w, h);
                var a = Math.Min(Math.Max(AdjustRaw(adjustments, "adj", 25000), 0), ss > 0 ? (50000 * w) / ss : 50000);
                var vf = 115470;
                var shd2 = ((h / 2) * vf) / 100000.0d;
                var x1 = (ss * a) / 100000.0d;
                var x2 = w - x1;
                var _hc = w / 2;
                var vc = h / 2;
                // dy1 = sin(shd2, 60°) = shd2 * sin(60°)
                var dy1 = shd2 * Math.Sin((60 * Math.PI) / 180);
                var y1 = vc - dy1;
                var y2 = vc + dy1;

                return string.Join(" ", [
                    $"M0,{vc}",
                    $"L{x1},{y1}",
                    $"L{x2},{y1}",
                    $"L{w},{vc}",
                    $"L{x2},{y2}",
                    $"L{x1},{y2}",
                    "Z"
                ]);
            });

            PresetShapes.Add("octagon", (w, h, adjustments) =>
            {
                // OOXML octagon: adj=29289 (≈1-1/√2). Uses ss-based cuts for both x and y.
                var ss = Math.Min(w, h);
                var a = Math.Min(Math.Max(AdjustRaw(adjustments, "adj", 29289), 0), 50000);
                var x1 = (ss * a) / 100000.0;
                var x2 = w - x1;
                var y2 = h - x1;

                return string.Join(" ", [
                    $"M0,{x1}",
                    $"L{x1},0",
                    $"L{x2},0",
                    $"L{w},{x1}",
                    $"L{w},{y2}",
                    $"L{x2},{h}",
                    $"L{x1},{h}",
                    $"L0,{y2}",
                    "Z"]);
            });

            PresetShapes.Add("heptagon", (w, h, a) =>
            {
                // OOXML heptagon: hf=102572, vf=105210 with shifted center.
                var hc = w / 2;
                var swd2 = (hc * 102572) / 100000.0;
                var shd2 = ((h / 2) * 105210) / 100000.0;
                var svc = ((h / 2) * 105210) / 100000.0;
                // Pre-computed trig ratios from OOXML spec (scaled by 100000)
                var dx1 = (swd2 * 97493) / 100000.0; // cos(12.857°) ≈ sin(77.14°)
                var dx2 = (swd2 * 78183) / 100000.0; // cos(38.57°)
                var dx3 = (swd2 * 43388) / 100000.0; // cos(64.29°)
                var dy1 = (shd2 * 62349) / 100000.0; // sin(38.57°)
                var dy2 = (shd2 * 22252) / 100000.0; // sin(12.857°)
                var dy3 = (shd2 * 90097) / 100000.0; // sin(64.29°)

                return string.Join(" ", [
                    $"M{hc - dx1},{svc + dy2}", // x1, y2 (left)
                    $"L{hc - dx2},{svc - dy1}", // x2, y1 (upper-left)
                    $"L{hc},0", // hc, t (top: svc - shd2 = 0)
                    $"L{hc + dx2},{svc - dy1}", // x5, y1 (upper-right)
                    $"L{hc + dx1},{svc + dy2}", // x6, y2 (right)
                    $"L{hc + dx3},{svc + dy3}", // x4, y3 (lower-right)
                    $"L{hc - dx3},{svc + dy3}", // x3, y3 (lower-left)
                    "Z"]);
            });

            PresetShapes.Add("decagon", (w, h, a) =>
            {
                // OOXML decagon: vf=105146 (no hf, uses wd2 for x). 10 vertices starting from left.
                var hc = w / 2;
                var vc = h / 2;
                var shd2 = (vc * 105146) / 100000.0d;
                // OOXML angles: 2160000=36°, 4320000=72°
                var dx1 = hc * Math.Cos((36 * Math.PI) / 180); // cos(wd2, 2160000)
                var dx2 = hc * Math.Cos((72 * Math.PI) / 180); // cos(wd2, 4320000)
                var dy1 = shd2 * Math.Sin((72 * Math.PI) / 180); // sin(shd2, 4320000)
                var dy2 = shd2 * Math.Sin((36 * Math.PI) / 180); // sin(shd2, 2160000)

                return string.Join(" ", [
                    $"M0,{vc}", // l, vc
                    $"L{hc - dx1},{vc - dy2}", // x1, y2
                    $"L{hc - dx2},{vc - dy1}", // x2, y1
                    $"L{hc + dx2},{vc - dy1}", // x3, y1
                    $"L{hc + dx1},{vc - dy2}", // x4, y2
                    $"L{w},{vc}", // r, vc
                    $"L{hc + dx1},{vc + dy2}", // x4, y3
                    $"L{hc + dx2},{vc + dy1}", // x3, y4
                    $"L{hc - dx2},{vc + dy1}", // x2, y4
                    $"L{hc - dx1},{vc + dy2}", // x1, y3
                    "Z"]);
            });

            PresetShapes.Add("dodecagon", (w, h, a) =>
            {
                // OOXML dodecagon: 21600-unit coordinate space, simple ratios.
                var x1 = (w * 2894) / 21600.0;
                var x2 = (w * 7906) / 21600.0;
                var x3 = (w * 13694) / 21600.0;
                var x4 = (w * 18706) / 21600.0;
                var y1 = (h * 2894) / 21600.0;
                var y2 = (h * 7906) / 21600.0;
                var y3 = (h * 13694) / 21600.0;
                var y4 = (h * 18706) / 21600.0;

                return string.Join(" ", [
                    $"M0,{y2}",
                    $"L{x1},{y1}",
                    $"L{x2},0",
                    $"L{x3},0",
                    $"L{x4},{y1}",
                    $"L{w},{y2}",
                    $"L{w},{y3}",
                    $"L{x4},{y4}",
                    $"L{x3},{h}",
                    $"L{x2},{h}",
                    $"L{x1},{y4}",
                    $"L0,{y3}",
                    "Z"]);
            });

            PresetShapes.Add("parallelogram", (w, h, adjustments) =>
            {
                // OOXML: adj=25000, x2 = ss * a / 100000, path: M(l,b)→L(x2,t)→L(r,t)→L(r-x2,b)→Z
                var ss = Math.Min(w, h);
                var maxAdj = ss > 0 ? (100000 * w) / ss : 100000.0;
                var a = Math.Min(Math.Max(AdjustRaw(adjustments, "adj", 25000), 0), maxAdj);
                var x2 = (ss * a) / 100000.0;
                var x5 = w - x2;

                return $"M0,{h} L{x2},0 L{w},0 L{x5},{h} Z";
            });

            PresetShapes.Add("trapezoid", (w, h, adjustments) =>
            {
                // OOXML: adj=25000, x2 = ss * a / 100000, x3 = r - x2
                var ss = Math.Min(w, h);
                var maxAdj = ss > 0 ? (50000 * w) / ss : 50000;
                var a = Math.Min(Math.Max(AdjustRaw(adjustments, "adj", 25000), 0), maxAdj);
                var x2 = (ss * a) / 100000.0;
                var x3 = w - x2;

                return $"M0,{h} L{x2},0 L{x3},0 L{w},{h} Z";
            });

            PresetShapes.Add("nonIsoscelesTrapezoid", (w, h, adjustments) =>
            {
                // OOXML: Two independent top insets. adj1=25000, adj2=25000
                var ss = Math.Min(w, h);
                var maxAdj = ss > 0 ? (50000 * w) / ss : 50000;
                var a1 = Math.Min(Math.Max(AdjustRaw(adjustments, "adj1", 25000), 0), maxAdj);
                var a2 = Math.Min(Math.Max(AdjustRaw(adjustments, "adj2", 25000), 0), maxAdj);
                var x2 = (ss * a1) / 100000.0;
                var dx3 = (ss * a2) / 100000.0;
                var x3 = w - dx3;

                return $"M0,{h} L{x2},0 L{x3},0 L{w},{h} Z";
            });

            PresetShapes.Add("corner", (w, h, adjustments) =>
            {
                // OOXML corner: two adjustments control horizontal and vertical arm thickness.
                // adj1 (default 50000) → vertical arm height from bottom: dy1 = ss * a1, y1 = h - dy1
                // adj2 (default 50000) → horizontal arm width from left: x1 = ss * a2
                var ss = Math.Min(w, h);
                var a1 = Math.Min(Math.Max(Adjust(adjustments, "adj1", 50000), 0), 1);
                var a2 = Math.Min(Math.Max(Adjust(adjustments, "adj2", 50000), 0), 1);
                var x1 = ss * a2;
                var dy1 = ss * a1;
                var y1 = h - dy1;

                return string.Join(" ", ["M0,0", $"L{x1},0", $"L{x1},{y1}", $"L{w},{y1}", $"L{w},{h}", $"L0,{h}", "Z"]);
            });

            PresetShapes.Add("diagStripe", (w, h, adjustments) =>
            {
                var a = Math.Min(Math.Max(Adjust(adjustments, "adj", 50000), 0), 1);
                var x2 = w * a;
                var y2 = h * a;

                return string.Join(" ", [$"M0,{y2}", $"L{x2},0", $"L{w},0", $"L0,{h}", "Z"]);
            });

            // ==== Star Shapes ====
            PresetShapes.Add("star4", (w, h, adjustments) =>
            {
                // OOXML default adj=12500 → innerRatio = 12500/50000 = 0.25
                var a = Adjust(adjustments, "adj", 12500) * 2;

                return StarShape(w, h, 4, Math.Min(Math.Max(a, 0), 1));
            });

            PresetShapes.Add("star5", (w, h, adjustments) =>
            {
                // OOXML: adj=19098, hf=105146, vf=110557 — scaling factors for non-square bounding box
                var aRaw = adjustments?["adj"] ?? 19098;
                var a = Math.Min(Math.Max(aRaw, 0), 50000);
                var hf = 105146;
                var vf = 110557;
                var swd2 = ((w / 2) * hf) / 100000.0;
                var shd2 = ((h / 2) * vf) / 100000.0;
                var svc = ((h / 2) * vf) / 100000.0;
                var iwd2 = (swd2 * a) / 50000.0;
                var ihd2 = (shd2 * a) / 50000.0;
                var cx = w / 2;
                var step = (2 * Math.PI) / 5;
                var halfStep = step / 2;
                var startAngle = -Math.PI / 2;
                var parts = new List<string>();

                for (var i = 0; i < 5; i++)
                {
                    var outerAngle = startAngle + step * i;
                    var innerAngle = outerAngle + halfStep;
                    var ox = cx + swd2 * Math.Cos(outerAngle);
                    var oy = svc + shd2 * Math.Sin(outerAngle);
                    var ix = cx + iwd2 * Math.Cos(innerAngle);
                    var iy = svc + ihd2 * Math.Sin(innerAngle);

                    parts.Add(i == 0 ? $"M{ox},{oy}" : $"L{ox},{oy}");
                    parts.Add($"L{ix},{iy}");
                }

                parts.Add("Z");
                return string.Join(" ", parts);
            });

            PresetShapes.Add("star6", (w, h, adjustments) =>
            {
                // OOXML: adj=28868, hf=115470 — horizontal scaling factor
                var aRaw = adjustments?["adj"] ?? 28868;
                var a = Math.Min(Math.Max(aRaw, 0), 50000);
                var hf = 115470;
                var swd2 = ((w / 2) * hf) / 100000.0;
                var shd2 = h / 2; // no vf for star6
                var iwd2 = (swd2 * a) / 50000.0;
                var ihd2 = (shd2 * a) / 50000.0;
                var cx = w / 2;
                var cy = h / 2;
                var step = (2 * Math.PI) / 6;
                var halfStep = step / 2;
                var startAngle = -Math.PI / 2;
                var parts = new List<string>();

                for (var i = 0; i < 6; i++)
                {
                    var outerAngle = startAngle + step * i;
                    var innerAngle = outerAngle + halfStep;
                    var ox = cx + swd2 * Math.Cos(outerAngle);
                    var oy = cy + shd2 * Math.Sin(outerAngle);
                    var ix = cx + iwd2 * Math.Cos(innerAngle);
                    var iy = cy + ihd2 * Math.Sin(innerAngle);
                    parts.Add(i == 0 ? $"M{ox},{oy}" : $"L{ox},{oy}");
                    parts.Add($"L{ix},{iy}");
                }

                parts.Add("Z");
                return string.Join(" ", parts);
            });

            PresetShapes.Add("star7", (w, h, adjustments) =>
            {
                // OOXML star7: adj=34601, hf=102572, vf=105210 — center shifted to svc
                var aRaw = adjustments?["adj"] ?? 34601;
                var a = Math.Min(Math.Max(aRaw, 0), 50000);
                var swd2 = ((w / 2) * 102572) / 100000.0;
                var shd2 = ((h / 2) * 105210) / 100000.0;
                var svc = shd2; // = vc * vf/100000 so top vertex at svc - shd2 = 0
                var iwd2 = (swd2 * a) / 50000.0;
                var ihd2 = (shd2 * a) / 50000.0;
                var cx = w / 2;
                var step = (2 * Math.PI) / 7.0;
                var halfStep = step / 2.0;
                var startAngle = -Math.PI / 2;
                var parts = new List<string>();

                for (var i = 0; i < 7; i++)
                {
                    var outerAngle = startAngle + step * i;
                    var innerAngle = outerAngle + halfStep;
                    var ox = cx + swd2 * Math.Cos(outerAngle);
                    var oy = svc + shd2 * Math.Sin(outerAngle);
                    var ix = cx + iwd2 * Math.Cos(innerAngle);
                    var iy = svc + ihd2 * Math.Sin(innerAngle);
                    parts.Add(i == 0 ? $"M{ox},{oy}" : $"L{ox},{oy}");
                    parts.Add($"L{ix},{iy}");
                }

                parts.Add("Z");
                return string.Join(" ", parts);
            });

            PresetShapes.Add("star8", (w, h, adjustments) =>
            {
                // OOXML: iwd2 = wd2 * adj / 50000. adj default=37500 → innerRatio = 37500/50000 = 0.75
                // Adjust() divides by 100000, so we multiply by 2 to get adj/50000.
                var a = Adjust(adjustments, "adj", 37500) * 2;
                return StarShape(w, h, 8, Math.Min(Math.Max(a, 0), 1));
            });

            PresetShapes.Add("star10", (w, h, adjustments) =>
            {
                // OOXML: adj=42533, hf=105146 — horizontal scaling factor
                var aRaw = adjustments?["adj"] ?? 42533;
                var a = Math.Min(Math.Max(aRaw, 0), 50000);
                var hf = 105146;
                var swd2 = ((w / 2) * hf) / 100000.0;
                var shd2 = h / 2; // no vf for star10
                var iwd2 = (swd2 * a) / 50000.0;
                var ihd2 = (shd2 * a) / 50000.0;
                var cx = w / 2;
                var cy = h / 2;
                var step = (2 * Math.PI) / 10;
                var halfStep = step / 2;
                var startAngle = -Math.PI / 2;
                var parts = new List<string>();

                for (var i = 0; i < 10; i++)
                {
                    var outerAngle = startAngle + step * i;
                    var innerAngle = outerAngle + halfStep;
                    var ox = cx + swd2 * Math.Cos(outerAngle);
                    var oy = cy + shd2 * Math.Sin(outerAngle);
                    var ix = cx + iwd2 * Math.Cos(innerAngle);
                    var iy = cy + ihd2 * Math.Sin(innerAngle);

                    parts.Add(i == 0 ? $"M{ox},{oy}" : $"L{ox},{oy}");
                    parts.Add($"L{ix},{iy}");
                }
                parts.Add("Z");
                return string.Join(" ", parts);
            });

            PresetShapes.Add("star12", (w, h, adjustments) =>
            {
                // OOXML default adj=37500 → innerRatio = 0.75
                var a = Adjust(adjustments, "adj", 37500) * 2;
                return StarShape(w, h, 12, Math.Min(Math.Max(a, 0), 1));
            });

            PresetShapes.Add("star16", (w, h, adjustments) =>
            {
                // OOXML default adj=37500 → innerRatio = 0.75
                var a = Adjust(adjustments, "adj", 37500) * 2;
                return StarShape(w, h, 16, Math.Min(Math.Max(a, 0), 1));
            });

            PresetShapes.Add("star24", (w, h, adjustments) =>
            {
                // OOXML default adj=37500 → innerRatio = 0.75
                var a = Adjust(adjustments, "adj", 37500) * 2;
                return StarShape(w, h, 24, Math.Min(Math.Max(a, 0), 1));
            });

            PresetShapes.Add("star32", (w, h, adjustments) =>
            {
                // OOXML default adj=37500 → innerRatio = 0.75
                var a = Adjust(adjustments, "adj", 37500) * 2;
                return StarShape(w, h, 32, Math.Min(Math.Max(a, 0), 1));
            });

            // ==== Lines & Connectors ====
            // OOXML line: diagonal (0,0→w,h) when both extents are non-zero.
            // Keep explicit horizontal/vertical handling for zero-extent cases so 1px SVGs remain visible.
            PresetShapes.Add("line", (w, h, a) =>
            {
                var safeH = h == 0 ? 1 : h;
                var safeW = w == 0 ? 1 : w;
                if (w == 0)
                    return $"M0.5,0 L0.5,{safeH}";
                if (h == 0)
                    return $"M0,0.5 L{safeW},0.5";
                return $"M0,0 L{w},{h}";
            });

            // Inverse diagonal line (top-right to bottom-left).
            PresetShapes.Add("lineInv", (w, h, a) =>
            {
                var safeH = h == 0 ? 1 : h;
                var safeW = w == 0 ? 1 : w;
                if (w == 0)
                    return $"M0.5,0 L0.5,{safeH}";
                if (h == 0)
                    return $"M0,0.5 L{safeW},0.5";
                return $"M{w},0 L0,{h}";
            });

            // When one dimension is 0, draw horizontal or vertical line (same as "line") so gradient and stroke are correct
            PresetShapes.Add("straightConnector1", (w, h, a) =>
            {
                var safeH = h == 0 ? 1 : h;
                var safeW = w == 0 ? 1 : w;
                if (w == 0)
                    return $"M0.5,0 L0.5,{safeH}";
                if (h == 0)
                    return $"M0,0.5 L{safeW},0.5";
                return $"M0,0 L{w},{h}";
            });

            PresetShapes.Add("bentConnector2", (w, h, a) => $"M0,0 L{w},0 L{w},{h}");

            PresetShapes.Add("bentConnector3", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj1", 50000);
                var midX = w * a;
                return $"M0,0 L{midX},0 L{midX},{h} L{w},{h}";
            });

            PresetShapes.Add("bentConnector4", (w, h, adjustments) =>
            {
                var a1 = Adjust(adjustments, "adj1", 50000);
                var a2 = Adjust(adjustments, "adj2", 50000);
                var midX = w * a1;
                var midY = h * a2;
                return $"M0,0 L{midX},0 L{midX},{midY} L{w},{midY} L{w},{h}";
            });

            PresetShapes.Add("curvedConnector2", (w, h, a) =>
            {
                return $"M0,0 C{w / 2},0 {w},{h / 2} {w},{h}";
            });

            PresetShapes.Add("curvedConnector3", (w, h, adjustments) =>
            {
                // OOXML: two cubic Bezier segments joined at midpoint (x2, vc)
                var x2 = w * Adjust(adjustments, "adj1", 50000);
                var x1 = x2 / 2; // +/ l x2 2
                var x3 = (w + x2) / 2; // +/ r x2 2
                var vc = h / 2;
                var hd4 = h / 4;
                var y3 = (h * 3) / 4;
                return $"M0,0 C{x1},0 {x2},{hd4} {x2},{vc} C{x2},{y3} {x3},{h} {w},{h}";
            });

            PresetShapes.Add("curvedConnector4", (w, h, adjustments) =>
            {
                // OOXML: three cubic Bezier segments
                var x2 = w * Adjust(adjustments, "adj1", 50000);
                var y4 = h * Adjust(adjustments, "adj2", 50000);
                var x1 = x2 / 2; // +/ l x2 2
                var x3 = (w + x2) / 2; // +/ r x2 2
                var x4 = (x2 + x3) / 2; // +/ x2 x3 2
                var x5 = (x3 + w) / 2; // +/ x3 r 2
                var y1 = y4 / 2; // +/ t y4 2
                var y2 = y1 / 2; // +/ t y1 2
                var y3 = (y1 + y4) / 2; // +/ y1 y4 2
                var y5 = (h + y4) / 2; // +/ b y4 2
                return string.Join(" ", [
                    "M0,0",
                    $"C{x1},0 {x2},{y2} {x2},{y1}",
                    $"C{x2},{y3} {x4},{y4} {x3},{y4}",
                    $"C{x5},{y4} {w},{y5} {w},{h}",
                ]);
            });

            PresetShapes.Add("curvedConnector5", (w, h, adjustments) =>
            {
                // OOXML: four cubic Bezier segments
                var x3 = w * Adjust(adjustments, "adj1", 50000);
                var y4 = h * Adjust(adjustments, "adj2", 50000);
                var x6 = w * Adjust(adjustments, "adj3", 50000);
                var x1 = (x3 + x6) / 2; // +/ x3 x6 2
                var x2 = x3 / 2; // +/ l x3 2
                var x4 = (x3 + x1) / 2; // +/ x3 x1 2
                var x5 = (x6 + x1) / 2; // +/ x6 x1 2
                var x7 = (x6 + w) / 2; // +/ x6 r 2
                var y1 = y4 / 2; // +/ t y4 2
                var y2 = y1 / 2; // +/ t y1 2
                var y3 = (y1 + y4) / 2; // +/ y1 y4 2
                var y5 = (h + y4) / 2; // +/ b y4 2
                var y6 = (y5 + y4) / 2; // +/ y5 y4 2
                var y7 = (y5 + h) / 2; // +/ y5 b 2
                return string.Join(" ", [
                    "M0,0",
                    $"C{x2},0 {x3},{y2} {x3},{y1}",
                    $"C{x3},{y3} {x4},{y4} {x1},{y4}",
                    $"C{x5},{y4} {x6},{y6} {x6},{y5}",
                    $"C{x6},{y7} {x7},{h} {w},{h}",
                ]);
            });

            PresetShapes.Add("bentConnector5", (w, h, adjustments) =>
            {
                var a1 = Adjust(adjustments, "adj1", 50000);
                var a2 = Adjust(adjustments, "adj2", 50000);
                var a3 = Adjust(adjustments, "adj3", 50000);
                var x1 = w * a1;
                var y1 = h * a2;
                var x2 = w * a3;

                return $"M0,0 L{x1},0 L{x1},{y1} L{x2},{y1} L{x2},{h} L{w},{h}";
            });

            // ==== Arrow Shapes ====
            PresetShapes.Add("rightArrow", (w, h, adjustments) =>
            {
                var a1 = Adjust(adjustments, "adj1", 50000); // shaft width ratio
                var a2 = Adjust(adjustments, "adj2", 50000); // head length ratio
                var ss = Math.Min(w, h); // OOXML uses short side for head length
                var shaftHalfH = (h * a1) / 2;
                var headLen = ss * a2;
                var cy = h / 2;
                var shaftEnd = w - headLen;

                return string.Join(" ", [
                    $"M0,{cy - shaftHalfH}",
                    $"L{shaftEnd},{cy - shaftHalfH}",
                    $"L{shaftEnd},0",
                    $"L{w},{cy}",
                    $"L{shaftEnd},{h}",
                    $"L{shaftEnd},{cy + shaftHalfH}",
                    $"L0,{cy + shaftHalfH}",
                    "Z",
                ]);
            });

            PresetShapes.Add("leftArrow", (w, h, adjustments) =>
            {
                var a1 = Adjust(adjustments, "adj1", 50000);
                var a2 = Adjust(adjustments, "adj2", 50000);
                var ss = Math.Min(w, h);
                var shaftHalfH = (h * a1) / 2;
                var headLen = ss * a2;
                var cy = h / 2;

                return string.Join(" ", [
                    $"M{w},{cy - shaftHalfH}",
                    $"L{headLen},{cy - shaftHalfH}",
                    $"L{headLen},0",
                    $"L0,{cy}",
                    $"L{headLen},{h}",
                    $"L{headLen},{cy + shaftHalfH}",
                    $"L{w},{cy + shaftHalfH}",
                    "Z",
                ]);
            });

            PresetShapes.Add("upArrow", (w, h, adjustments) =>
            {
                var a1 = Adjust(adjustments, "adj1", 50000);
                var a2 = Adjust(adjustments, "adj2", 50000);
                var shaftHalfW = (w * a1) / 2;
                var headLen = h * a2;
                var cx = w / 2;

                return string.Join(" ", [
                    $"M{cx - shaftHalfW},{h}",
                    $"L{cx - shaftHalfW},{headLen}",
                    $"L0,{headLen}",
                    $"L{cx},0",
                    $"L{w},{headLen}",
                    $"L{cx + shaftHalfW},{headLen}",
                    $"L{cx + shaftHalfW},{h}",
                    "Z",
                ]);
            });

            PresetShapes.Add("downArrow", (w, h, adjustments) =>
            {
                var a1 = Adjust(adjustments, "adj1", 50000);
                var a2 = Adjust(adjustments, "adj2", 50000);
                var shaftHalfW = (w * a1) / 2;
                var headLen = h * a2;
                var cx = w / 2;
                var shaftEnd = h - headLen;
                return string.Join(" ", [
                        $"M{cx - shaftHalfW},0",
                        $"L{cx + shaftHalfW},0",
                        $"L{cx + shaftHalfW},{shaftEnd}",
                        $"L{w},{shaftEnd}",
                        $"L{cx},{h}",
                        $"L0,{shaftEnd}",
                        $"L{cx - shaftHalfW},{shaftEnd}",
                        "Z",
                ]);
            });

            PresetShapes.Add("downArrowCallout", (w, h, adjustments) =>
            {
                // ECMA-like callout geometry (4 adjustments).
                var adj1 = adjustments?["adj1"] ?? 25000;
                var adj2 = adjustments?["adj2"] ?? 25000;
                var adj3 = adjustments?["adj3"] ?? 25000;
                var adj4 = adjustments?["adj4"] ?? 64977;
                var ss = Math.Min(w, h);
                var a2 = Math.Max(0, Math.Min(adj2, (50000 * w) / Math.Max(ss, 1)));
                var a1 = Math.Max(0, Math.Min(adj1, a2 * 2));
                var a3 = Math.Max(0, Math.Min(adj3, (100000 * h) / Math.Max(ss, 1)));
                var q2 = (a3 * ss) / Math.Max(h, 1);
                var a4 = Math.Max(0, Math.Min(adj4, 100000 - q2));
                var hc = w / 2;
                var dx1 = (ss * a2) / 100000.0;
                var dx2 = (ss * a1) / 200000.0;
                var x1 = hc - dx1;
                var x2 = hc - dx2;
                var x3 = hc + dx2;
                var x4 = hc + dx1;
                var y3 = h - (ss * a3) / 100000.0;
                var y2 = (h * a4) / 100000.0;

                return string.Join(" ", [
                    "M0,0",
                    $"L{w},0",
                    $"L{w},{y2}",
                    $"L{x3},{y2}",
                    $"L{x3},{y3}",
                    $"L{x4},{y3}",
                    $"L{hc},{h}",
                    $"L{x1},{y3}",
                    $"L{x2},{y3}",
                    $"L{x2},{y2}",
                    $"L0,{y2}",
                    "Z",
                ]);
            });

            PresetShapes.Add("rightArrowCallout", (w, h, adjustments) =>
            {
                // OOXML: Rectangle body + right-pointing arrowhead (11-point polygon, 4 adj)
                var ss = Math.Min(w, h);
                var maxAdj2 = (50000 * h) / Math.Max(ss, 1);
                var a2 = Math.Max(0, Math.Min(adjustments?["adj2"] ?? 25000, maxAdj2));
                var a1 = Math.Max(0, Math.Min(adjustments?["adj1"] ?? 25000, a2 * 2));
                var maxAdj3 = (100000 * w) / Math.Max(ss, 1);
                var a3 = Math.Max(0, Math.Min(adjustments?["adj3"] ?? 25000, maxAdj3));
                var q2 = (a3 * ss) / Math.Max(w, 1);
                var a4 = Math.Max(0, Math.Min(adjustments?["adj4"] ?? 64977, 100000 - q2));
                var vc = h / 2;
                var dy1 = (ss * a2) / 100000.0;
                var dy2 = (ss * a1) / 200000.0;
                var y1 = vc - dy1;
                var y2 = vc - dy2;
                var y3 = vc + dy2;
                var y4 = vc + dy1;
                var dx3 = (ss * a3) / 100000.0;
                var x3 = w - dx3;
                var x2 = (w * a4) / 100000.0;

                return string.Join(" ", [
                    "M0,0",
                    $"L{x2},0",
                    $"L{x2},{y2}",
                    $"L{x3},{y2}",
                    $"L{x3},{y1}",
                    $"L{w},{vc}",
                    $"L{x3},{y4}",
                    $"L{x3},{y3}",
                    $"L{x2},{y3}",
                    $"L{x2},{h}",
                    $"L0,{h}",
                    "Z",
                ]);
            });

            PresetShapes.Add("leftArrowCallout", (w, h, adjustments) =>
            {
                // OOXML: Mirror of rightArrowCallout — arrowhead points left
                var ss = Math.Min(w, h);
                var maxAdj2 = (50000 * h) / Math.Max(ss, 1);
                var a2 = Math.Max(0, Math.Min(adjustments?["adj2"] ?? 25000, maxAdj2));
                var a1 = Math.Max(0, Math.Min(adjustments?["adj1"] ?? 25000, a2 * 2));
                var maxAdj3 = (100000 * w) / Math.Max(ss, 1);
                var a3 = Math.Max(0, Math.Min(adjustments?["adj3"] ?? 25000, maxAdj3));
                var q2 = (a3 * ss) / Math.Max(w, 1);
                var a4 = Math.Max(0, Math.Min(adjustments?["adj4"] ?? 64977, 100000 - q2));
                var vc = h / 2;
                var dy1 = (ss * a2) / 100000.0;
                var dy2 = (ss * a1) / 200000.0;
                var y1 = vc - dy1;
                var y2 = vc - dy2;
                var y3 = vc + dy2;
                var y4 = vc + dy1;
                var x1 = (ss * a3) / 100000.0;
                var dx2 = (w * a4) / 100000.0;
                var x2 = w - dx2;
                return string.Join(" ", [
                    $"M0,{vc}",
                    $"L{x1},{y1}",
                    $"L{x1},{y2}",
                    $"L{x2},{y2}",
                    $"L{x2},0",
                    $"L{w},0",
                    $"L{w},{h}",
                    $"L{x2},{h}",
                    $"L{x2},{y3}",
                    $"L{x1},{y3}",
                    $"L{x1},{y4}",
                    "Z",
                ]);
            });

            PresetShapes.Add("upArrowCallout", (w, h, adjustments) =>
            {
                // OOXML: Vertical variant — arrowhead points up
                var ss = Math.Min(w, h);
                var maxAdj2 = (50000 * w) / Math.Max(ss, 1);
                var a2 = Math.Max(0, Math.Min(adjustments?["adj2"] ?? 25000, maxAdj2));
                var a1 = Math.Max(0, Math.Min(adjustments?["adj1"] ?? 25000, a2 * 2));
                var maxAdj3 = (100000 * h) / Math.Max(ss, 1);
                var a3 = Math.Max(0, Math.Min(adjustments?["adj3"] ?? 25000, maxAdj3));
                var q2 = (a3 * ss) / Math.Max(h, 1);
                var a4 = Math.Max(0, Math.Min(adjustments?["adj4"] ?? 64977, 100000 - q2));
                var hc = w / 2;
                var dx1 = (ss * a2) / 100000.0;
                var dx2 = (ss * a1) / 200000.0;
                var x1 = hc - dx1;
                var x2 = hc - dx2;
                var x3 = hc + dx2;
                var x4 = hc + dx1;
                var y1 = (ss * a3) / 100000.0;
                var dy2 = (h * a4) / 100000.0;
                var y2 = h - dy2;
                return string.Join(" ", [
                    $"M0,{y2}",
                    $"L{x2},{y2}",
                    $"L{x2},{y1}",
                    $"L{x1},{y1}",
                    $"L{hc},0",
                    $"L{x4},{y1}",
                    $"L{x3},{y1}",
                    $"L{x3},{y2}",
                    $"L{w},{y2}",
                    $"L{w},{h}",
                    $"L0,{h}",
                    "Z",
                ]);
            });

            PresetShapes.Add("upDownArrowCallout", (w, h, adjustments) =>
            {
                // OOXML spec: 4 adjustments
                var adj1Raw = adjustments?["adj1"] ?? 25000;
                var adj2Raw = adjustments?["adj2"] ?? 25000;
                var adj3Raw = adjustments?["adj3"] ?? 25000;
                var adj4Raw = adjustments?["adj4"] ?? 48123;
                var ss = Math.Min(w, h);
                var a2 = Math.Max(0, Math.Min(adj2Raw, (50000 * w) / Math.Max(ss, 1)));
                var a1 = Math.Max(0, Math.Min(adj1Raw, a2 * 2));
                var a3 = Math.Max(0, Math.Min(adj3Raw, (50000 * h) / Math.Max(ss, 1)));
                var q2 = (a3 * ss) / Math.Max(h, 1);
                var a4 = Math.Max(0, Math.Min(adj4Raw, 100000 - q2 - q2));
                var dx1 = (ss * a2) / 100000.0;
                var dx2 = (ss * a1) / 200000.0;
                var hc = w / 2;
                var x1 = hc - dx1;
                var x2 = hc - dx2;
                var x3 = hc + dx2;
                var x4 = hc + dx1;
                var y1 = (ss * a3) / 100000.0;
                var dy2 = (h * a4) / 200000.0;
                var y2 = h / 2 - dy2;
                var y3 = h / 2 + dy2;
                var y4 = h - y1;

                return string.Join(" ", [
                    $"M{hc},0",
                    $"L{x4},{y1}",
                    $"L{x3},{y1}",
                    $"L{x3},{y2}",
                    $"L{w},{y2}",
                    $"L{w},{y3}",
                    $"L{x3},{y3}",
                    $"L{x3},{y4}",
                    $"L{x4},{y4}",
                    $"L{hc},{h}",
                    $"L{x1},{y4}",
                    $"L{x2},{y4}",
                    $"L{x2},{y3}",
                    $"L0,{y3}",
                    $"L0,{y2}",
                    $"L{x2},{y2}",
                    $"L{x2},{y1}",
                    $"L{x1},{y1}",
                    "Z",
                ]);
            });

            PresetShapes.Add("leftRightArrowCallout", (w, h, adjustments) =>
            {
                // OOXML spec: 4 adjustments
                var adj1Raw = adjustments?["adj1"] ?? 25000;
                var adj2Raw = adjustments?["adj2"] ?? 25000;
                var adj3Raw = adjustments?["adj3"] ?? 25000;
                var adj4Raw = adjustments?["adj4"] ?? 48123;
                var ss = Math.Min(w, h);
                var a2 = Math.Max(0, Math.Min(adj2Raw, (50000 * h) / Math.Max(ss, 1)));
                var a1 = Math.Max(0, Math.Min(adj1Raw, a2 * 2));
                var a3 = Math.Max(0, Math.Min(adj3Raw, (50000 * w) / Math.Max(ss, 1)));
                var q2 = (a3 * ss) / Math.Max(w, 1);
                var a4 = Math.Max(0, Math.Min(adj4Raw, 100000 - q2 - q2));
                var dy1 = (ss * a2) / 100000.0;
                var dy2 = (ss * a1) / 200000.0;
                var vc = h / 2;
                var y1 = vc - dy1;
                var y2 = vc - dy2;
                var y3 = vc + dy2;
                var y4 = vc + dy1;
                var x1 = (ss * a3) / 100000.0;
                var dx2 = (w * a4) / 200000.0;
                var x2 = w / 2 - dx2;
                var x3 = w / 2 + dx2;
                var x4 = w - x1;

                return string.Join(" ", [
                    $"M0,{vc}",
                    $"L{x1},{y1}",
                    $"L{x1},{y2}",
                    $"L{x2},{y2}",
                    $"L{x2},0",
                    $"L{x3},0",
                    $"L{x3},{y2}",
                    $"L{x4},{y2}",
                    $"L{x4},{y1}",
                    $"L{w},{vc}",
                    $"L{x4},{y4}",
                    $"L{x4},{y3}",
                    $"L{x3},{y3}",
                    $"L{x3},{h}",
                    $"L{x2},{h}",
                    $"L{x2},{y3}",
                    $"L{x1},{y3}",
                    $"L{x1},{y4}",
                    "Z",
                ]);
            });

            PresetShapes.Add("uturnArrow", (w, h, adjustments) =>
            {
                // ECMA-like U-turn arrow geometry (5 adjustments).
                var adj1 = adjustments?["adj1"] ?? 25000;
                var adj2 = adjustments?["adj2"] ?? 25000;
                var adj3 = adjustments?["adj3"] ?? 25000;
                var adj4 = adjustments?["adj4"] ?? 43750;
                var adj5 = adjustments?["adj5"] ?? 75000;
                var ss = Math.Min(w, h);
                var a2 = Math.Max(0, Math.Min(adj2, 25000));
                var a1 = Math.Max(0, Math.Min(adj1, a2 * 2));
                var q2 = (a1 * ss) / Math.Max(h, 1);
                var q3 = 100000 - q2;
                var a3 = Math.Max(0, Math.Min(adj3, (q3 * h) / Math.Max(ss, 1)));
                var minAdj5 = ((a3 + a1) * ss) / Math.Max(h, 1);
                var a5 = Math.Max(minAdj5, Math.Min(adj5, 100000));
                var th = (ss * a1) / 100000.0;
                var aw2 = (ss * a2) / 100000.0;
                var th2 = th / 2;
                var dh2 = aw2 - th2;
                var y5 = (h * a5) / 100000.0;
                var ah = (ss * a3) / 100000.0;
                var y4 = y5 - ah;
                var x9 = w - dh2;
                var bs = Math.Min(x9 / 2, y4);
                var a4 = Math.Max(0, Math.Min(adj4, (100000 * bs) / Math.Max(ss, 1)));
                var bd = (ss * a4) / 100000.0;
                var bd2 = Math.Max(bd - th, 0);
                var x3 = th + bd2;
                var x8 = w - aw2;
                var x6 = x8 - aw2;
                var x7 = x6 + dh2;
                var x4 = x9 - bd;
                var x5 = x7 - bd2;
                return string.Join(" ", [
                    $"M0,{h}",
                    $"L0,{bd}",
                    bd > 0.1 ? $"A{bd},{bd} 0 0,1 {bd},0" : "L0,0",
                    $"L{x4},0",
                    bd > 0.1 ? $"A{bd},{bd} 0 0,1 {x9},{bd}" : "L{x9},0",
                    $"L{x9},{y4}",
                    $"L{w},{y4}",
                    $"L{x8},{y5}",
                    $"L{x6},{y4}",
                    $"L{x7},{y4}",
                    $"L{x7},{x3}",
                    bd2 > 0.1 ? $"A{bd2},{bd2} 0 0,0 {x5},{th}" : "L{x5},{th}",
                    $"L{x3},{th}",
                    bd2 > 0.1 ? $"A{bd2},{bd2} 0 0,0 {th},{x3}" : "L{th},{x3}",
                    $"L{th},{h}",
                    "Z",
                ]);
            });

            PresetShapes.Add("leftRightArrow", (w, h, adjustments) =>
            {
                // OOXML: adj1=50000 (shaft width), adj2=50000 (head length based on ss)
                var ss = Math.Min(w, h);
                var hd2 = h / 2;
                var maxAdj2 = ss > 0 ? (50000 * w) / ss : 0;
                var a1 = Math.Min(Math.Max(adjustments?["adj1"] ?? 50000, 0), 100000);
                var a2 = Math.Min(Math.Max(adjustments?["adj2"] ?? 50000, 0), maxAdj2);
                var x2 = (ss * a2) / 100000.0;
                var x3 = w - x2;
                var dy = (h * a1) / 200000.0;
                var vc = hd2;
                var y1 = vc - dy;
                var y2 = vc + dy;
                var dx1 = hd2 > 0 ? (y1 * x2) / hd2 : 0;
                var _x1 = x2 - dx1;
                var _x4 = x3 + dx1;

                return string.Join(" ", [
                    $"M0,{vc}",
                    $"L{x2},0",
                    $"L{x2},{y1}",
                    $"L{x3},{y1}",
                    $"L{x3},0",
                    $"L{w},{vc}",
                    $"L{x3},{h}",
                    $"L{x3},{y2}",
                    $"L{x2},{y2}",
                    $"L{x2},{h}",
                    "Z",
                ]);
            });

            PresetShapes.Add("leftUpArrow", (w, h, adjustments) =>
            {
                // OOXML preset formula (presetShapeDefinitions.xml -> leftUpArrow)
                var rawAdj2 = Math.Max(0, Math.Min(adjustments?["adj2"] ?? 25000, 50000));
                var maxAdj1 = rawAdj2 * 2;
                var rawAdj1 = Math.Max(0, Math.Min(adjustments?["adj1"] ?? 25000, maxAdj1));
                var maxAdj3 = 100000 - maxAdj1;
                var rawAdj3 = Math.Max(0, Math.Min(adjustments?["adj3"] ?? 25000, maxAdj3));
                var ss = Math.Min(w, h);
                var x1 = (ss * rawAdj3) / 100000.0;
                var dx2 = (ss * rawAdj2) / 50000.0;
                var x2 = w - dx2;
                var y2 = h - dx2;
                var dx4 = (ss * rawAdj2) / 100000.0;
                var x4 = w - dx4;
                var y4 = h - dx4;
                var dx3 = (ss * rawAdj1) / 200000.0;
                var x3 = x4 - dx3;
                var x5 = x4 + dx3;
                var y3 = y4 - dx3;
                var y5 = y4 + dx3;

                return string.Join(" ", [
                    $"M0,{y4}",
                    $"L{x1},{y2}",
                    $"L{x1},{y3}",
                    $"L{x3},{y3}",
                    $"L{x3},{x1}",
                    $"L{x2},{x1}",
                    $"L{x4},0",
                    $"L{w},{x1}",
                    $"L{x5},{x1}",
                    $"L{x5},{y5}",
                    $"L{x1},{y5}",
                    $"L{x1},{h}",
                    "Z",
                ]);
            });

            PresetShapes.Add("upDownArrow", (w, h, adjustments) =>
            {
                // OOXML spec: adj1=50000 (shaft width), adj2=50000 (head length on ss)
                var adj1Raw = adjustments?["adj1"] ?? 50000;
                var adj2Raw = adjustments?["adj2"] ?? 50000;
                var ss = Math.Min(w, h);
                var maxAdj2 = (50000 * h) / Math.Max(ss, 1);
                var a2 = Math.Max(0, Math.Min(adj2Raw, maxAdj2));
                var a1 = Math.Max(0, Math.Min(adj1Raw, 100000));
                var dx1 = (ss * a1) / 200000.0; // shaft half-width
                var dy = (ss * a2) / 100000.0; // arrowhead length
                var hc = w / 2;

                return string.Join(" ", [
                    $"M{hc},0",
                    $"L{w},{dy}",
                    $"L{hc + dx1},{dy}",
                    $"L{hc + dx1},{h - dy}",
                    $"L{w},{h - dy}",
                    $"L{hc},{h}",
                    $"L0,{h - dy}",
                    $"L{hc - dx1},{h - dy}",
                    $"L{hc - dx1},{dy}",
                    $"L0,{dy}",
                    "Z",
                ]);
            });

            PresetShapes.Add("notchedRightArrow", (w, h, adjustments) =>
            {
                var a1 = Adjust(adjustments, "adj1", 50000); // shaft width ratio
                var a2 = Adjust(adjustments, "adj2", 50000); // head length ratio
                var ss = Math.Min(w, h); // OOXML uses short side for head length
                var shaftHalfH = (h * a1) / 2;
                var headLen = ss * a2;
                var cy = h / 2;
                var shaftEnd = w - headLen;
                // Notch depth: OOXML formula dxn = dy1 * dx2 / hd2 = shaftHalfH * headLen / (h/2)
                var notchDepth = cy > 0 ? (shaftHalfH * headLen) / cy : 0;

                return string.Join(" ", [
                    $"M0,{cy - shaftHalfH}",
                    $"L{shaftEnd},{cy - shaftHalfH}",
                    $"L{shaftEnd},0",
                    $"L{w},{cy}",
                    $"L{shaftEnd},{h}",
                    $"L{shaftEnd},{cy + shaftHalfH}",
                    $"L0,{cy + shaftHalfH}",
                    $"L{notchDepth},{cy}",
                    "Z",
                ]);
            });

            PresetShapes.Add("chevron", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 50000);
                var ss = Math.Min(w, h);
                var offset = ss * a;

                return string.Join(" ", [
                    "M0,0",
                    $"L{w - offset},0",
                    $"L{w},{h / 2}",
                    $"L{w - offset},{h}",
                    $"L0,{h}",
                    $"L{offset},{h / 2}",
                    "Z",
                ]);
            });

            PresetShapes.Add("homePlate", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 50000);
                var ss = Math.Min(w, h);
                var offset = ss * a;
                var shoulderX = w - offset;
                return string.Join(" ", ["M0,0", $"L{shoulderX},0", $"L{w},{h / 2}", $"L{shoulderX},{h}", $"L0,{h}", "Z"]);
            });

            PresetShapes.Add("stripedRightArrow", (w, h, adjustments) =>
            {
                // OOXML: adj1=50000, adj2=50000 (max 84375). Stripes at ssd32, ssd16-ssd8, x4=ss*5/32.
                var ss = Math.Min(w, h);
                var maxAdj2 = ss > 0 ? (84375 * w) / ss : 84375;
                var a1 = Math.Min(Math.Max(AdjustRaw(adjustments, "adj1", 50000), 0), 100000);
                var a2 = Math.Min(Math.Max(AdjustRaw(adjustments, "adj2", 50000), 0), maxAdj2);
                var dy1 = (h * a1) / 200000.0;
                var dx5 = (ss * a2) / 100000.0;
                var x5 = w - dx5;
                var vc = h / 2;
                var y1 = vc - dy1;
                var y2 = vc + dy1;
                var ssd32 = ss / 32;
                var ssd16 = ss / 16;
                var ssd8 = ss / 8;
                var x4 = (ss * 5) / 32;

                return string.Join(" ", [
                    // Stripe 1: 0 to ssd32
                    $"M0,{y1} L{ssd32},{y1} L{ssd32},{y2} L0,{y2} Z",
                    // Stripe 2: ssd16 to ssd8
                    $"M{ssd16},{y1} L{ssd8},{y1} L{ssd8},{y2} L{ssd16},{y2} Z",
                    // Main body + arrowhead: x4 to r
                    $"M{x4},{y1}",
                    $"L{x5},{y1}",
                    $"L{x5},0",
                    $"L{w},{vc}",
                    $"L{x5},{h}",
                    $"L{x5},{y2}",
                    $"L{x4},{y2}",
                    "Z",
                ]);
            });

            // ==== Bent / Curved / Special Arrows ====
            PresetShapes.Add("bentArrow", (w, h, adjustments) =>
            {
                // OOXML bentArrow: L-shaped arrow with rounded bend, arrowhead pointing right.
                // Uses 4 adjustments per ECMA-376 spec.
                var ss = Math.Min(w, h);
                // varrained adjustments (raw values, not fractions — we do our own math)
                var adj2Raw = Math.Max(0, Math.Min(adjustments?["adj2"] ?? 25000, 50000));
                var maxAdj1 = adj2Raw * 2;
                var adj1Raw = Math.Max(0, Math.Min(adjustments?["adj1"] ?? 25000, maxAdj1));
                var adj3Raw = Math.Max(0, Math.Min(adjustments?["adj3"] ?? 25000, 50000));
                var th = (ss * adj1Raw) / 100000.0; // shaft width
                var aw2 = (ss * adj2Raw) / 100000.0; // arrowhead half-width
                var th2 = th / 2;
                var dh2 = aw2 - th2; // arrowhead extension beyond shaft
                var ah = (ss * adj3Raw) / 100000.0; // arrowhead length
                var bw = w - ah;
                var bh = h - dh2;
                var bs = Math.Min(bw, bh);
                var maxAdj4 = bs > 0 ? (100000 * bs) / ss : 0;
                var adj4Raw = Math.Max(0, Math.Min(adjustments?["adj4"] ?? 43750, maxAdj4));
                var bd = (ss * adj4Raw) / 100000.0; // outer bend radius
                var bd2 = Math.Max(bd - th, 0); // inner bend radius
                var x3 = th + bd2;
                var x4 = w - ah;
                var y3 = dh2 + th;
                var y4 = y3 + dh2;
                var y5 = dh2 + bd;
                // OOXML arcTo: from current point, arc with radii (wR, hR), start angle stAng, sweep swAng.
                // Arc 1: outer bend — from (0, y5), radii=bd, 180°→270° (sweep +90°)
                //   Center of arc is at (bd, y5) relative, endpoint at (bd, y5-bd) = (bd, dh2)
                //   SVG: A bd,bd 0 0,1 bd,dh2
                // Arc 2: inner bend — from (x3, y3), radii=bd2, 270°→180° (sweep -90°)
                //   Center at (x3, y3+bd2), endpoint at (x3-bd2, y3+bd2) = (th, y3+bd2)
                //   SVG: A bd2,bd2 0 0,0 th,y6  where y6 = y3+bd2
                var y6 = y3 + bd2;

                var parts = new List<string>() {
                    $"M0,{h}", // bottom-left
                    $"L0,{y5}", // up left edge to arc start
                };

                // Outer arc (rounded bend, going from left edge up to top edge)
                if (bd > 0.1)
                {
                    parts.Add($"A{bd},{bd} 0 0,1 {bd},{dh2}");
                }
                else
                {
                    parts.Add($"L0,{dh2}"); // degenerate: straight corner
                }

                parts.Add($"L{x4},{dh2}"); // horizontal to arrowhead base (top)
                parts.Add($"L{x4},0"); // up to arrowhead top-left wing
                parts.Add($"L{w},{aw2}"); // arrowhead tip (pointing right)
                parts.Add($"L{x4},{y4}");// arrowhead bottom wing
                parts.Add($"L{x4},{y3}"); // back to arrowhead base (bottom)
                parts.Add($"L{x3},{y3}");

                // Inner arc (rounded bend, going from top down to right side of shaft)
                if (bd2 > 0.1)
                {
                    parts.Add($"A{bd2},{bd2} 0 0,0 {th},{y6}");
                }
                else
                {
                    parts.Add($"L{th},{y3}"); // degenerate: straight corner
                }

                parts.Add($"L{th},{h}"); // down right side of shaft to bottom
                parts.Add("Z");

                return string.Join(" ", parts);
            });

            PresetShapes.Add("bentUpArrow", (w, h, adjustments) =>
            {
                // OOXML preset formula (presetShapeDefinitions.xml -> bentUpArrow):
                // x/y variables are solved from adj1/2/3 in [0..50000], ss=min(w,h).
                var raw1 = Math.Max(0, Math.Min(adjustments?["adj1"] ?? 25000, 50000));
                var raw2 = Math.Max(0, Math.Min(adjustments?["adj2"] ?? 25000, 50000));
                var raw3 = Math.Max(0, Math.Min(adjustments?["adj3"] ?? 25000, 50000));
                var ss = Math.Min(w, h);
                var y1 = (ss * raw3) / 100000.0;
                var dx1 = (ss * raw2) / 50000.0;
                var x1 = w - dx1;
                var dx3 = (ss * raw2) / 100000.0;
                var x3 = w - dx3;
                var dx2 = (ss * raw1) / 200000.0;
                var x2 = x3 - dx2;
                var x4 = x3 + dx2;
                var dy2 = (ss * raw1) / 100000.0;
                var y2 = h - dy2;

                return string.Join(" ", [
                    $"M0,{y2}",
                    $"L{x2},{y2}",
                    $"L{x2},{y1}",
                    $"L{x1},{y1}",
                    $"L{x3},0",
                    $"L{w},{y1}",
                    $"L{x4},{y1}",
                    $"L{x4},{h}",
                    $"L0,{h}",
                    "Z",
                ]);
            });

            PresetShapes.Add("curvedRightArrow", (w, h, adjustments) =>
            {
                // Keep geometry aligned with OOXML preset math. Use local arc helper here
                // because preset formulas mix positive/negative sweeps that do not map 1:1
                // to the generic shapeArc() helper used in other shapes.
                var adj1Raw = adjustments?["adj1"] ?? 25000;
                var adj2Raw = adjustments?["adj2"] ?? 50000;
                var adj3Raw = adjustments?["adj3"] ?? 25000;
                var cnstVal1 = 50000;
                var cnstVal2 = 100000;
                var hd2 = h / 2;
                var r = w;
                var b = h;
                var l = 0;
                var c3d4 = 270;
                var cd2 = 180;
                var cd4 = 90;
                var ss = Math.Max(Math.Min(w, h), 1);
                var maxAdj2 = (cnstVal1 * h) / ss;
                var a2 = Math.Max(0, Math.Min(adj2Raw, maxAdj2));
                var a1 = Math.Max(0, Math.Min(adj1Raw, a2));
                var th = (ss * a1) / cnstVal2;
                var aw = (ss * a2) / cnstVal2;
                var q1 = (th + aw) / 4;
                var hR = hd2 - q1;
                var q7 = hR * 2;
                var q8 = q7 * q7;
                var q9 = th * th;
                var q10 = Math.Max(q8 - q9, 0);
                var q11 = Math.Sqrt(q10);
                var iDx = (q11 * w) / Math.Max(q7, 1e-6);
                var maxAdj3 = (cnstVal2 * iDx) / ss;
                var a3 = Math.Max(0, Math.Min(adj3Raw, maxAdj3));
                var ah = (ss * a3) / cnstVal2;
                var y3 = hR + th;
                var q2 = w * w;
                var q3 = ah * ah;
                var q4 = Math.Max(q2 - q3, 0);
                var q5 = Math.Sqrt(q4);
                var dy = (q5 * hR) / Math.Max(w, 1e-6);
                var y5 = hR + dy;
                var y7 = y3 + dy;
                var q6 = aw - th;
                var dh = q6 / 2;
                var y4 = y5 - dh;
                var y8 = y7 + dh;
                var aw2 = aw / 2;
                var y6 = b - aw2;
                var x1 = r - ah;
                var swAng = Math.Atan(dy / Math.Max(ah, 1e-6));
                var stAng = Math.PI - swAng;
                var mswAng = -swAng;
                var q12 = th / 2;
                var dang2 = Math.Atan2(q12, Math.Max(iDx, 1e-6));
                var swAng2 = dang2 - Math.PI / 2;
                var stAngDg = (stAng * 180) / Math.PI;
                var mswAngDg = (mswAng * 180) / Math.PI;
                var swAngDg = (swAng * 180) / Math.PI;
                var swAng2Dg = (swAng2 * 180) / Math.PI;

                Func<double, double, double, double, double, double, string> arc = (cx, cy, rx, ry, startDeg, endDeg) =>
                {
                    var s = (startDeg * Math.PI) / 180;
                    var e = (endDeg * Math.PI) / 180;
                    var xS = cx + rx * Math.Cos(s);
                    var yS = cy + ry * Math.Sin(s);
                    var xE = cx + rx * Math.Cos(e);
                    var yE = cy + ry * Math.Sin(e);
                    var delta = endDeg - startDeg;
                    var largeArc = Math.Abs(delta) > 180 ? 1 : 0;
                    var sweep = delta >= 0 ? 1 : 0;

                    return $"M{xS},{yS} A{rx},{ry} 0 {largeArc},{sweep} {xE},{yE}";
                };

                return string.Join(" ", [
                    $"M{l},{hR}",
                    arc(w, hR, w, hR, cd2, cd2 + mswAngDg).Replace("M", "L"),
                    $"L{x1},{y5}",
                    $"L{x1},{y4}",
                    $"L{r},{y6}",
                    $"L{x1},{y8}",
                    $"L{x1},{y7}",
                    arc(w, y3, w, hR, stAngDg, stAngDg + swAngDg).Replace("M", "L"),
                    "Z",
                    arc(w, hR, w, hR, cd2, cd2 + cd4),
                    $"L{r},{th}",
                    arc(w, y3, w, hR, c3d4, c3d4 + swAng2Dg).Replace("M", "L"),
                    "Z",
                ]);
            });

            PresetShapes.Add("curvedLeftArrow", (w, h, adjustments) => MirrorAbsolutePathHorizontally(PresetShapes["curvedRightArrow"](w, h, adjustments), w));

            /**
            * Convert OOXML arcTo to SVG arc endpoint and command string.
            * OOXML arcTo: wR, hR (radii), stAng, swAng (degrees).
            * Current point is at stAng on the arc ellipse.
            * Returns { path, endX, endY }.
            */
            PresetShapes.Add("curvedUpArrow", (w, h, adjustments) =>
            {
                Func<double, double, double, double, double, double, string> arc = (cx, cy, rx, ry, startDeg, endDeg) =>
                 {
                     var s = (startDeg * Math.PI) / 180;
                     var e = (endDeg * Math.PI) / 180;
                     var xS = cx + rx * Math.Cos(s);
                     var yS = cy + ry * Math.Sin(s);
                     var xE = cx + rx * Math.Cos(e);
                     var yE = cy + ry * Math.Sin(e);
                     var delta = endDeg - startDeg;
                     var largeArc = Math.Abs(delta) > 180 ? 1 : 0;
                     var sweep = delta >= 0 ? 1 : 0;

                     return $"M{xS},{yS} A{rx},{ry} 0 {largeArc},{sweep} {xE},{yE}";
                 };

                var ss = Math.Min(w, h);
                var wd2 = w / 2;
                var a1Raw = adjustments?["adj1"] ?? 25000;
                var a2Raw = adjustments?["adj2"] ?? 50000;
                var a3Raw = adjustments?["adj3"] ?? 25000;
                var maxAdj2 = (50000 * w) / Math.Max(ss, 1);
                var a2 = Math.Max(0, Math.Min(a2Raw, maxAdj2));
                var a1 = Math.Max(0, Math.Min(a1Raw, 100000));
                var th = (ss * a1) / 100000.0;
                var aw = (ss * a2) / 100000.0;
                var q1 = (th + aw) / 4;
                var wR = wd2 - q1;
                var q7 = wR * 2;
                var idy = (Math.Sqrt(Math.Max(q7 * q7 - th * th, 0)) * h) / Math.Max(q7, 1);
                var maxAdj3 = (100000 * idy) / Math.Max(ss, 1);
                var a3 = Math.Max(0, Math.Min(a3Raw, maxAdj3));
                var ah = (ss * a3) / 100000.0;
                var x3 = wR + th;
                var dx = (Math.Sqrt(Math.Max(h * h - ah * ah, 0)) * wR) / Math.Max(h, 1);
                var x5 = wR + dx;
                var x7 = x3 + dx;
                var dh = (aw - th) / 2;
                var x4 = x5 - dh;
                var x8 = x7 + dh;
                var x6 = w - aw / 2;
                var y1 = ah;
                var swAng = Math.Atan2(dx, ah);
                var dang2 = Math.Atan2(th / 2, idy);
                var stAng2 = Math.PI / 2 - dang2;
                var swAng2 = dang2 - swAng;
                var stAng3 = Math.PI / 2 - swAng;
                var stAng2Deg = (stAng2 * 180) / Math.PI;
                var swAng2Deg = (swAng2 * 180) / Math.PI;
                var stAng3Deg = (stAng3 * 180) / Math.PI;
                var swAngDeg = (swAng * 180) / Math.PI;

                return string.Join(" ", [
                     arc(wR, 0, wR, h, stAng2Deg, stAng2Deg + swAng2Deg),
                    $"L{x5},{y1}",
                    $"L{x4},{y1}",
                    $"L{x6},0",
                    $"L{x8},{y1}",
                    $"L{x7},{y1}",
                    arc(x3, 0, wR, h, stAng3Deg, stAng3Deg + swAngDeg).Replace("M", "L"),
                    $"L{wR},{h}",
                    arc(wR, 0, wR, h, 90, 180).Replace("M", "L"),
                    $"L{th},0",
                    arc(x3, 0, wR, h, 180, 90).Replace("M", "L"),
                    "Z"]);
            });

            PresetShapes.Add("curvedDownArrow", (w, h, adjustments) =>
            {
                Func<double, double, double, double, double, double, string> arc = (cx, cy, rx, ry, startDeg, endDeg) =>
                {
                    var s = (startDeg * Math.PI) / 180;
                    var e = (endDeg * Math.PI) / 180;
                    var xS = cx + rx * Math.Cos(s);
                    var yS = cy + ry * Math.Sin(s);
                    var xE = cx + rx * Math.Cos(e);
                    var yE = cy + ry * Math.Sin(e);
                    var delta = endDeg - startDeg;
                    var largeArc = Math.Abs(delta) > 180 ? 1 : 0;
                    var sweep = delta >= 0 ? 1 : 0;
                    return $"M{xS},{yS} A{rx},{ry} 0 {largeArc},{sweep} {xE},{yE}";
                };

                var ss = Math.Min(w, h);
                var wd2 = w / 2;
                var a1Raw = adjustments?["adj1"] ?? 25000;
                var a2Raw = adjustments?["adj2"] ?? 50000;
                var a3Raw = adjustments?["adj3"] ?? 25000;
                var maxAdj2 = (50000 * w) / Math.Max(ss, 1);
                var a2 = Math.Max(0, Math.Min(a2Raw, maxAdj2));
                var a1 = Math.Max(0, Math.Min(a1Raw, 100000));
                var th = (ss * a1) / 100000.0;
                var aw = (ss * a2) / 100000.0;
                var q1 = (th + aw) / 4;
                var wR = wd2 - q1;
                var q7 = wR * 2;
                var idy = (Math.Sqrt(Math.Max(q7 * q7 - th * th, 0)) * h) / Math.Max(q7, 1);
                var maxAdj3 = (100000 * idy) / Math.Max(ss, 1);
                var a3 = Math.Max(0, Math.Min(a3Raw, maxAdj3));
                var ah = (ss * a3) / 100000.0;
                var x3 = wR + th;
                var dx = (Math.Sqrt(Math.Max(h * h - ah * ah, 0)) * wR) / Math.Max(h, 1);
                var x5 = wR + dx;
                var x7 = x3 + dx;
                var dh = (aw - th) / 2;
                var x4 = x5 - dh;
                var x8 = x7 + dh;
                var x6 = w - aw / 2;
                var y1 = h - ah;
                var swAng = Math.Atan2(dx, ah);
                var swAngDeg = (swAng * 180) / Math.PI;
                var dang2 = Math.Atan2(th / 2, idy);
                var dang2Deg = (dang2 * 180) / Math.PI;
                var stAng = 270 + swAngDeg;
                var stAng2 = 270 - dang2Deg;
                var swAng2 = dang2Deg - 90;
                var swAng3 = 90 + dang2Deg;

                return string.Join(" ", [
                    $"M{x6},{h}",
                    $"L{x4},{y1}",
                    $"L{x5},{y1}",
                    arc(wR, h, wR, h, stAng, stAng - swAngDeg).Replace("M", "L"),
                    $"L{x3},0",
                    arc(x3, h, wR, h, 270, 270 + swAngDeg).Replace("M", "L"),
                    $"L{x5 + th},{y1}",
                    $"L{x8},{y1}",
                    "Z",
                    $"M{x3},0",
                    arc(x3, h, wR, h, stAng2, stAng2 + swAng2).Replace("M", "L"),
                    arc(wR, h, wR, h, 180, 180 + swAng3).Replace("M", "L"),
                    "Z"]);
            });

            PresetShapes.Add("circularArrow", (w, h, adjustments) =>
            {
                return BuildCircularArrowPath(w, h, adjustments, false, "circularArrow");
            });

            // leftCircularArrow uses same OOXML guide formulas as circularArrow but different default adjustments.
            PresetShapes.Add("leftCircularArrow", (w, h, adjustments) =>
            {
                return BuildCircularArrowPath(w, h, adjustments, false, "leftCircularArrow");
            });

            PresetShapes.Add("leftRightCircularArrow", (w, h, _adjustments) =>
            {
                // Build from the actual oracle PDF vector path (shape id 0177),
                // normalized to a 400x280 reference box.
                var sx = w / 400;
                var sy = h / 280;
                Func<double, double, PositionInfo> p = (x, y) => new PositionInfo() { X = x * sx, Y = y * sy };
                var p1 = p(35.0, 140.0);
                var p2 = p(19.9536, 89.9471);
                var p3 = p(33.4296, 89.9471);
                var c1 = p(74.6127, 28.1974);
                var c2 = p(182.5744, 0.5489);
                var p4 = p(274.5688, 28.1924);
                var c3 = p(315.4978, 40.4912);
                var c4 = p(348.2481, 62.4743);
                var p5 = p(366.5707, 89.9471);
                var p6 = p(380.0463, 89.9471);
                var p7 = p(365.0, 140.0);
                var p8 = p(310.0463, 89.9471);
                var p9 = p(320.9838, 89.9471);
                var c5 = p(274.3848, 50.3095);
                var c6 = p(182.4425, 40.5864);
                var p10 = p(115.6249, 68.2298);
                var c7 = p(101.3589, 74.1319);
                var c8 = p(88.9651, 81.4842);
                var p11 = p(79.0159, 89.947);
                var p12 = p(89.9536, 89.9471);

                return string.Join(" ", [
                    $"M{p1.X},{p1.Y}",
                        $"L{p2.X},{p2.Y}",
                        $"L{p3.X},{p3.Y}",
                        $"C{c1.X},{c1.Y} {c2.X},{c2.Y} {p4.X},{p4.Y}",
                        $"C{c3.X},{c3.Y} {c4.X},{c4.Y} {p5.X},{p5.Y}",
                        $"L{p6.X},{p6.Y}",
                        $"L{p7.X},{p7.Y}",
                        $"L{p8.X},{p8.Y}",
                        $"L{p9.X},{p9.Y}",
                        $"C{c5.X},{c5.Y} {c6.X},{c6.Y} {p10.X},{p10.Y}",
                        $"C{c7.X},{c7.Y} {c8.X},{c8.Y} {p11.X},{p11.Y}",
                        $"L{p12.X},{p12.Y}",
                        "Z",
                    ]);
            });

            PresetShapes.Add("quadArrow", (w, h, adjustments) =>
            {
                var adj1Raw = adjustments?["adj1"] ?? 22500;
                var adj2Raw = adjustments?["adj2"] ?? 22500;
                var adj3Raw = adjustments?["adj3"] ?? 22500;
                var vc = h / 2;
                var hc = w / 2;
                var minWH = Math.Min(w, h);
                var a2 = Math.Max(0, Math.Min(adj2Raw, 50000));
                var a1 = Math.Max(0, Math.Min(adj1Raw, 2 * a2));
                var a3 = Math.Max(0, Math.Min(adj3Raw, (100000 - 2 * a2) / 2));
                var x1 = (minWH * a3) / 100000.0;
                var dx2 = (minWH * a2) / 100000.0;
                var x2 = hc - dx2;
                var x5 = hc + dx2;
                var dx3 = (minWH * a1) / 200000.0;
                var x3 = hc - dx3;
                var x4 = hc + dx3;
                var x6 = w - x1;
                var y2 = vc - dx2;
                var y5 = vc + dx2;
                var y3 = vc - dx3;
                var y4 = vc + dx3;
                var y6 = h - x1;

                return string.Join(" ", [
                    $"M0,{vc}",
                    $"L{x1},{y2}",
                    $"L{x1},{y3}",
                    $"L{x3},{y3}",
                    $"L{x3},{x1}",
                    $"L{x2},{x1}",
                    $"L{hc},0",
                    $"L{x5},{x1}",
                    $"L{x4},{x1}",
                    $"L{x4},{y3}",
                    $"L{x6},{y3}",
                    $"L{x6},{y2}",
                    $"L{w},{vc}",
                    $"L{x6},{y5}",
                    $"L{x6},{y4}",
                    $"L{x4},{y4}",
                    $"L{x4},{y6}",
                    $"L{x5},{y6}",
                    $"L{hc},{h}",
                    $"L{x2},{y6}",
                    $"L{x3},{y6}",
                    $"L{x3},{y4}",
                    $"L{x1},{y4}",
                    $"L{x1},{y5}",
                    "Z",
                ]);
            });

            PresetShapes.Add("quadArrowCallout", (w, h, adjustments) =>
            {
                // OOXML: 28-point polygon with 4 arrowheads (4 adj)
                var ss = Math.Min(w, h);
                var hc = w / 2;
                var vc = h / 2;
                var a2 = Math.Max(0, Math.Min(adjustments?["adj2"] ?? 18515, 50000));
                var a1 = Math.Max(0, Math.Min(adjustments?["adj1"] ?? 18515, a2 * 2));
                var maxAdj3 = 50000 - a2;
                var a3 = Math.Max(0, Math.Min(adjustments?["adj3"] ?? 18515, maxAdj3));
                var q2 = a3 * 2;
                var a4 = Math.Max(a1, Math.Min(adjustments?["adj4"] ?? 48123, 100000 - q2));
                var dx2 = (ss * a2) / 100000.0;
                var dx3 = (ss * a1) / 200000.0;
                var ah = (ss * a3) / 100000.0;
                var dx1 = (w * a4) / 200000.0;
                var dy1 = (h * a4) / 200000.0;
                var x8 = w - ah;
                var x2 = hc - dx1;
                var x7 = hc + dx1;
                var x3 = hc - dx2;
                var x6 = hc + dx2;
                var x4 = hc - dx3;
                var x5 = hc + dx3;
                var y8 = h - ah;
                var y2 = vc - dy1;
                var y7 = vc + dy1;
                var y3 = vc - dx2;
                var y6 = vc + dx2;
                var y4 = vc - dx3;
                var y5 = vc + dx3;

                return string.Join(" ", [
                    $"M0,{vc}",
                    $"L{ah},{y3}",
                    $"L{ah},{y4}",
                    $"L{x2},{y4}",
                    $"L{x2},{y2}",
                    $"L{x4},{y2}",
                    $"L{x4},{ah}",
                    $"L{x3},{ah}",
                    $"L{hc},0",
                    $"L{x6},{ah}",
                    $"L{x5},{ah}",
                    $"L{x5},{y2}",
                    $"L{x7},{y2}",
                    $"L{x7},{y4}",
                    $"L{x8},{y4}",
                    $"L{x8},{y3}",
                    $"L{w},{vc}",
                    $"L{x8},{y6}",
                    $"L{x8},{y5}",
                    $"L{x7},{y5}",
                    $"L{x7},{y7}",
                    $"L{x5},{y7}",
                    $"L{x5},{y8}",
                    $"L{x6},{y8}",
                    $"L{hc},{h}",
                    $"L{x3},{y8}",
                    $"L{x4},{y8}",
                    $"L{x4},{y7}",
                    $"L{x2},{y7}",
                    $"L{x2},{y5}",
                    $"L{ah},{y5}",
                    $"L{ah},{y6}",
                    "Z",
                ]);
            });

            PresetShapes.Add("leftRightUpArrow", (w, h, adjustments) =>
            {
                // OOXML preset formula (presetShapeDefinitions.xml -> leftRightUpArrow)
                var rawAdj2 = Math.Max(0, Math.Min(adjustments?["adj2"] ?? 25000, 50000));
                var maxAdj1 = rawAdj2 * 2;
                var rawAdj1 = Math.Max(0, Math.Min(adjustments?["adj1"] ?? 25000, maxAdj1));
                var q1 = 100000 - maxAdj1;
                var maxAdj3 = q1 / 2.0;
                var rawAdj3 = Math.Max(0, Math.Min(adjustments?["adj3"] ?? 25000, maxAdj3));
                var ss = Math.Min(w, h);
                var hc = w / 2;
                var x1 = (ss * rawAdj3) / 100000.0;
                var dx2 = (ss * rawAdj2) / 100000.0;
                var x2 = hc - dx2;
                var x5 = hc + dx2;
                var dx3 = (ss * rawAdj1) / 200000.0;
                var x3 = hc - dx3;
                var x4 = hc + dx3;
                var x6 = w - x1;
                var dy2 = (ss * rawAdj2) / 50000.0;
                var y2 = h - dy2;
                var y4 = h - dx2;
                var y3 = y4 - dx3;
                var y5 = y4 + dx3;

                return string.Join(" ", [
                    $"M0,{y4}",
                    $"L{x1},{y2}",
                    $"L{x1},{y3}",
                    $"L{x3},{y3}",
                    $"L{x3},{x1}",
                    $"L{x2},{x1}",
                    $"L{hc},0",
                    $"L{x5},{x1}",
                    $"L{x4},{x1}",
                    $"L{x4},{y3}",
                    $"L{x6},{y3}",
                    $"L{x6},{y2}",
                    $"L{w},{y4}",
                    $"L{x6},{h}",
                    $"L{x6},{y5}",
                    $"L{x1},{y5}",
                    $"L{x1},{h}",
                    "Z",
                ]);
            });

            PresetShapes.Add("swooshArrow", (w, h, adjustments) =>
            {
                // OOXML swooshArrow: curved swoosh with arrowhead on the right.
                var ss = Math.Min(w, h);
                var raw1 = adjustments?["adj1"] ?? 25000;
                var raw2 = adjustments?["adj2"] ?? 16667;
                var a1 = Math.Max(1, Math.Min(raw1, 75000));
                var maxAdj2 = (70000 * w) / ss;
                var a2 = Math.Max(0, Math.Min(raw2, maxAdj2));
                var ad1 = (h * a1) / 100000.0;
                var ad2 = (ss * a2) / 100000.0;
                var ssd8 = ss / 8;
                var hd6 = h / 6;
                var alfa = Math.PI / 2 / 14; // cd4/14 in radians
                var tanAlfa = Math.Tan(alfa);
                var xB = w - ad2;
                var yB = ssd8;
                var dx0 = ssd8 * tanAlfa;
                var xC = xB - dx0;
                var dx1 = ad1 * tanAlfa;
                var yF = yB + ad1;
                var xF = xB + dx1;
                var xE = xF + dx0;
                var yE = yF + ssd8;
                var dy2 = yE;
                var dy22 = dy2 / 2;
                var dy3 = h / 20;
                var yD = dy22 + dy3;
                var xP1 = w / 6;
                var yP1 = hd6 + hd6; // h/3
                var dy5 = hd6 / 2;
                var yP2 = yF + dy5;
                var xP2 = w / 4;

                return string.Join(" ", [
                    $"M0,{h}",
                    $"Q{xP1},{yP1} {xB},{yB}",
                    $"L{xC},0",
                    $"L{w},{yD}",
                    $"L{xE},{yE}",
                    $"L{xF},{yF}",
                    $"Q{xP2},{yP2} 0,{h}",
                    "Z",
                ]);
            });

            // ==== Flowchart Shapes ====
            PresetShapes.Add("flowChartProcess", (w, h, a) => $"M0,0 L{w},0 L{w},{h} L0,{h} Z");

            PresetShapes.Add("flowChartDecision", (w, h, a) =>
            {
                var cx = w / 2;
                var cy = h / 2;
                return $"M{cx},0 L{w},{cy} L{cx},{h} L0,{cy} Z";
            });

            PresetShapes.Add("flowChartTerminator", (w, h, a) =>
            {
                // OOXML: path w=21600 h=21600, wR=3475, hR=10800 (elliptical caps, not circular)
                var x1 = (w * 3475) / 21600;
                var x2 = (w * 18125) / 21600;
                var wR = x1; // w * 3475/21600
                var hR = h / 2; // h * 10800/21600

                return string.Join(" ", [
                    $"M{x1},0",
                    $"L{x2},0",
                    $"A{wR},{hR} 0 0,1 {x2},{h}",
                    $"L{x1},{h}",
                    $"A{wR},{hR} 0 0,1 {x1},0",
                    "Z",
                ]);
            });

            PresetShapes.Add("flowChartDocument", (w, h, a) =>
            {
                // OOXML: path w=21600 h=21600, cubic (21600,17322)(10800,17322)(10800,23922)(0,20172)
                var y1 = (h * 17322) / 21600;
                var cy1 = y1; // h * 17322/21600
                var cy2 = (h * 23922) / 21600; // extends below h (overshoot for curve)
                var y2 = (h * 20172) / 21600;

                return string.Join(" ", ["M0,0", $"L{w},0", $"L{w},{y1}", $"C{w / 2},{cy1} {w / 2},{cy2} 0,{y2}", "Z"]);
            });

            PresetShapes.Add("flowChartInputOutput", (w, h, a) =>
            {
                // OOXML: path w=5 h=5, points: (0,5)(1,0)(5,0)(4,5) — offset = w/5
                var offset = w / 5;

                return $"M{offset},0 L{w},0 L{w - offset},{h} L0,{h} Z";
            });

            PresetShapes.Add("flowChartPredefinedProcess", (w, h, a) =>
            {
                var inset = w * 0.1;

                return string.Join(" ", [
                    // Outer rectangle
                    $"M0,0 L{w},0 L{w},{h} L0,{h} Z",
                    // Left inner line
                    $"M{inset},0 L{inset},{h}",
                    // Right inner line
                    $"M{w - inset},0 L{w - inset},{h}",
                ]);
            });

            PresetShapes.Add("flowChartAlternateProcess", (w, h, a) =>
            {
                // OOXML spec: corner radius = ssd6 = min(w,h)/6
                var r = Math.Min(w, h) / 6;

                return string.Join(" ", [
                    $"M{r},0",
                    $"L{w - r},0",
                    $"A{r},{r} 0 0,1 {w},{r}",
                    $"L{w},{h - r}",
                    $"A{r},{r} 0 0,1 {w - r},{h}",
                    $"L{r},{h}",
                    $"A{r},{r} 0 0,1 0,{h - r}",
                    $"L0,{r}",
                    $"A{r},{r} 0 0,1 {r},0",
                    "Z",
                ]);
            });

            PresetShapes.Add("flowChartManualInput", (w, h, a) =>
            {
                var topOffset = h * 0.2;
                return $"M0,{topOffset} L{w},0 L{w},{h} L0,{h} Z";
            });

            PresetShapes.Add("flowChartManualOperation", (w, h, a) =>
            {
                // OOXML: path w=5 h=5: (0,0)→(5,0)→(4,5)→(1,5)→close → inset = w/5
                return $"M0,0 L{w},0 L{(w * 4) / 5},{h} L{w / 5},{h} Z";
            });

            PresetShapes.Add("flowChartPreparation", (w, h, a) =>
            {
                var inset = w * 0.2;
                var cy = h / 2;

                return $"M{inset},0 L{w - inset},0 L{w},{cy} L{w - inset},{h} L{inset},{h} L0,{cy} Z";
            });

            PresetShapes.Add("flowChartData", (w, h, a) =>
            {
                var offset = w * 0.15;

                return $"M{offset},0 L{w},0 L{w - offset},{h} L0,{h} Z";
            });

            PresetShapes.Add("flowChartInternalStorage", (w, h, a) =>
            {
                var inset = Math.Min(w, h) * 0.12;

                return string.Join(" ", [
                    $"M0,0 L{w},0 L{w},{h} L0,{h} Z",
                    $"M{inset},0 L{inset},{h}",
                    $"M0,{inset} L{w},{inset}",
                ]);
            });

            PresetShapes.Add("flowChartMagneticDisk", (w, h, a) =>
            {
                // OOXML spec: path w=6 h=6, top at y=1, arc hR=1 → ry = h/6
                var ry = h / 6;
                var bodyTop = ry;
                var bodyBottom = h - ry;

                return string.Join(" ", [
                    // Top ellipse
                    $"M0,{bodyTop}",
                    $"A{w / 2},{ry} 0 1,1 {w},{bodyTop}",
                    // Right side down
                    $"L{w},{bodyBottom}",
                    // Bottom ellipse
                    $"A{w / 2},{ry} 0 1,1 0,{bodyBottom}",
                    // Left side up
                    $"L0,{bodyTop}",
                    "Z",
                    // Top ellipse visible arc (back half)
                    $"M{w},{bodyTop}",
                    $"A{w / 2},{ry} 0 1,1 0,{bodyTop}",
                ]);
            });

            PresetShapes.Add("flowChartDelay", (w, h, adjustments) =>
            {
                // OOXML: M(0,0) L(hc,0) arcTo(wd2,hd2, 270°, 180°) L(0,h) Z
                // Arc from (hc,0) with wR=w/2 hR=h/2, stAng=270° swAng=180° → semicircle right side
                var hc = w / 2;
                var a = OoArcTo(hc, 0, hc, h / 2, 270, 180);

                return string.Join(" ", ["M0,0", $"L{hc},0", a.SVG, $"L0,{h}", "Z"]);
            });

            PresetShapes.Add("flowChartDisplay", (w, h, adjustments) =>
            {
                // OOXML: path w=6 h=6, points: (0,3)(1,0)(5,0) arcTo(1,3,270°,180°) (1,6) close
                // Scaled: left point at (0, h/2), top-left at (w/6, 0), arc center at (5w/6, h/2)
                var sx = w / 6;
                var sy = h / 6;
                var arcWR = sx; // wR = 1 * (w/6)
                var arcHR = sy * 3; // hR = 3 * (h/6) = h/2
                var a = OoArcTo(5 * sx, 0, arcWR, arcHR, 270, 180);

                return string.Join(" ", [$"M0,{3 * sy}", $"L{sx},0", $"L{5 * sx},0", a.SVG, $"L{sx},{h}", "Z"]);
            });

            PresetShapes.Add("flowChartExtract", (w, h, a) => $"M{w / 2},0 L{w},{h} L0,{h} Z");

            PresetShapes.Add("flowChartMerge", (w, h, a) => $"M0,0 L{w},0 L{w / 2},{h} Z");

            PresetShapes.Add("flowChartOffpageConnector", (w, h, a) =>
            {
                var arrowH = h * 0.2;

                return string.Join(" ", ["M0,0", $"L{w},0", $"L{w},{h - arrowH}", $"L{w / 2},{h}", $"L0,{h - arrowH}", "Z"]);
            });

            PresetShapes.Add("flowChartConnector", (w, h, a) =>
            {
                var rx = w / 2;
                var ry = h / 2;
                return string.Join(" ", [$"M{w},{ry}", $"A{rx},{ry} 0 1,1 0,{ry}", $"A{rx},{ry} 0 1,1 {w},{ry}", "Z"]);
            });

            PresetShapes.Add("flowChartSort", (w, h, a) =>
            {
                var cx = w / 2;
                var cy = h / 2;

                return string.Join(" ", [$"M{cx},0 L{w},{cy} L{cx},{h} L0,{cy} Z", $"M0,{cy} L{w},{cy}"]);
            });

            PresetShapes.Add("flowChartCollate", (w, h, a) =>
            {
                var cx = w / 2;
                var cy = h / 2;

                return string.Join(" ", [
                    // top inverted triangle
                    $"M0,0 L{w},0 L{cx},{cy} Z",
                    // bottom upright triangle
                    $"M0,{h} L{w},{h} L{cx},{cy} Z",
                ]);
            });

            PresetShapes.Add("flowChartPunchedTape", (w, h, adjustments) =>
            {
                // OOXML: path w="20" h="20" with arcTo operations.
                // Start at (0, 2), four arcs for wavy top/bottom.
                var sx = w / 20;
                var sy = h / 20;

                Func<double, double, double, double, double, double, ArcToInfo> arcTo = (curX, curY, wR, hR, stAng60k, swAng60k) =>
                {
                    var stDeg = stAng60k / 60000;
                    var swDeg = swAng60k / 60000;
                    var stRad = (stDeg * Math.PI) / 180;
                    var endRad = ((stDeg + swDeg) * Math.PI) / 180;
                    var cx = curX - wR * Math.Cos(stRad);
                    var cy = curY - hR * Math.Sin(stRad);
                    var endX = cx + wR * Math.Cos(endRad);
                    var endY = cy + hR * Math.Sin(endRad);
                    var largeArc = Math.Abs(swDeg) > 180 ? 1 : 0;
                    var sweep = swDeg > 0 ? 1 : 0;

                    return new ArcToInfo() { EndX = endX, EndY = endY, SVG = $"A{wR},{hR} 0 {largeArc},{sweep} {endX},{endY}" };
                };

                // cd2 = 10800000 (180°)
                var wR = 5 * sx;
                var hR = 2 * sy;
                double x = 0, y = 2 * sy;
                var parts = new List<string>() { $"M{x},{y}" };
                // Top-left: stAng=cd2(180°), swAng=-cd2(-180°) → dips down
                var a = arcTo(x, y, wR, hR, 10800000, -10800000);
                parts.Add(a.SVG);
                x = a.EndX;
                y = a.EndY;
                // Top-right: stAng=cd2(180°), swAng=+cd2(+180°) → bumps up
                a = arcTo(x, y, wR, hR, 10800000, 10800000);
                parts.Add(a.SVG);
                x = a.EndX;
                y = a.EndY;
                // Line to bottom-right
                double bx = 20 * sx, by = 18 * sy;
                parts.Add($"L{bx},{by}");
                x = bx;
                y = by;
                // Bottom-right: stAng=0, swAng=-cd2(-180°) → bumps up
                a = arcTo(x, y, wR, hR, 0, -10800000);
                parts.Add(a.SVG);
                x = a.EndX;
                y = a.EndY;
                // Bottom-left: stAng=0, swAng=+cd2(+180°) → dips down
                a = arcTo(x, y, wR, hR, 0, 10800000);
                parts.Add(a.SVG);
                parts.Add("Z");

                return string.Join(" ", parts);
            });

            PresetShapes.Add("flowChartPunchedCard", (w, h, a) =>
            {
                // OOXML spec: path w=5, h=5. Points: (0,1)(1,0)(5,0)(5,5)(0,5)
                var sx = w / 5;
                var sy = h / 5;

                return $"M0,{sy} L{sx},0 L{w},0 L{w},{h} L0,{h} Z";
            });

            PresetShapes.Add("flowChartSummingJunction", (w, h, a) =>
            {
                // OOXML: Circle with X cross. Returns single path with circle + X lines.
                var wd2 = w / 2;
                var hd2 = h / 2;
                var idx = wd2 * Math.Cos(Math.PI / 4); // cos(45°)
                var idy = hd2 * Math.Sin(Math.PI / 4);
                var il = wd2 - idx;
                var ir = wd2 + idx;
                var it = hd2 - idy;
                var ib = hd2 + idy;

                return string.Join(" ", [
                    // Circle
                    $"M0,{hd2}",
                    $"A{wd2},{hd2} 0 1,1 {w},{hd2}",
                    $"A{wd2},{hd2} 0 1,1 0,{hd2}",
                    "Z",
                    // X cross
                    $"M{il},{it} L{ir},{ib}",
                    $"M{ir},{it} L{il},{ib}",
                ]);
            });

            PresetShapes.Add("flowChartOr", (w, h, a) =>
            {
                // OOXML: Circle with + cross.
                var wd2 = w / 2;
                var hd2 = h / 2;

                return string.Join(" ", [
                    // Circle
                    $"M0,{hd2}",
                    $"A{wd2},{hd2} 0 1,1 {w},{hd2}",
                    $"A{wd2},{hd2} 0 1,1 0,{hd2}",
                    "Z",
                    // + cross
                    $"M{wd2},0 L{wd2},{h}",
                    $"M0,{hd2} L{w},{hd2}",
                ]);
            });

            PresetShapes.Add("flowChartOnlineStorage", (w, h, a) =>
            {
                // OOXML: Rounded left side rectangle with concave right cap.
                // Normalized: left arc (convex) at x=w/6, right arc (concave) at x=w
                var x1 = w / 6;

                return string.Join(" ", [
                    $"M{x1},0",
                    $"L{w},0",
                    $"A{x1},{h / 2} 0 0,0 {w},{h}",
                    $"L{x1},{h}",
                    $"A{x1},{h / 2} 0 0,1 {x1},0",
                    "Z",
                ]);
            });

            PresetShapes.Add("flowChartMagneticDrum", (w, h, a) =>
            {
                // OOXML: Horizontal cylinder (magnetic drum). Right ellipse cap visible.
                var x1 = w / 6;
                var x2 = (w * 5) / 6;
                var ry = h / 2;

                return string.Join(" ", [
                    // Body
                    $"M{x1},0",
                    $"L{x2},0",
                    $"A{x1},{ry} 0 0,1 {x2},{h}",
                    $"L{x1},{h}",
                    $"A{x1},{ry} 0 0,1 {x1},0",
                    "Z",
                    // Right ellipse back-face (visible part)
                    $"M{x2},{h}",
                    $"A{x1},{ry} 0 0,1 {x2},0",
                ]);
            });

            PresetShapes.Add("flowChartMagneticTape", (w, h, a) =>
            {
                // OOXML: Nearly full ellipse (circle) with a tape tail to the bottom-right.
                // 3 quarter-arcs (270°) + partial arc of ang1 = at2(w,h) = atan2(h,w),
                // then line to (r, ib) → (r, b) → close.
                var wd2 = w / 2;
                var hd2 = h / 2;
                var hc = wd2;
                var vc = hd2;
                var ang1 = Math.Atan2(h, w); // OOXML at2 w h = atan2(h, w)
                var ib = vc + hd2 * Math.Sin(Math.PI / 4); // sin(45°) * hd2

                // arcTo helper: compute SVG arc from OOXML arcTo parameters
                Func<double, double, double, double, double, double, ArcToInfo> arcTo = (curX, curY, wR, hR, stDeg, swDeg) =>
                {
                    var stRad = (stDeg * Math.PI) / 180;
                    var endRad = ((stDeg + swDeg) * Math.PI) / 180;
                    var cx = curX - wR * Math.Cos(stRad);
                    var cy = curY - hR * Math.Sin(stRad);
                    var endX = cx + wR * Math.Cos(endRad);
                    var endY = cy + hR * Math.Sin(endRad);
                    var largeArc = Math.Abs(swDeg) > 180 ? 1 : 0;
                    var sweep = swDeg > 0 ? 1 : 0;

                    return new ArcToInfo() { EndX = endX, EndY = endY, SVG = $"A{wR},{hR} 0 {largeArc},{sweep} {endX},{endY}" };
                };

                // Start at bottom center: M(hc, b)
                var curX = hc;
                var curY = h;
                var a1 = arcTo(curX, curY, wd2, hd2, 90, 90); // cd4, cd4 → 90° to 180°
                curX = a1.EndX;
                curY = a1.EndY;
                var a2 = arcTo(curX, curY, wd2, hd2, 180, 90); // cd2, cd4 → 180° to 270°
                curX = a2.EndX;
                curY = a2.EndY;
                var a3 = arcTo(curX, curY, wd2, hd2, 270, 90); // 3cd4, cd4 → 270° to 360°
                curX = a3.EndX;
                curY = a3.EndY;
                var ang1Deg = (ang1 * 180) / Math.PI;
                var a4 = arcTo(curX, curY, wd2, hd2, 0, ang1Deg); // 0, ang1

                return string.Join(" ", [$"M{hc},{h}", a1.SVG, a2.SVG, a3.SVG, a4.SVG, $"L{w},{ib}", $"L{w},{h}", "Z"]);
            });

            PresetShapes.Add("flowChartMultidocument", (w, h, a) =>
            {
                // OOXML: 21600-unit coordinates. Three stacked documents with cubic bezier waves.
                Func<double, double> s = (x) => (w * x) / 21600;
                Func<double, double> t = (y) => (h * y) / 21600;

                return string.Join(" ", [
                    // Front doc (bottom layer, with wave)
                    $"M0,{t(20782)}",
                    $"C{s(9298)},{t(23542)} {s(9298)},{t(18022)} {s(18595)},{t(18022)}",
                    $"L{s(18595)},{t(3675)} L0,{t(3675)} Z",
                    // Middle doc
                    $"M{s(1532)},{t(3675)} L{s(1532)},{t(1815)} L{s(20000)},{t(1815)}",
                    $"L{s(20000)},{t(16252)}",
                    $"C{s(19298)},{t(16252)} {s(18595)},{t(16352)} {s(18595)},{t(16352)}",
                    $"L{s(18595)},{t(3675)} Z",
                    // Back doc (top layer)
                    $"M{s(2972)},{t(1815)} L{s(2972)},0 L{w},0",
                    $"L{w},{t(14392)}",
                    $"C{s(20800)},{t(14392)} {s(20000)},{t(14467)} {s(20000)},{t(14467)}",
                    $"L{s(20000)},{t(1815)} Z",
                ]);
            });

            // ==== Callout Shapes ====
            PresetShapes.Add("wedgeRectCallout", (w, h, adjustments) =>
            {
                // OOXML spec: adaptive callout pointer on the edge closest to the tip
                var hc = w / 2;
                var vc = h / 2;
                var dxPos = (w * (adjustments?["adj1"] ?? -20833)) / 100000.0;
                var dyPos = (h * (adjustments?["adj2"] ?? 62500)) / 100000.0;
                var xPos = hc + dxPos;
                var yPos = vc + dyPos;
                var dq = (dxPos * h) / w;
                var ady = Math.Abs(dyPos);
                var adq = Math.Abs(dq);
                var dz = ady - adq;
                // Notch bracket positions (7/12 or 2/12 depending on tip direction)
                var x1 = (w * (dxPos >= 0 ? 7 : 2)) / 12;
                var x2 = (w * (dxPos >= 0 ? 10 : 5)) / 12;
                var y1 = (h * (dyPos >= 0 ? 7 : 2)) / 12;
                var y2 = (h * (dyPos >= 0 ? 10 : 5)) / 12;
                // Conditional notch points per edge (collapse to edge if not the active edge)
                var xl = dz > 0 ? 0 : dxPos >= 0 ? 0 : xPos;
                var xt = dz > 0 ? (dyPos >= 0 ? x1 : xPos) : x1;
                var xr = dz > 0 ? w : dxPos >= 0 ? xPos : w;
                var xb = dz > 0 ? (dyPos >= 0 ? xPos : x1) : x1;
                var yl = dz > 0 ? y1 : dxPos >= 0 ? y1 : yPos;
                var yt = dz > 0 ? (dyPos >= 0 ? 0 : yPos) : 0;
                var yr = dz > 0 ? y1 : dxPos >= 0 ? yPos : y1;
                var yb = dz > 0 ? (dyPos >= 0 ? yPos : h) : h;

                return string.Join(" ", [
                    "M0,0",
                    $"L{x1},0",
                    $"L{xt},{yt}",
                    $"L{x2},0",
                    $"L{w},0",
                    $"L{w},{y1}",
                    $"L{xr},{yr}",
                    $"L{w},{y2}",
                    $"L{w},{h}",
                    $"L{x2},{h}",
                    $"L{xb},{yb}",
                    $"L{x1},{h}",
                    $"L0,{h}",
                    $"L0,{y2}",
                    $"L{xl},{yl}",
                    $"L0,{y1}",
                    "Z",
                ]);
            });

            PresetShapes.Add("wedgeRoundRectCallout", (w, h, adjustments) =>
            {
                // OOXML spec: rounded rect with adaptive callout pointer
                var hc = w / 2;
                var vc = h / 2;
                var ss = Math.Min(w, h);
                var dxPos = (w * (adjustments?["adj1"] ?? -20833)) / 100000.0;
                var dyPos = (h * (adjustments?["adj2"] ?? 62500)) / 100000.0;
                var u1 = (ss * (adjustments?["adj3"] ?? 16667)) / 100000.0;
                var xPos = hc + dxPos;
                var yPos = vc + dyPos;
                var dq = (dxPos * h) / w;
                var ady = Math.Abs(dyPos);
                var adq = Math.Abs(dq);
                var dz = ady - adq;
                var u2 = w - u1;
                var v2 = h - u1;
                var x1 = (w * (dxPos >= 0 ? 7 : 2)) / 12;
                var x2 = (w * (dxPos >= 0 ? 10 : 5)) / 12;
                var y1 = (h * (dyPos >= 0 ? 7 : 2)) / 12;
                var y2 = (h * (dyPos >= 0 ? 10 : 5)) / 12;
                var xl = dz > 0 ? 0 : dxPos >= 0 ? 0 : xPos;
                var xt = dz > 0 ? (dyPos >= 0 ? x1 : xPos) : x1;
                var xr = dz > 0 ? w : dxPos >= 0 ? xPos : w;
                var xb = dz > 0 ? (dyPos >= 0 ? xPos : x1) : x1;
                var yl = dz > 0 ? y1 : dxPos >= 0 ? y1 : yPos;
                var yt = dz > 0 ? (dyPos >= 0 ? 0 : yPos) : 0;
                var yr = dz > 0 ? y1 : dxPos >= 0 ? yPos : y1;
                var yb = dz > 0 ? (dyPos >= 0 ? yPos : h) : h;

                return string.Join(" ", [
                    $"M0,{u1}",
                    $"A{u1},{u1} 0 0,1 {u1},0",
                    $"L{x1},0",
                    $"L{xt},{yt}",
                    $"L{x2},0",
                    $"L{u2},0",
                    $"A{u1},{u1} 0 0,1 {w},{u1}",
                    $"L{w},{y1}",
                    $"L{xr},{yr}",
                    $"L{w},{y2}",
                    $"L{w},{v2}",
                    $"A{u1},{u1} 0 0,1 {u2},{h}",
                    $"L{x2},{h}",
                    $"L{xb},{yb}",
                    $"L{x1},{h}",
                    $"L{u1},{h}",
                    $"A{u1},{u1} 0 0,1 0,{v2}",
                    $"L0,{y2}",
                    $"L{xl},{yl}",
                    $"L0,{y1}",
                    "Z",
                ]);
            });

            PresetShapes.Add("wedgeEllipseCallout", (w, h, adjustments) =>
            {
                var ax = Adjust(adjustments, "adj1", -20833);
                var ay = Adjust(adjustments, "adj2", 62500);
                var rx = w / 2;
                var ry = h / 2;
                var tipX = rx + w * ax;
                var tipY = ry + h * ay;
                // Approximate: ellipse with a pointer
                var angle = Math.Atan2(tipY - ry, tipX - rx);
                var gapAngle = 0.15;
                var _x1 = rx + rx * Math.Cos(angle - gapAngle);
                var _y1 = ry + ry * Math.Sin(angle - gapAngle);
                var _x2 = rx + rx * Math.Cos(angle + gapAngle);
                var _y2 = ry + ry * Math.Sin(angle + gapAngle);

                return string.Join(" ", [
                    ShapeArc.Arc(rx, ry, rx, ry, ((angle + gapAngle) * 180) / Math.PI, ((angle - gapAngle + 2 * Math.PI) * 180) / Math.PI, false),
                    $"L{tipX},{tipY}",
                    "Z",
                ]);
            });

            PresetShapes.Add("cloudCallout", (w, h, adjustments) =>
            {
                var ax = Adjust(adjustments, "adj1", -20833);
                var ay = Adjust(adjustments, "adj2", 62500);
                var tipX = w / 2 + w * ax;
                var tipY = h / 2 + h * ay;
                // Simplified cloud with callout circles
                var cloud = PresetShapes["cloud"](w, h, null);
                // Small circles leading to tip
                var cx = w / 2;
                var cy = h / 2;
                var dx = tipX - cx;
                var dy = tipY - cy;
                var r1 = Math.Min(w, h) * 0.04;
                var r2 = Math.Min(w, h) * 0.025;
                var c1x = cx + dx * 0.5;
                var c1y = cy + dy * 0.5;
                var c2x = cx + dx * 0.75;
                var c2y = cy + dy * 0.75;

                return string.Join(" ", [
                    cloud,
                    // Connector circles (approximated as small ellipses)
                    $"M{c1x + r1},{c1y} A{r1},{r1} 0 1,1 {c1x - r1},{c1y} A{r1},{r1} 0 1,1 {c1x + r1},{c1y} Z",
                    $"M{c2x + r2},{c2y} A{r2},{r2} 0 1,1 {c2x - r2},{c2y} A{r2},{r2} 0 1,1 {c2x + r2},{c2y} Z",
                ]);
            });

            PresetShapes.Add("borderCallout1", (w, h, adjustments) =>
            {
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 112500)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -38333)) / 100000.0;

                return $"M0,0 L{w},0 L{w},{h} L0,{h} Z M{x1},{y1} L{x2},{y2}";
            });

            // ==== Block / 3D Shapes ====
            PresetShapes.Add("cube", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 25000);
                var depth = Math.Min(w, h) * a;

                return string.Join(" ", [
                    // Front face
                    $"M0,{depth} L{w - depth},{depth} L{w - depth},{h} L0,{h} Z",
                    // Top face
                    $"M0,{depth} L{depth},0 L{w},0 L{w - depth},{depth} Z",
                    // Right face
                    $"M{w - depth},{depth} L{w},0 L{w},{h - depth} L{w - depth},{h} Z",
                ]);
            });

            // can is implemented as multiPathPreset (see multiPathPresets below)
            // ribbon2 is implemented as multiPathPreset (see multiPathPresets below)
            PresetShapes.Add("plus", (w, h, adjustments) =>
            {
                // OOXML: adj=25000 (max 50000), x1 = ss * a / 100000 (uses ss for both x and y)
                var ss = Math.Min(w, h);
                var a = Math.Min(Math.Max(AdjustRaw(adjustments, "adj", 25000), 0), 50000);
                var x1 = (ss * a) / 100000.0;
                var x2 = w - x1;
                var y2 = h - x1;

                return string.Join(" ", [
                    $"M0,{x1}",
                    $"L{x1},{x1}",
                    $"L{x1},0",
                    $"L{x2},0",
                    $"L{x2},{x1}",
                    $"L{w},{x1}",
                    $"L{w},{y2}",
                    $"L{x2},{y2}",
                    $"L{x2},{h}",
                    $"L{x1},{h}",
                    $"L{x1},{y2}",
                    $"L0,{y2}",
                    "Z",
                ]);
            });

            PresetShapes.Add("heart", (w, h, a) =>
            {
                // OOXML spec: two cubic beziers from (hc, hd4) through (hc, b) and back.
                // dx1 = w*49/48 (slightly wider than w/2), dx2 = w*10/48
                // y1 = t - hd3 (above top edge)
                var hc = w / 2;
                var hd4 = h / 4;
                var hd3 = h / 3;
                var dx1 = (w * 49) / 48;
                var dx2 = (w * 10) / 48;
                var x1 = hc - dx1; // far left control
                var x2 = hc - dx2; // inner left control
                var x3 = hc + dx2; // inner right control
                var x4 = hc + dx1; // far right control
                var y1 = -hd3; // above top (negative y)

                return string.Join(" ", [
                    $"M{hc},{hd4}",
                    $"C{x3},{y1} {x4},{hd4} {hc},{h}",
                    $"C{x1},{hd4} {x2},{y1} {hc},{hd4}",
                    "Z",
                ]);
            });

            PresetShapes.Add("cloud", (w, h, a) =>
            {
                // OOXML cloud: 11 arcTo operations in 43200×43200 coordinate space
                var sx = w / 43200;
                var sy = h / 43200;

                // OOXML arcTo: wR/hR are radii, stAng/swAng in 60000ths of degree
                double[][] arcs = [
                    [6753, 9190, -11429249, 7426832],
                    [5333, 7267, -8646143, 5396714],
                    [4365, 5945, -8748475, 5983381],
                    [4857, 6595, -7859164, 7034504],
                    [5333, 7273, -4722533, 6541615],
                    [6775, 9220, -2776035, 7816140],
                    [5785, 7867, 37501, 6842000],
                    [6752, 9215, 1347096, 6910353],
                    [7720, 10543, 3974558, 4542661],
                    [4360, 5918, -16496525, 8804134],
                    [4345, 5945, -14809710, 9151131],
                ];

                var curX = 3900 * sx;
                var curY = 14370 * sy;
                var parts = new List<string>() { $"M{curX},{curY}" };
                // Track position in unscaled 43200×43200 space for accurate arcTo computation.
                // OOXML arcTo angles are visual (geometric ray) angles in the path coordinate space.
                // Convert to parametric before computing center/endpoint positions.
                double ux = 3900, uy = 14370; // unscaled current position

                foreach (var item in arcs)
                {
                    double wR = item[0], hR = item[1], stAng60k = item[2], swAng60k = item[3];

                    var stDeg = stAng60k / 60000;
                    var swDeg = swAng60k / 60000;
                    // Visual→parametric using UNSCALED radii (path coordinate space)
                    var stVisRad = (stDeg * Math.PI) / 180;
                    var stRad = Math.Atan2(wR * Math.Sin(stVisRad), hR * Math.Cos(stVisRad));
                    var endVisRad = ((stDeg + swDeg) * Math.PI) / 180;
                    var endRad = Math.Atan2(wR * Math.Sin(endVisRad), hR * Math.Cos(endVisRad));
                    // Compute center and endpoint in unscaled space
                    var acx = ux - wR * Math.Cos(stRad);
                    var acy = uy - hR * Math.Sin(stRad);
                    var endUX = acx + wR * Math.Cos(endRad);
                    var endUY = acy + hR * Math.Sin(endRad);
                    // Scale to pixel space for SVG output
                    var endX = endUX * sx;
                    var endY = endUY * sy;
                    var rwS = wR * sx;
                    var rhS = hR * sy;
                    var largeArc = Math.Abs(swDeg) > 180 ? 1 : 0;
                    var sweep = swDeg > 0 ? 1 : 0;
                    parts.Add($"A{rwS},{rhS} 0 {largeArc},{sweep} {endX},{endY}");
                    ux = endUX;
                    uy = endUY;
                    curX = endX;
                    curY = endY;
                }

                parts.Add("Z");
                return string.Join(" ", parts);
            });

            // ==== Frame, Donut, Misc ====
            PresetShapes.Add("frame", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj1", 12500);
                var t = Math.Min(w, h) * a;

                return string.Join(" ", [
                    // Outer rectangle
                    $"M0,0 L{w},0 L{w},{h} L0,{h} Z",
                    // Inner rectangle (counter-clockwise for hole)
                    $"M{t},{t} L{t},{h - t} L{w - t},{h - t} L{w - t},{t} Z",
                ]);
            });

            PresetShapes.Add("halfFrame", (w, h, adjustments) =>
            {
                // OOXML spec defaults: adj1=33333, adj2=33333
                var adj1Raw = adjustments?["adj1"] ?? 33333;
                var adj2Raw = adjustments?["adj2"] ?? 33333;
                var minWH = Math.Min(w, h);
                var a2 = Math.Max(0, Math.Min(adj2Raw, (100000 * w) / Math.Max(minWH, 1)));
                var x1 = (minWH * a2) / 100000.0;
                var g1 = (h * x1) / Math.Max(w, 1);
                var g2 = h - g1;
                var a1 = Math.Max(0, Math.Min(adj1Raw, (100000 * g2) / Math.Max(minWH, 1)));
                var y1 = (minWH * a1) / 100000.0;
                var x2 = w - (y1 * w) / Math.Max(h, 1);
                var y2 = h - (x1 * h) / Math.Max(w, 1);

                return string.Join(" ", ["M0,0", $"L{w},0", $"L{x2},{y1}", $"L{x1},{y1}", $"L{x1},{y2}", $"L0,{h}", "Z"]);
            });

            PresetShapes.Add("donut", (w, h, adjustments) =>
            {
                // OOXML: adj=25000, dr = ss * a / 100000, inner radii = wd2-dr, hd2-dr
                var ss = Math.Min(w, h);
                var a = Math.Min(Math.Max(AdjustRaw(adjustments, "adj", 25000), 0), 50000);
                var dr = (ss * a) / 100000.0;
                var rx = w / 2;
                var ry = h / 2;
                var iwd2 = Math.Max(0, rx - dr);
                var ihd2 = Math.Max(0, ry - dr);

                return string.Join(" ", [
                    // Outer circle (CW)
                    $"M0,{ry}",
                    $"A{rx},{ry} 0 1,1 {w},{ry}",
                    $"A{rx},{ry} 0 1,1 0,{ry}",
                    "Z",
                    // Inner circle (CCW for evenodd hole)
                    $"M{dr},{ry}",
                    $"A{iwd2},{ihd2} 0 1,0 {w - dr},{ry}",
                    $"A{iwd2},{ihd2} 0 1,0 {dr},{ry}",
                    "Z",
                ]);
            });

            PresetShapes.Add("noSmoking", (w, h, adjustments) =>
            {
                // OOXML: adj=18750. Ring thickness = ss*a/100000. Diagonal band via inner ellipse arcs + evenodd.
                var ss = Math.Min(w, h);
                var a = Math.Min(Math.Max(AdjustRaw(adjustments, "adj", 18750), 0), 50000);
                var dr = (ss * a) / 100000.0;
                var rx = w / 2;
                var ry = h / 2;
                var hc = w / 2;
                var vc = h / 2;
                var iwd2 = rx - dr;
                var ihd2 = ry - dr;
                // Compute diagonal angle and band intersection with inner ellipse
                var ang = Math.Atan2(h, w); // at2(w, h) in OOXML: at2 x y = atan2(y, x)
                // Inner ellipse radius at diagonal angle
                var ct = ihd2 * Math.Cos(ang);
                var st = iwd2 * Math.Sin(ang);
                var s = ct * ct + st * st;
                var m = s != 0 ? Math.Sqrt(ct * ct + st * st) : 1;
                var n = (iwd2 * ihd2) / m;
                var drd2 = dr / 2;
                var dang = Math.Atan2(drd2, n);
                var dang2 = dang * 2;
                // Sweep for inner arcs: -(180° - dang2) expressed as OOXML 60000ths then converted
                var swAngRad = -(Math.PI - dang2);
                var stAng1 = ang - dang;
                var stAng2 = stAng1 - Math.PI;

                // Compute points on inner ellipse for the two diagonal band arcs
                Func<double, PositionInfo> innerPt = (angle) =>
                {
                    var ct2 = ihd2 * Math.Cos(angle);
                    var st2 = iwd2 * Math.Sin(angle);
                    var t = ct2 * ct2 + st2 * st2;
                    var m2 = t != 0 ? Math.Sqrt(t) : 1;
                    var n2 = (iwd2 * ihd2) / m2;
                    return new PositionInfo() { X = hc + n2 * Math.Cos(angle), Y = vc + n2 * Math.Sin(angle) };
                };

                var p1 = innerPt(stAng1);
                var p2 = innerPt(stAng2);
                // End points of arcs
                var endAng1 = stAng1 + swAngRad;
                var endAng2 = stAng2 + swAngRad;
                var e1 = innerPt(endAng1);
                var e2 = innerPt(endAng2);
                var largeArc = Math.Abs(swAngRad) > Math.PI ? 1 : 0;
                var sweep = swAngRad > 0 ? 1 : 0;

                return string.Join(" ", [
                    // Outer circle (CW)
                    $"M0,{vc}",
                    $"A{rx},{ry} 0 1,1 {w},{vc}",
                    $"A{rx},{ry} 0 1,1 0,{vc}",
                    "Z",
                    // First diagonal band arc (inner ellipse)
                    $"M{p1.X},{p1.Y}",
                    $"A{iwd2},{ihd2} 0 {largeArc},{sweep} {e1.X},{e1.Y}",
                    "Z",
                    // Second diagonal band arc (opposite quadrant)
                    $"M{p2.X},{p2.Y}",
                    $"A{iwd2},{ihd2} 0 {largeArc},{sweep} {e2.X},{e2.Y}",
                    "Z",
                ]);
            });

            PresetShapes.Add("blockArc", (w, h, adjustments) =>
            {
                var adj1Raw = adjustments?["adj1"] ?? 10800000; // start angle
                var adj2Raw = adjustments?["adj2"] ?? 0; // sweep/end angle
                var adj3Raw = adjustments?["adj3"] ?? 25000; // thickness ratio
                var startDeg = Math.Min(Math.Max(adj1Raw / 60000, 0), 360);
                var innerStartDeg = Math.Min(Math.Max(adj2Raw / 60000, 0), 360);
                var s = (innerStartDeg - startDeg + 360) % 360;
                var sweepDeg = s == 0 ? 360 : s;
                var endDeg = startDeg + sweepDeg;
                var innerEndDeg = innerStartDeg - sweepDeg;
                var wd2 = w / 2;
                var hd2 = h / 2;
                var dr = (Math.Min(w, h) * Math.Max(0, Math.Min(adj3Raw, 50000))) / 100000.0;
                var iwd2 = Math.Max(1, wd2 - dr);
                var ihd2 = Math.Max(1, hd2 - dr);
                Func<double, double, double, double, double, PositionInfo> p = (cx, cy, rx, ry, deg) =>
                {
                    var r = (deg * Math.PI) / 180;
                    return new PositionInfo() { X = cx + rx * Math.Cos(r), Y = cy + ry * Math.Sin(r) };
                };

                var oStart = p(wd2, hd2, wd2, hd2, startDeg);
                var oEnd = p(wd2, hd2, wd2, hd2, endDeg);
                var iStart = p(wd2, hd2, iwd2, ihd2, innerStartDeg);
                var iEnd = p(wd2, hd2, iwd2, ihd2, innerEndDeg);
                var largeArc = sweepDeg > 180 ? 1 : 0;

                return string.Join(" ", [
                    $"M{oStart.X},{oStart.Y}",
                    $"A{wd2},{hd2} 0 {largeArc},1 {oEnd.X},{oEnd.Y}",
                    $"L{iStart.X},{iStart.Y}",
                    $"A{iwd2},{ihd2} 0 {largeArc},0 {iEnd.X},{iEnd.Y}",
                    "Z",
                ]);
            });

            // ==== Gear Shapes ====
            PresetShapes.Add("gear6", (w, h, adjustments) =>
            {
                var a1 = adjustments?["adj1"] ?? 15000;
                var a2 = adjustments?["adj2"] ?? 3526;
                return GearShape(w, h, 6, a1, a2);
            });

            PresetShapes.Add("gear9", (w, h, adjustments) =>
            {
                var a1 = adjustments?["adj1"] ?? 10000;
                var a2 = adjustments?["adj2"] ?? 1763;

                return GearShape(w, h, 9, a1, a2);
            });

            // ==== Misc Shapes ====
            PresetShapes.Add("mathPlus", (w, h, adjustments) =>
            {
                // OOXML: adj1=23520 (max 73490). dx1 = w*73490/200000, dx2 = ss*a/200000
                var ss = Math.Min(w, h);
                var a1 = Math.Min(Math.Max(AdjustRaw(adjustments, "adj", 23520), 0), 73490);
                var dx1 = (w * 73490) / 200000.0;
                var dy1 = (h * 73490) / 200000.0;
                var dx2 = (ss * a1) / 200000.0;
                var hc = w / 2;
                var vc = h / 2;
                var x1 = hc - dx1;
                var x2 = hc - dx2;
                var x3 = hc + dx2;
                var x4 = hc + dx1;
                var y1 = vc - dy1;
                var y2 = vc - dx2;
                var y3 = vc + dx2;
                var y4 = vc + dy1;

                return string.Join(" ", [
                    $"M{x1},{y2}",
                    $"L{x2},{y2}",
                    $"L{x2},{y1}",
                    $"L{x3},{y1}",
                    $"L{x3},{y2}",
                    $"L{x4},{y2}",
                    $"L{x4},{y3}",
                    $"L{x3},{y3}",
                    $"L{x3},{y4}",
                    $"L{x2},{y4}",
                    $"L{x2},{y3}",
                    $"L{x1},{y3}",
                    "Z",
                ]);
            });

            PresetShapes.Add("mathMinus", (w, h, adjustments) =>
            {
                // OOXML: adj1=23520 (max 100000). dy1 = h*a1/200000, dx1 = w*73490/200000
                var a1 = Math.Min(Math.Max(AdjustRaw(adjustments, "adj1", 23520), 0), 100000);
                var dy1 = (h * a1) / 200000.0;
                var dx1 = (w * 73490) / 200000.0;
                var hc = w / 2;
                var vc = h / 2;
                var x1 = hc - dx1;
                var x2 = hc + dx1;
                var y1 = vc - dy1;
                var y2 = vc + dy1;

                return $"M{x1},{y1} L{x2},{y1} L{x2},{y2} L{x1},{y2} Z";
            });

            PresetShapes.Add("mathMultiply", (w, h, adjustments) =>
            {
                // OOXML: adj1=23520 (max 51965). X shape with diagonal arms.
                // Key: a = at2 w h → atan2(w, h), coordinates are absolute from top-left.
                var ss = Math.Min(w, h);
                var hc = w / 2;
                var vc = h / 2;
                var a1 = Math.Min(Math.Max(AdjustRaw(adjustments, "adj1", 23520), 0), 51965);
                var th = (ss * a1) / 100000.0;
                var a = Math.Atan2(h, w);
                var sa = Math.Sin(a);
                var ca = Math.Cos(a);
                var ta = sa / ca; // tan(a)
                var dl = Math.Sqrt(w * w + h * h);
                var rw = (dl * 51965) / 100000.0;
                var lM = dl - rw;
                // xM, yM: half-distance along the diagonal from the outer tip to the outer tip
                var xM = (ca * lM) / 2;
                var yM = (sa * lM) / 2;
                // Perpendicular offset for arm thickness
                var dxAM = (sa * th) / 2;
                var dyAM = (ca * th) / 2;
                // xA, yA = upper-left outer tip (left side of arm), coordinates from (0,0)
                var xA = xM - dxAM;
                var yA = yM + dyAM;
                var xB = xM + dxAM;
                var yB = yM - dyAM;
                // yC = center notch: where the inner edge of one arm meets the inner edge of the other
                var xBC = hc - xB;
                var yBC = xBC * ta;
                var yC = yBC + yB;
                // Mirror points for upper-right quadrant
                var xD = w - xB;
                var xE = w - xA;
                // xF: where the arm inner edge meets vc (center y)
                var yFE = vc - yA;
                var xFE = yFE / ta;
                var xF = xE - xFE;
                var xL = xA + xFE;
                // Bottom half mirrors
                var yG = h - yA;
                var yH = h - yB;
                var yI = h - yC;

                return string.Join(" ", [
                    $"M{xA},{yA}",
                    $"L{xB},{yB}",
                    $"L{hc},{yC}",
                    $"L{xD},{yB}",
                    $"L{xE},{yA}",
                    $"L{xF},{vc}",
                    $"L{xE},{yG}",
                    $"L{xD},{yH}",
                    $"L{hc},{yI}",
                    $"L{xB},{yH}",
                    $"L{xA},{yG}",
                    $"L{xL},{vc}",
                    "Z",
                ]);
            });

            PresetShapes.Add("mathDivide", (w, h, adjustments) =>
            {
                var adj1 = adjustments?["adj1"] ?? 23520;
                var adj2 = adjustments?["adj2"] ?? 5880;
                var adj3 = adjustments?["adj3"] ?? 11760;
                var a1 = Math.Min(Math.Max(adj1, 1000), 36745);
                var maxAdj3 = Math.Min((73490 - a1) / 4, (36745 * w) / Math.Max(h, 1));
                var a3 = Math.Min(Math.Max(adj3, 1000), maxAdj3);
                var maxAdj2 = 73490 - 4 * a3 - a1;
                var a2 = Math.Min(Math.Max(adj2, 0), maxAdj2);
                var hc = w / 2;
                var vc = h / 2;
                var dy1 = (h * a1) / 200000.0;
                var yg = (h * a2) / 100000.0;
                var rad = (h * a3) / 100000.0;
                var dx1 = (w * 73490) / 200000.0;
                var y3 = vc - dy1;
                var y4 = vc + dy1;
                var y2 = y3 - (yg + rad);
                var y1 = y2 - rad;
                var y5 = h - y1;
                var x1 = hc - dx1;
                var x3 = hc + dx1;
                return string.Join(" ", [
                    // Top dot
                    $"M{hc + rad},{y1 + rad} A{rad},{rad} 0 1,1 {hc - rad},{y1 + rad} A{rad},{rad} 0 1,1 {hc + rad},{y1 + rad} Z",
                    // Bottom dot
                    $"M{hc + rad},{y5 - rad} A{rad},{rad} 0 1,1 {hc - rad},{y5 - rad} A{rad},{rad} 0 1,1 {hc + rad},{y5 - rad} Z",
                    // Bar
                    $"M{x1},{y3} L{x3},{y3} L{x3},{y4} L{x1},{y4} Z",
                ]);
            });

            PresetShapes.Add("mathEqual", (w, h, adjustments) =>
            {
                // OOXML: adj1=23520 (bar thickness, max 36745), adj2=11760 (gap, max 100000-2*a1)
                var adj1Raw = adjustments?["adj1"] ?? 23520;
                var adj2Raw = adjustments?["adj2"] ?? 11760;
                var a1 = Math.Min(Math.Max(adj1Raw, 0), 36745);
                var mAdj2 = 100000 - a1 * 2;
                var a2 = Math.Min(Math.Max(adj2Raw, 0), Math.Max(mAdj2, 0));
                var dy1 = (h * a1) / 100000.0;
                var dy2 = (h * a2) / 200000.0;
                var dx1 = (w * 73490) / 200000.0;
                var hc = w / 2;
                var vc = h / 2;
                var y2 = vc - dy2; // center of top bar
                var y3 = vc + dy2; // center of bottom bar
                var y1 = y2 - dy1; // top of top bar
                var y4 = y3 + dy1; // bottom of bottom bar
                var x1 = hc - dx1;
                var x2 = hc + dx1;

                return string.Join(" ", [
                    $"M{x1},{y1} L{x2},{y1} L{x2},{y2} L{x1},{y2} Z",
                    $"M{x1},{y3} L{x2},{y3} L{x2},{y4} L{x1},{y4} Z",
                ]);
            });

            PresetShapes.Add("mathNotEqual", (w, h, adjustments) =>
            {
                // Follow OOXML mathNotEqual geometry (single closed contour), which keeps
                // bar thickness/slash width and intersections aligned with PowerPoint.
                var adj1Raw = adjustments?["adj1"] ?? 23520;
                var adj2Raw = adjustments?["adj2"];
                var adj3Raw = adjustments?["adj3"] ?? 11760;
                var hc = w / 2;
                var vc = h / 2;
                var hd2 = h / 2;
                var a1 = Math.Min(Math.Max(adj1Raw, 0), 50000);

                Func<double> crAng = (() =>
                {
                    if (adj2Raw == null)
                        return (110 * Math.PI) / 180;

                    double rad = ((adj2Raw.Value / 60000) * Math.PI) / 180;
                    double min = (70 * Math.PI) / 180;
                    double max = (110 * Math.PI) / 180;

                    return Math.Min(Math.Max(rad, min), max);
                });

                var maxAdj3 = 100000 - a1 * 2;
                var a3 = Math.Min(Math.Max(adj3Raw, 0), maxAdj3);
                var dy1 = (h * a1) / 100000.0;
                var dy2 = (h * a3) / 200000.0;
                var dx1 = (w * 73490) / 200000.0;
                var x1 = hc - dx1;
                var x8 = hc + dx1;
                var y2 = vc - dy2;
                var y3 = vc + dy2;
                var y1 = y2 - dy1;
                var y4 = y3 + dy1;
                var cadj2 = crAng() - Math.PI / 2;
                var xadj2 = hd2 * Math.Tan(cadj2);
                var s = Math.Sqrt(Math.Pow(xadj2, 2) + Math.Pow(hd2, 2));
                var len = s == 0 ? 1 : s;
                var bhw = (len * dy1) / hd2;
                var bhw2 = bhw / 2;
                var x7 = hc + xadj2 - bhw2;
                var x6 = x7 - (xadj2 * y1) / hd2;
                var x5 = x7 - (xadj2 * y2) / hd2;
                var x4 = x7 - (xadj2 * y3) / hd2;
                var x3 = x7 - (xadj2 * y4) / hd2;
                var rx7 = x7 + bhw;
                var rx6 = x6 + bhw;
                var rx5 = x5 + bhw;
                var rx4 = x4 + bhw;
                var rx3 = x3 + bhw;
                var dx7 = (dy1 * hd2) / len;
                var rx = cadj2 > 0 ? x7 + dx7 : rx7;
                var lx = cadj2 > 0 ? x7 : rx7 - dx7;
                var dy3 = (dy1 * xadj2) / len;
                var ry = cadj2 > 0 ? dy3 : 0;
                var ly = cadj2 > 0 ? 0 : -dy3;
                var dlx = w - rx;
                var drx = w - lx;
                var dly = h - ry;
                var dry = h - ly;

                return string.Join(" ", [
                    $"M{x1},{y1}",
                    $"L{x6},{y1}",
                    $"L{lx},{ly}",
                    $"L{rx},{ry}",
                    $"L{rx6},{y1}",
                    $"L{x8},{y1}",
                    $"L{x8},{y2}",
                    $"L{rx5},{y2}",
                    $"L{rx4},{y3}",
                    $"L{x8},{y3}",
                    $"L{x8},{y4}",
                    $"L{rx3},{y4}",
                    $"L{drx},{dry}",
                    $"L{dlx},{dly}",
                    $"L{x3},{y4}",
                    $"L{x1},{y4}",
                    $"L{x1},{y3}",
                    $"L{x4},{y3}",
                    $"L{x5},{y2}",
                    $"L{x1},{y2}",
                    "Z",
                ]);
            });

            PresetShapes.Add("round1Rect", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 16667);
                var r = Math.Min(w, h) * a;
                return string.Join(" ", ["M0,0", $"L{w - r},0", $"A{r},{r} 0 0,1 {w},{r}", $"L{w},{h}", "L0,{h}", "Z"]);
            });

            PresetShapes.Add("round2SameRect", (w, h, adjustments) =>
            {
                var a1 = Adjust(adjustments, "adj1", 16667);
                var a2 = Adjust(adjustments, "adj2", 0);
                var r1 = Math.Min(w, h) * a1;
                var r2 = Math.Min(w, h) * a2;
                return string.Join(" ", [
                    $"M{r1},0",
                    $"L{w - r1},0",
                    $"A{r1},{r1} 0 0,1 {w},{r1}",
                    $"L{w},{h - r2}",
                    $"A{r2},{r2} 0 0,1 {w - r2},{h}",
                    $"L{r2},{h}",
                    $"A{r2},{r2} 0 0,1 0,{h - r2}",
                    $"L0,{r1}",
                    $"A{r1},{r1} 0 0,1 {r1},0",
                    "Z",
                ]);
            });

            PresetShapes.Add("round2DiagRect", (w, h, adjustments) =>
            {
                var a1 = Adjust(adjustments, "adj1", 16667);
                var a2 = Adjust(adjustments, "adj2", 0);
                var r1 = Math.Min(w, h) * a1;
                var r2 = Math.Min(w, h) * a2;
                return string.Join(" ", [
                    $"M{r1},0",
                    $"L{w},0",
                    $"L{w},{h - r2}",
                    $"A{r2},{r2} 0 0,1 {w - r2},{h}",
                    $"L0,{h}",
                    $"L0,{r1}",
                    $"A{r1},{r1} 0 0,1 {r1},0",
                    "Z",
                ]);
            });

            PresetShapes.Add("snip1Rect", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 16667);
                var d = Math.Min(w, h) * a;

                return $"M0,0 L{w - d},0 L{w},{d} L{w},{h} L0,{h} Z";
            });

            PresetShapes.Add("snip2SameRect", (w, h, adjustments) =>
            {
                var a1 = Adjust(adjustments, "adj1", 16667);
                var a2 = Adjust(adjustments, "adj2", 0);
                var d1 = Math.Min(w, h) * a1;
                var d2 = Math.Min(w, h) * a2;

                return $"M{d1},0 L{w - d1},0 L{w},{d1} L{w},{h - d2} L{w - d2},{h} L{d2},{h} L0,{h - d2} L0,{d1} Z";
            });

            PresetShapes.Add("snip2DiagRect", (w, h, adjustments) =>
            {
                // OOXML spec: diagonal snipped rectangle. adj1=top-left/bottom-right, adj2=top-right/bottom-left
                var ss = Math.Min(w, h);
                var a1 = Math.Min(Math.Max(adjustments?["adj1"] ?? 0, 0), 50000);
                var a2 = Math.Min(Math.Max(adjustments?["adj2"] ?? 16667, 0), 50000);
                var lx1 = (ss * a1) / 100000.0;
                var lx2 = w - lx1;
                var ly1 = h - lx1;
                var rx1 = (ss * a2) / 100000.0;
                var rx2 = w - rx1;
                var ry1 = h - rx1;

                return $"M{lx1},0 L{rx2},0 L{w},{rx1} L{w},{ly1} L{lx2},{h} L{rx1},{h} L0,{ry1} L0,{lx1} Z";
            });

            PresetShapes.Add("snipRoundRect", (w, h, adjustments) =>
            {
                var a1 = Adjust(adjustments, "adj1", 16667);
                var a2 = Adjust(adjustments, "adj2", 16667);
                var r = Math.Min(w, h) * a1;
                var d = Math.Min(w, h) * a2;
                return string.Join(" ", [
                    $"M{r},0",
                    $"L{w - d},0",
                    $"L{w},{d}",
                    $"L{w},{h}",
                    $"L0,{h}",
                    $"L0,{r}",
                    $"A{r},{r} 0 0,1 {r},0",
                    "Z",
                ]);
            });

            PresetShapes.Add("bevel", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 12500);
                var t = Math.Min(w, h) * a;

                return string.Join(" ", [
                    // Outer
                    $"M0,0 L{w},0 L{w},{h} L0,{h} Z",
                    // Inner
                    $"M{t},{t} L{t},{h - t} L{w - t},{h - t} L{w - t},{t} Z",
                    // Connecting triangles (top)
                    $"M0,0 L{w},0 L{w - t},{t} L{t},{t} Z",
                    // Right
                    $"M{w},0 L{w},{h} L{w - t},{h - t} L{w - t},{t} Z",
                    // Bottom
                    $"M{w},{h} L0,{h} L{t},{h - t} L{w - t},{h - t} Z",
                    // Left
                    $"M0,{h} L0,0 L{t},{t} L{t},{h - t} Z",
                ]);
            });

            PresetShapes.Add("foldedCorner", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 16667);
                var fold = Math.Min(w, h) * a * 0.7;

                return string.Join(" ", [
                    $"M0,0 L{w},0 L{w},{h} L0,{h} Z",
                    // Fold triangle
                    $"M{w - fold},{h} L{w},{h} L{w},{h - fold}",
                ]);
            });

            // smileyFace is implemented as multiPathPreset (see multiPathPresets below)
            PresetShapes.Add("sun", (w, h, adjustments) =>
            {
                // OOXML spec: adj default=25000, pinned 12500..46875
                var AdjustRaw = adjustments?["adj"] ?? 25000;
                var a = Math.Min(Math.Max(AdjustRaw, 12500), 46875);
                var g0 = 50000 - a;
                // OOXML guide formulas
                var g1 = (g0 * 30274) / 32768;
                var g2 = (g0 * 12540) / 32768;
                var _g3 = g1 + 50000;
                var _g4 = g2 + 50000;
                var g5 = 50000 - g1;
                var g6 = 50000 - g2;
                var g7 = (g0 * 23170) / 32768;
                var g8 = 50000 + g7;
                var g9 = 50000 - g7;
                var g10 = (g5 * 3) / 4;
                var g11 = (g6 * 3) / 4;
                var g12 = g10 + 3662;
                var g13 = g11 + 3662;
                var g14 = g11 + 12500;
                var g15 = 100000 - g10;
                var g16 = 100000 - g12;
                var g17 = 100000 - g13;
                var g18 = 100000 - g14;
                // Pixel coordinates
                var hc = w / 2;
                var vc = h / 2;
                var ox1 = (w * 18436) / 21600;
                var oy1 = (h * 3163) / 21600;
                var ox2 = (w * 3163) / 21600;
                var oy2 = (h * 18436) / 21600;
                Func<double, double, double> s = (pct, dim) => (dim * pct) / 100000.0;
                var _x8 = s(g8, w);
                var _x9 = s(g9, w);
                var x10 = s(g10, w);
                var x12 = s(g12, w);
                var x13 = s(g13, w);
                var x14 = s(g14, w);
                var x15 = s(g15, w);
                var x16 = s(g16, w);
                var x17 = s(g17, w);
                var x18 = s(g18, w);
                var wR = s(g0, w);
                var hR = s(g0, h);
                var _y8 = s(g8, h);
                var _y9 = s(g9, h);
                var y10 = s(g10, h);
                var y12 = s(g12, h);
                var y13 = s(g13, h);
                var y14 = s(g14, h);
                var y15 = s(g15, h);
                var y16 = s(g16, h);
                var y17 = s(g17, h);
                var y18 = s(g18, h);
                var x19 = s(a, w);

                return string.Join(" ", [
                    // Ray 0: right
                    $"M{w},{vc} L{x15},{y18} L{x15},{y14} Z",
                    // Ray 1: top-right
                    $"M{ox1},{oy1} L{x16},{y13} L{x17},{y12} Z",
                    // Ray 2: top
                    $"M{hc},0 L{x18},{y10} L{x14},{y10} Z",
                    // Ray 3: top-left
                    $"M{ox2},{oy1} L{x13},{y12} L{x12},{y13} Z",
                    // Ray 4: left
                    $"M0,{vc} L{x10},{y14} L{x10},{y18} Z",
                    // Ray 5: bottom-left
                    $"M{ox2},{oy2} L{x12},{y17} L{x13},{y16} Z",
                    // Ray 6: bottom
                    $"M{hc},{h} L{x14},{y15} L{x18},{y15} Z",
                    // Ray 7: bottom-right
                    $"M{ox1},{oy2} L{x17},{y16} L{x16},{y17} Z",
                    // Center ellipse (arcTo from x19,vc with wR,hR, startAngle=180°, sweep=360°)
                    $"M{x19},{vc}",
                    $"A{wR},{hR} 0 1,1 {x19 + 2 * wR},{vc}",
                    $"A{wR},{hR} 0 1,1 {x19},{vc}",
                    "Z",
                ]);
            });

            PresetShapes.Add("moon", (w, h, adjustments) =>
            {
                if (w <= 0 || h <= 0)
                    return $"M0,0 L{w},0 L{w},{h} L0,{h} Z";
                // OOXML moon: outer semicircle (rx=w, ry=h/2) + inner semicircle (rx=g18w, ry=dy1).
                // Both arcs share endpoints (w,0) and (w,h). Inner ellipse centered at (g0w+g18w, h/2).
                var ss = Math.Min(w, h);
                var hd2 = h / 2;
                var a = Math.Min(Math.Max(adjustments?["adj"] ?? 50000, 0), 87500);
                var g0 = (ss * a) / 100000.0;
                var g1 = ss - g0;
                if (g1 <= 0)
                    return $"M0,0 L{w},0 L{w},{h} L0,{h} Z";
                var g0w = (g0 * w) / ss;
                var g5 = (2 * ss * ss - g0 * g0) / g1;
                var g6w = ((g5 - g0) * w) / ss;
                var g8 = g5 / 2 - g0;
                var dy1 = (g8 * hd2) / ss;
                var g18w = (g6w - g0w) / 2;

                return string.Join(" ", [
                    $"M{w},{h}",
                    $"A{w},{hd2} 0 0,1 {w},0", // outer: (w,h) → left semicircle → (w,0)
                    $"A{g18w},{dy1} 0 0,0 {w},{h}", // inner: (w,0) → concave arc → (w,h)
                    "Z",
                ]);
            });

            PresetShapes.Add("lightningBolt", (w, h, a) =>
            {
                // Calibrated against OOXML preset rendering (PowerPoint PDF export):
                // the old simplified 7-point bolt was too wide and lacked the inner notches.
                // This normalized 11-point contour follows the default lightningBolt geometry.
                return string.Join(" ", [
                    $"M{w * 0.3895},{h * 0.0}",
                    $"L{w * 0.0},{h * 0.1821}",
                    $"L{w * 0.3425},{h * 0.3845}",
                    $"L{w * 0.2265},{h * 0.4452}",
                    $"L{w * 0.5497},{h * 0.6391}",
                    $"L{w * 0.453},{h * 0.683}",
                    $"L{w * 0.9972},{h * 0.9983}",
                    $"L{w * 0.6796},{h * 0.5919}",
                    $"L{w * 0.7624},{h * 0.5514}",
                    $"L{w * 0.5138},{h * 0.3153}",
                    $"L{w * 0.5939},{h * 0.2816}",
                    "Z",
                ]);
            });

            PresetShapes.Add("bracketPair", (w, h, adjustments) =>
            {
                // OOXML: adj=16667 (max 50000), radius = ss * a / 100000
                var ss = Math.Min(w, h);
                var a = Math.Min(Math.Max(AdjustRaw(adjustments, "adj", 16667), 0), 50000);
                var r = (ss * a) / 100000.0;
                var x2 = w - r;
                var y2 = h - r;

                return string.Join(" ", [
                    // Left bracket: bottom-left arc → vertical → top-left arc
                    $"M{r},{h}",
                    $"A{r},{r} 0 0,1 0,{y2}",
                    $"L0,{r}",
                    $"A{r},{r} 0 0,1 {r},0",
                    // Right bracket: top-right arc → vertical → bottom-right arc
                    $"M{x2},0",
                    $"A{r},{r} 0 0,1 {w},{r}",
                    $"L{w},{y2}",
                    $"A{r},{r} 0 0,1 {x2},{h}",
                ]);
            });

            PresetShapes.Add("bracePair", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 8333);
                var r = Math.Min(w, h) * a;
                var cy = h / 2;

                return string.Join(" ", [
                    // Left brace
                    $"M{r * 2},0",
                    $"A{r},{r} 0 0,0 {r},{r}",
                    $"L{r},{cy - r}",
                    $"A{r},{r} 0 0,1 0,{cy}",
                    $"A{r},{r} 0 0,1 {r},{cy + r}",
                    $"L{r},{h - r}",
                    $"A{r},{r} 0 0,0 {r * 2},{h}",
                    // Right brace
                    $"M{w - r * 2},0",
                    $"A{r},{r} 0 0,1 {w - r},{r}",
                    $"L{w - r},{cy - r}",
                    $"A{r},{r} 0 0,0 {w},{cy}",
                    $"A{r},{r} 0 0,0 {w - r},{cy + r}",
                    $"L{w - r},{h - r}",
                    $"A{r},{r} 0 0,1 {w - r * 2},{h}",
                ]);
            });

            PresetShapes.Add("leftBracket", (w, h, adjustments) =>
            {
                var ss = Math.Min(w, h);
                var maxAdj = ss > 0 ? (50000 * h) / ss : 0;
                var a = Math.Max(0, Math.Min(adjustments?["adj"] ?? 8333, maxAdj));
                var y1 = (ss * a) / 100000.0;
                Func<double, double> toDeg = (ooxmlAng) => ooxmlAng / 60000;
                Func<double, double, double, double, double, double, ArcFromInfo> arcFrom = (x0, y0, rx, ry, stAng, swAng) =>
                {
                    var st = (toDeg(stAng) * Math.PI) / 180;
                    var sw = (toDeg(swAng) * Math.PI) / 180;
                    var cx = x0 - rx * Math.Cos(st);
                    var cy = y0 - ry * Math.Sin(st);
                    var x1 = cx + rx * Math.Cos(st + sw);
                    var y1p = cy + ry * Math.Sin(st + sw);
                    var large = Math.Abs(toDeg(swAng)) > 180 ? 1 : 0;
                    var sweep = swAng >= 0 ? 1 : 0;

                    return new ArcFromInfo() { Cmd = $"A{rx},{ry} 0 {large},{sweep} {x1},{y1p}", X = x1, Y = y1p };
                };

                var a1 = arcFrom(w, h, w, y1, 5400000, 5400000); // cd4, cd4
                var a2 = arcFrom(0, y1, w, y1, 10800000, 5400000); // cd2, cd4

                return string.Join(" ", [$"M{w},{h}", a1.Cmd, $"L0,{y1}", a2.Cmd]);
            });

            PresetShapes.Add("rightBracket", (w, h, adjustments) =>
            {
                var ss = Math.Min(w, h);
                var maxAdj = ss > 0 ? (50000 * h) / ss : 0;
                var a = Math.Max(0, Math.Min(adjustments?["adj"] ?? 8333, maxAdj));
                var y1 = (ss * a) / 100000.0;
                var y2 = h - y1;
                Func<double, double> toDeg = (ooxmlAng) => ooxmlAng / 60000;
                Func<double, double, double, double, double, double, ArcFromInfo> arcFrom = (x0, y0, rx, ry, stAng, swAng) =>
                {
                    var st = (toDeg(stAng) * Math.PI) / 180;
                    var sw = (toDeg(swAng) * Math.PI) / 180;
                    var cx = x0 - rx * Math.Cos(st);
                    var cy = y0 - ry * Math.Sin(st);
                    var x1 = cx + rx * Math.Cos(st + sw);
                    var y1p = cy + ry * Math.Sin(st + sw);
                    var large = Math.Abs(toDeg(swAng)) > 180 ? 1 : 0;
                    var sweep = swAng >= 0 ? 1 : 0;

                    return new ArcFromInfo() { Cmd = $"A{rx},{ry} 0 {large},{sweep} {x1},{y1p}", X = x1, Y = y1p };
                };

                var a1 = arcFrom(0, 0, w, y1, 16200000, 5400000); // 3cd4, cd4
                var a2 = arcFrom(w, y2, w, y1, 0, 5400000); // 0, cd4

                return string.Join(" ", ["M0,0", a1.Cmd, $"L{w},{y2}", a2.Cmd]);
            });

            PresetShapes.Add("leftBrace", (w, h, adjustments) =>
            {
                var ss = Math.Min(w, h);
                var a2 = Math.Max(0, Math.Min(adjustments?["adj2"] ?? 50000, 100000));
                var q1 = 100000 - a2;
                var q2 = Math.Min(q1, a2);
                var q3 = q2 / 2;
                var maxAdj1 = ss > 0 ? (q3 * h) / ss : 0;
                var a1 = Math.Max(0, Math.Min(adjustments?["adj1"] ?? 8333, maxAdj1));
                var y1 = (ss * a1) / 100000.0;
                var y3 = (h * a2) / 100000.0;
                var y4 = y3 + y1;
                var wd2 = w / 2;
                var hc = w / 2;
                Func<double, double> toDeg = (ooxmlAng) => ooxmlAng / 60000;
                Func<double, double, double, double, double, double, ArcFromInfo> arcFrom = (x0, y0, rx, ry, stAng, swAng) =>
                {
                    var st = (toDeg(stAng) * Math.PI) / 180;
                    var sw = (toDeg(swAng) * Math.PI) / 180;
                    var cx = x0 - rx * Math.Cos(st);
                    var cy = y0 - ry * Math.Sin(st);
                    var x1 = cx + rx * Math.Cos(st + sw);
                    var y1p = cy + ry * Math.Sin(st + sw);
                    var large = Math.Abs(toDeg(swAng)) > 180 ? 1 : 0;
                    var sweep = swAng >= 0 ? 1 : 0;

                    return new ArcFromInfo() { Cmd = $"A{rx},{ry} 0 {large},{sweep} {x1},{y1p}", X = x1, Y = y1p };
                };

                var x = w;
                var y = h;
                var aTop = arcFrom(x, y, wd2, y1, 5400000, 5400000); // cd4, cd4
                x = aTop.X;
                y = aTop.Y;
                var aMid1 = arcFrom(hc, y4, wd2, y1, 0, -5400000);
                var aMid2 = arcFrom(aMid1.X, aMid1.Y, wd2, y1, 5400000, -5400000);
                var aBot = arcFrom(hc, y1, wd2, y1, 10800000, 5400000); // cd2, cd4

                return string.Join(" ", [
                    $"M{w},{h}",
                    aTop.Cmd,
                    $"L{hc},{y4}",
                    aMid1.Cmd,
                    aMid2.Cmd,
                    $"L{hc},{y1}",
                    aBot.Cmd,
                ]);
            });

            PresetShapes.Add("rightBrace", (w, h, adjustments) =>
            {
                var ss = Math.Min(w, h);
                var a2 = Math.Max(0, Math.Min(adjustments?["adj2"] ?? 50000, 100000));
                var q1 = 100000 - a2;
                var q2 = Math.Min(q1, a2);
                var q3 = q2 / 2;
                var maxAdj1 = ss > 0 ? (q3 * h) / ss : 0;
                var a1 = Math.Max(0, Math.Min(adjustments?["adj1"] ?? 8333, maxAdj1));
                var y1 = (ss * a1) / 100000.0;
                var y3 = (h * a2) / 100000.0;
                var y2 = y3 - y1;
                var y4 = h - y1;
                var wd2 = w / 2;
                var hc = w / 2;
                Func<double, double> toDeg = (ooxmlAng) => ooxmlAng / 60000;
                Func<double, double, double, double, double, double, ArcFromInfo> arcFrom = (x0, y0, rx, ry, stAng, swAng) =>
                {
                    var st = (toDeg(stAng) * Math.PI) / 180;
                    var sw = (toDeg(swAng) * Math.PI) / 180;
                    var cx = x0 - rx * Math.Cos(st);
                    var cy = y0 - ry * Math.Sin(st);
                    var x1 = cx + rx * Math.Cos(st + sw);
                    var y1p = cy + ry * Math.Sin(st + sw);
                    var large = Math.Abs(toDeg(swAng)) > 180 ? 1 : 0;
                    var sweep = swAng >= 0 ? 1 : 0;

                    return new ArcFromInfo() { Cmd = $"A{rx},{ry} 0 {large},{sweep} {x1},{y1p}", X = x1, Y = y1p };
                };

                var aTop = arcFrom(0, 0, wd2, y1, 16200000, 5400000); // 3cd4, cd4
                var aMid1 = arcFrom(hc, y2, wd2, y1, 10800000, -5400000); // cd2,-cd4
                var aMid2 = arcFrom(aMid1.X, aMid1.Y, wd2, y1, 16200000, -5400000); //3cd4,-cd4
                var aBot = arcFrom(hc, y4, wd2, y1, 0, 5400000); //0,cd4

                return string.Join(" ", ["M0,0", aTop.Cmd, $"L{hc},{y2}", aMid1.Cmd, aMid2.Cmd, $"L{hc},{y4}", aBot.Cmd]);
            });

            // ==== Action Buttons ====
            // Action buttons are multi-path shapes: background rect + icon with darken fill + icon outline + rect outline.
            // OOXML spec uses ss*3/8 as the icon half-size (dx2), with the icon centred at (hc, vc).
            // Shapes with multiPathPresets entries below get proper 3D treatment. Remaining shapes
            // fall back to the legacy actionButtonIcons overlay (single flat icon path).
            PresetShapes.Add("actionButtonBlank", (w, h, a) => $"M0,0 L{w},0 L{w},{h} L0,{h} Z");

            // ==== Aliases and common alternative names ====
            // Some shapes are known by multiple names in different OOXML versions
            // flowChartOfflineStorage: registered as multiPathPreset (see below)
            // ribbon is implemented as multiPathPreset (see multiPathPresets below)
            PresetShapes.Add("wave", (w, h, adjustments) =>
            {
                // OOXML: adj1=12500 (max 20000), adj2=0 (phase shift, range -10000..10000)
                var a1 = Math.Min(Math.Max(AdjustRaw(adjustments, "adj1", 12500), 0), 20000);
                var a2 = Math.Min(Math.Max(AdjustRaw(adjustments, "adj2", 0), -10000), 10000);
                var y1 = (h * a1) / 100000.0;
                var dy2 = (y1 * 10) / 3;
                var y2 = y1 - dy2; // control above crest
                var y3 = y1 + dy2; // control below crest
                var y4 = h - y1; // bottom wave y
                var y5 = y4 - dy2;
                var y6 = y4 + dy2;
                // Phase shift
                var of2 = (w * a2) / 50000.0;
                var dx2 = of2 < 0 ? 0 : of2;
                var dx5 = of2 < 0 ? of2 : 0;
                var x2 = -dx2;
                var x5 = w - dx5;
                var dx3 = (x5 - x2) / 3;
                var x3 = x2 + dx3;
                var x4 = (x3 + x5) / 2;
                var x6 = dx5;
                var x10 = w + dx2;
                var x7 = x6 + (x10 - x6) / 3;
                var x8 = (x7 + x10) / 2;

                return string.Join(" ", [
                    $"M{x2},{y1}",
                    $"C{x3},{y2} {x4},{y3} {x5},{y1}",
                    $"L{x10},{y4}",
                    $"C{x8},{y6} {x7},{y5} {x6},{y4}",
                    "Z",
                ]);
            });

            PresetShapes.Add("doubleWave", (w, h, adjustments) =>
            {
                // OOXML: adj1=6250 (max 12500), adj2=0 (phase shift)
                var a1 = Math.Min(Math.Max(AdjustRaw(adjustments, "adj1", 6250), 0), 12500);
                var a2 = Math.Min(Math.Max(AdjustRaw(adjustments, "adj2", 0), -10000), 10000);
                var y1 = (h * a1) / 100000.0;
                var dy2 = (y1 * 10) / 3;
                var y2 = y1 - dy2;
                var y3 = y1 + dy2;
                var y4 = h - y1;
                var y5 = y4 - dy2;
                var y6 = y4 + dy2;
                var of2 = (w * a2) / 50000.0;
                var dx2 = of2 < 0 ? 0 : of2;
                var dx8 = of2 < 0 ? of2 : 0;
                var x2 = -dx2;
                var x8 = w - dx8;
                var dx3 = (x8 - x2) / 6;
                var x3 = x2 + dx3;
                var dx4 = (x8 - x2) / 3;
                var x4 = x2 + dx4;
                var x5 = (x2 + x8) / 2;
                var x6 = x5 + dx3;
                var x7 = (x6 + x8) / 2;
                var x9 = dx8;
                var x15 = w + dx2;
                var dx3b = (x15 - x9) / 6;
                var x10 = x9 + dx3b;
                var x11 = x9 + (x15 - x9) / 3;
                var x12 = (x9 + x15) / 2;
                var x13 = x12 + dx3b;
                var x14 = (x13 + x15) / 2;

                return string.Join(" ", [
                    $"M{x2},{y1}",
                    $"C{x3},{y2} {x4},{y3} {x5},{y1}",
                    $"C{x6},{y2} {x7},{y3} {x8},{y1}",
                    $"L{x15},{y4}",
                    $"C{x14},{y6} {x13},{y5} {x12},{y4}",
                    $"C{x11},{y6} {x10},{y5} {x9},{y4}",
                    "Z",
                ]);
            });

            // verticalScroll and horizontalScroll are implemented as multi-path presets
            // (see multiPathPresets below) for accurate OOXML rendering with darkenLess shadows.
            PresetShapes.Add("irregularSeal1", (w, h, a) =>
            {
                // OOXML spec: exact coordinates on 21600x21600 grid
                Func<double, double> sx = (x) => (w * x) / 21600;
                Func<double, double> sy = (y) => (h * y) / 21600;

                return string.Join(" ", [
                    $"M{sx(10800)},{sy(5800)}",
                    $"L{sx(14522)},0",
                    $"L{sx(14155)},{sy(5325)}",
                    $"L{sx(18380)},{sy(4457)}",
                    $"L{sx(16702)},{sy(7315)}",
                    $"L{sx(21097)},{sy(8137)}",
                    $"L{sx(17607)},{sy(10475)}",
                    $"L{sx(21600)},{sy(13290)}",
                    $"L{sx(16837)},{sy(12942)}",
                    $"L{sx(18145)},{sy(18095)}",
                    $"L{sx(14020)},{sy(14457)}",
                    $"L{sx(13247)},{sy(19737)}",
                    $"L{sx(10532)},{sy(14935)}",
                    $"L{sx(8485)},{sy(21600)}",
                    $"L{sx(7715)},{sy(15627)}",
                    $"L{sx(4762)},{sy(17617)}",
                    $"L{sx(5667)},{sy(13937)}",
                    $"L{sx(135)},{sy(14587)}",
                    $"L{sx(3722)},{sy(11775)}",
                    $"L0,{sy(8615)}",
                    $"L{sx(4627)},{sy(7617)}",
                    $"L{sx(370)},{sy(2295)}",
                    $"L{sx(7312)},{sy(6320)}",
                    $"L{sx(8352)},{sy(2295)}",
                    "Z",
                ]);
            });

            PresetShapes.Add("irregularSeal2", (w, h, a) =>
            {
                // Office-like irregularSeal2 coordinates (21600 design grid).
                return string.Join(" ", [
                    $"M{(w * 11462) / 21600},{(h * 4342) / 21600}",
                    $"L{(w * 14790) / 21600},0",
                    $"L{(w * 14525) / 21600},{(h * 5777) / 21600}",
                    $"L{(w * 18007) / 21600},{(h * 3172) / 21600}",
                    $"L{(w * 16380) / 21600},{(h * 6532) / 21600}",
                    $"L{w},{(h * 6645) / 21600}",
                    $"L{(w * 16985) / 21600},{(h * 9402) / 21600}",
                    $"L{(w * 18270) / 21600},{(h * 11290) / 21600}",
                    $"L{(w * 16380) / 21600},{(h * 12310) / 21600}",
                    $"L{(w * 18877) / 21600},{(h * 15632) / 21600}",
                    $"L{(w * 14640) / 21600},{(h * 14350) / 21600}",
                    $"L{(w * 14942) / 21600},{(h * 17370) / 21600}",
                    $"L{(w * 12180) / 21600},{(h * 15935) / 21600}",
                    $"L{(w * 11612) / 21600},{(h * 18842) / 21600}",
                    $"L{(w * 9872) / 21600},{(h * 17370) / 21600}",
                    $"L{(w * 8700) / 21600},{(h * 19712) / 21600}",
                    $"L{(w * 7527) / 21600},{(h * 18125) / 21600}",
                    $"L{(w * 4917) / 21600},{h}",
                    $"L{(w * 4805) / 21600},{(h * 18240) / 21600}",
                    $"L{(w * 1285) / 21600},{(h * 17825) / 21600}",
                    $"L{(w * 3330) / 21600},{(h * 15370) / 21600}",
                    $"L0,{(h * 12877) / 21600}",
                    $"L{(w * 3935) / 21600},{(h * 11592) / 21600}",
                    $"L{(w * 1172) / 21600},{(h * 8270) / 21600}",
                    $"L{(w * 5372) / 21600},{(h * 7817) / 21600}",
                    $"L{(w * 4502) / 21600},{(h * 3625) / 21600}",
                    $"L{(w * 8550) / 21600},{(h * 6382) / 21600}",
                    $"L{(w * 9722) / 21600},{(h * 1887) / 21600}",
                    "Z",
                ]);
            });

            PresetShapes.Add("teardrop", (w, h, a) =>
            {
                var rx = w / 2;
                var ry = h / 2;

                return string.Join(" ", [$"M{w},{ry}", $"A{rx},{ry} 0 1,1 {rx},0", $"L{w},0", $"L{w},{ry}", "Z"]);
            });

            PresetShapes.Add("pie", (w, h, adjustments) =>
            {
                // OOXML pie: adj1 = start angle, adj2 = end angle (60000ths of a degree). Sweep clockwise from start to end.
                // OOXML angles are "visual" (geometric) — must convert to parametric for ellipses (rx≠ry).
                var adj1Raw = adjustments?["adj1"] ?? 0;
                var adj2Raw = adjustments?["adj2"] ?? 16200000; // 270° end default
                var startDeg = (adj1Raw / 60000) % 360;
                var endDeg = (adj2Raw / 60000) % 360;
                var sweepDeg = (((endDeg - startDeg) % 360) + 360) % 360;
                if (sweepDeg == 0 && startDeg != endDeg)
                    sweepDeg = 360;
                var rx = w / 2;
                var ry = h / 2;
                Func<double, double> toRad = (d) => (d * Math.PI) / 180;
                Func<double, double> visualToParam = (deg) => Math.Atan2(Math.Sin(toRad(deg)) / ry, Math.Cos(toRad(deg)) / rx);
                var startParam = visualToParam(startDeg);
                var endParam = visualToParam(endDeg);
                var x1 = rx + rx * Math.Cos(startParam);
                var y1 = ry + ry * Math.Sin(startParam);
                var x2 = rx + rx * Math.Cos(endParam);
                var y2 = ry + ry * Math.Sin(endParam);
                var largeArc = sweepDeg > 180 ? 1 : 0;

                return string.Join(" ", [$"M{rx},{ry}", $"L{x1},{y1}", $"A{rx},{ry} 0 {largeArc},1 {x2},{y2}", "Z"]);
            });

            PresetShapes.Add("pieWedge", (w, h, a) =>
            {
                // OOXML: Quarter-ellipse pie wedge. Center at (w, h), radii = (w, h).
                // Arc from 180° sweeping 90° CW: starts at (0, h), ends at (w, 0).
                // The arc bulges toward the upper-left.
                return string.Join(" ", [$"M0,{h}", $"A{w},{h} 0 0,1 {w},0", $"L{w},{h}", "Z"]);
            });

            PresetShapes.Add("arc", (w, h, adjustments) =>
            {
                // OOXML arc: adj1/adj2 are angles in 60000ths of a degree
                // OOXML angles are "visual" (geometric) — must convert to parametric for ellipses (rx≠ry).
                var adj1Raw = adjustments?["adj1"] ?? 16200000; // default 270°
                var adj2Raw = adjustments?["adj2"] ?? 0; // default 0°
                var startDeg = adj1Raw / 60000;
                var endDeg = adj2Raw / 60000;
                var rx = w / 2;
                var ry = h / 2;
                Func<double, double> toRad = (d) => (d * Math.PI) / 180;
                Func<double, double> visualToParam = (deg) => Math.Atan2(Math.Sin(toRad(deg)) / ry, Math.Cos(toRad(deg)) / rx);
                var startParam = visualToParam(startDeg);
                var endParam = visualToParam(endDeg);
                var x1 = rx + rx * Math.Cos(startParam);
                var y1 = ry + ry * Math.Sin(startParam);
                var x2 = rx + rx * Math.Cos(endParam);
                var y2 = ry + ry * Math.Sin(endParam);
                var sweepDeg = (((endDeg - startDeg) % 360) + 360) % 360;
                if (sweepDeg == 0 && startDeg != endDeg)
                    sweepDeg = 360;
                var largeArc = sweepDeg > 180 ? 1 : 0;

                return $"M{x1},{y1} A{rx},{ry} 0 {largeArc},1 {x2},{y2}";
            });

            PresetShapes.Add("chord", (w, h, adjustments) =>
            {
                // OOXML chord: arc + chord line. Spec uses ellipse (arcTo wR="wd2" hR="hd2") per presetShapeDefinitions.
                // OOXML angles are "visual" (geometric) angles — the angle of the ray from center to the point.
                // For ellipses (rx≠ry), convert to parametric angle: t = atan2(sin(θ)/ry, cos(θ)/rx)
                var adj1Raw = adjustments?["adj1"] ?? 2700000; // default 45°
                var adj2Raw = adjustments?["adj2"] ?? 16200000; // default 270°
                var startDeg = adj1Raw / 60000;
                var endDeg = adj2Raw / 60000;
                var cx = w / 2;
                var cy = h / 2;
                var rx = w / 2;
                var ry = h / 2;
                Func<double, double> toRad = (d) => (d * Math.PI) / 180;
                // Convert OOXML visual angles to parametric angles on the ellipse
                Func<double, double> visualToParam = (deg) => Math.Atan2(Math.Sin(toRad(deg)) / ry, Math.Cos(toRad(deg)) / rx);
                var startParam = visualToParam(startDeg);
                var endParam = visualToParam(endDeg);
                var x1 = cx + rx * Math.Cos(startParam);
                var y1 = cy + ry * Math.Sin(startParam);
                var x2 = cx + rx * Math.Cos(endParam);
                var y2 = cy + ry * Math.Sin(endParam);
                // Use OOXML visual sweep to determine large-arc-flag
                var sweepDeg = (((endDeg - startDeg) % 360) + 360) % 360;
                if (sweepDeg == 0 && startDeg != endDeg)
                    sweepDeg = 360;
                // When adj1 == adj2, the chord covers the full ellipse (360° sweep)
                if (sweepDeg == 0)
                {
                    return $"M{cx - rx},{cy} A{rx},{ry} 0 1,1 {cx + rx},{cy} A{rx},{ry} 0 1,1 {cx - rx},{cy} Z";
                }
                var largeArc = sweepDeg > 180 ? 1 : 0;

                return $"M{x1},{y1} A{rx},{ry} 0 {largeArc},1 {x2},{y2} Z";
            });

            PresetShapes.Add("funnel", (w, h, a) =>
            {
                // OOXML funnel: top rim ellipse arc + tapered sides + bottom spout arc + inset top ellipse.
                // From presetShapeDefinitions.xml (ECMA-376).
                var ss = Math.Min(w, h);
                var wd2 = w / 2;
                var hd4 = h / 4;
                var hc = w / 2;
                var b = h;
                var d = ss / 20; // inset margin
                var rw2 = wd2 - d; // inset top-ellipse x-radius
                var rh2 = hd4 - d; // inset top-ellipse y-radius
                                   // Angle (in radians) where funnel sides are tangent to top ellipse.
                                   // OOXML: t1 = cos(wd2, 480000), t2 = sin(hd4, 480000) → da = atan2(t1, t2)
                                   // 480000 = 8° in 60000ths of a degree
                var ang8 = (8 * Math.PI) / 180;
                var t1 = wd2 * Math.Cos(ang8);
                var t2 = hd4 * Math.Sin(ang8);
                var da = Math.Atan2(t2, t1); // radians
                                             // Angles for the top rim arc (OOXML convention: sweep from stAng1 by swAng1)
                var stAng1 = Math.PI - da; // cd2 - da
                var swAng1 = Math.PI + 2 * da; // cd2 + 2*da
                                               // Sweep for the bottom spout arc
                var swAng3 = Math.PI - 2 * da; // cd2 - 2*da
                                               // Bottom spout ellipse radii: 1/4 of top ellipse
                var rw3 = wd2 / 4;
                var rh3 = hd4 / 4;
                // Start point on top ellipse at stAng1 (visual angle → ellipse point)
                // OOXML uses: n = (wR*hR) / mod(cos(hR,ang), sin(wR,ang), 0), then x = hc + cos(n,ang), y = hd4 + sin(n,ang)
                // This is equivalent to the parametric ellipse point at the "visual" angle.
                var ct1 = hd4 * Math.Cos(stAng1);
                var st1 = wd2 * Math.Sin(stAng1);
                var m1 = Math.Sqrt(ct1 * ct1 + st1 * st1);
                var n1 = (wd2 * hd4) / m1;
                var dx1 = n1 * Math.Cos(stAng1);
                var dy1 = n1 * Math.Sin(stAng1);
                var x1 = hc + dx1;
                var y1 = hd4 + dy1;
                // End point of top arc (at stAng1 + swAng1 = pi + da)
                var endAng1 = stAng1 + swAng1;
                var ct1e = hd4 * Math.Cos(endAng1);
                var st1e = wd2 * Math.Sin(endAng1);
                var m1e = Math.Sqrt(ct1e * ct1e + st1e * st1e);
                var n1e = (wd2 * hd4) / m1e;
                var dx1e = n1e * Math.Cos(endAng1);
                var dy1e = n1e * Math.Sin(endAng1);
                var x1e = hc + dx1e;
                var y1e = hd4 + dy1e;
                // Point on spout ellipse at angle da
                var vc3 = b - rh3; // vertical center of spout ellipse
                var ct3 = rh3 * Math.Cos(da);
                var st3 = rw3 * Math.Sin(da);
                var m3 = Math.Sqrt(ct3 * ct3 + st3 * st3);
                var n3 = (rw3 * rh3) / m3;
                var dx3 = n3 * Math.Cos(da);
                var dy3 = n3 * Math.Sin(da);
                var x3 = hc + dx3;
                var y2 = vc3 + dy3;
                // End point of spout arc (at da + swAng3)
                var endAng3 = da + swAng3;
                var ct3e = rh3 * Math.Cos(endAng3);
                var st3e = rw3 * Math.Sin(endAng3);
                var m3e = Math.Sqrt(ct3e * ct3e + st3e * st3e);
                var n3e = (rw3 * rh3) / m3e;
                var dx3e = n3e * Math.Cos(endAng3);
                var dy3e = n3e * Math.Sin(endAng3);
                var x3e = hc + dx3e;
                var y2e = vc3 + dy3e;
                // Determine arc flags
                var swDeg1 = (swAng1 * 180) / Math.PI;
                var largeArc1 = Math.Abs(swDeg1) > 180 ? 1 : 0;
                var sweep1 = swAng1 > 0 ? 1 : 0;
                var swDeg3 = (swAng3 * 180) / Math.PI;
                var largeArc3 = Math.Abs(swDeg3) > 180 ? 1 : 0;
                var sweep3 = swAng3 > 0 ? 1 : 0;
                // Sub-path 1: Funnel body (top arc → line to spout → spout arc → close)
                var body = string.Join(" ", [
                    $"M{x1},{y1}",
                    $"A{wd2},{hd4} 0 {largeArc1},{sweep1} {x1e},{y1e}",
                    $"L{x3},{y2}",
                    $"A{rw3},{rh3} 0 {largeArc3},{sweep3} {x3e},{y2e}",
                    "Z",
                    ]);
                // Sub-path 2: Inset top ellipse (full ellipse, counter-clockwise for even-odd hole)
                var x2 = wd2 - rw2; // leftmost point of inset ellipse
                var x2r = wd2 + rw2; // rightmost point
                var inset = string.Join(" ", [
                    $"M{x2},{hd4}",
                    $"A{rw2},{rh2} 0 1,0 {x2r},{hd4}",
                    $"A{rw2},{rh2} 0 1,0 {x2},{hd4}",
                    "Z",
                ]);

                return $"{body} {inset}";
            });
        }

        private static void InitactionButtonIcons()
        {
            // actionButtonHome icon removed — uses multiPathPresets entry below
            ActionButtonIcons.Add("actionButtonForwardNext", (w, h) =>
            {
                // Right-pointing triangle (▶)
                double cx = w / 2, cy = h / 2, s = Math.Min(w, h) * 0.3;

                return $"M{cx - s * 0.5},{cy - s} L{cx + s},{cy} L{cx - s * 0.5},{cy + s} Z";
            });

            ActionButtonIcons.Add("actionButtonBackPrevious", (w, h) =>
            {
                // Left-pointing triangle (◀)
                double cx = w / 2, cy = h / 2, s = Math.Min(w, h) * 0.3;

                return $"M{cx + s * 0.5},{cy - s} L{cx - s},{cy} L{cx + s * 0.5},{cy + s} Z";
            });

            ActionButtonIcons.Add("actionButtonReturn", (w, h) =>
            {
                // Curved return arrow (↩) — shaft goes right at bottom, curves UP at right end,
                // returns left at top with arrowhead pointing left (standard PowerPoint icon).
                double cx = w / 2, cy = h / 2, s = Math.Min(w, h) * 0.28;
                var thick = s * 0.22; // shaft thickness
                var bottomY = cy + s * 0.4;
                var topY = cy - s * 0.4;
                var leftX = cx - s * 0.6;
                var rightX = cx + s * 0.6;
                var r = (bottomY - topY) / 2; // semicircle radius

                return string.Join(" ", [
                    // Outer edge: bottom-left → right → arc up → left to arrowhead junction
                    $"M{leftX},{bottomY}",
                    $"L{rightX},{bottomY}",
                    $"A{r},{r} 0 0,1 {rightX},{topY}",
                    $"L{leftX + s * 0.15},{topY}",
                    // Inner edge: top → right → arc down → bottom-left
                    $"L{leftX + s * 0.15},{topY + thick}",
                    $"L{rightX - thick * 0.3},{topY + thick}",
                    $"A{r - thick},{r - thick} 0 0,0 {rightX - thick * 0.3},{bottomY - thick}",
                    $"L{leftX},{bottomY - thick}",
                    "Z",
                    // Arrowhead pointing left at top-left
                    $"M{leftX - s * 0.3},{topY + thick / 2}",
                    $"L{leftX + s * 0.15},{topY - s * 0.2}",
                    $"L{leftX + s * 0.15},{topY + thick + s * 0.2}",
                    "Z",
                ]);
            });

            ActionButtonIcons.Add("actionButtonBeginning", (w, h) =>
            {
                // Skip-to-beginning (|◀)
                double cx = w / 2, cy = h / 2, s = Math.Min(w, h) * 0.28;

                return string.Join(" ", [
                    // Left bar
                    $"M{cx - s},{cy - s} L{cx - s + s * 0.2},{cy - s} L{cx - s + s * 0.2},{cy + s} L{cx - s},{cy + s} Z",
                    // Left-pointing triangle
                    $"M{cx + s},{cy - s} L{cx - s + s * 0.35},{cy} L{cx + s},{cy + s} Z",
                ]);
            });

            ActionButtonIcons.Add("actionButtonEnd", (w, h) =>
            {
                // Skip-to-end (▶|)
                double cx = w / 2, cy = h / 2, s = Math.Min(w, h) * 0.28;

                return string.Join(" ", [
                    // Right bar
                    $"M{cx + s - s * 0.2},{cy - s} L{cx + s},{cy - s} L{cx + s},{cy + s} L{cx + s - s * 0.2},{cy + s} Z",
                    // Right-pointing triangle
                    $"M{cx - s},{cy - s} L{cx + s - s * 0.35},{cy} L{cx - s},{cy + s} Z",
                ]);
            });

            // actionButtonHelp icon removed — uses multiPathPresets entry below
            ActionButtonIcons.Add("actionButtonInformation", (w, h) =>
            {
                // Info icon (i)
                double cx = w / 2, cy = h / 2, s = Math.Min(w, h) * 0.28;

                return string.Join(" ", [
                    // Dot
                    $"M{cx - s * 0.1},{cy - s * 0.65} L{cx + s * 0.1},{cy - s * 0.65} L{cx + s * 0.1},{cy - s * 0.4} L{cx - s * 0.1},{cy - s * 0.4} Z",
                    // Stem
                    $"M{cx - s * 0.12},{cy - s * 0.2} L{cx + s * 0.12},{cy - s * 0.2} L{cx + s * 0.12},{cy + s * 0.65} L{cx - s * 0.12},{cy + s * 0.65} Z",
                ]);
            });

            ActionButtonIcons.Add("actionButtonDocument", (w, h) =>
            {
                // Document with folded corner
                double cx = w / 2, cy = h / 2, s = Math.Min(w, h) * 0.28;
                double dx = s * 0.7, dy = s, fold = s * 0.3;

                return string.Join(" ", [
                    $"M{cx - dx},{cy - dy}",
                    $"L{cx + dx - fold},{cy - dy} L{cx + dx},{cy - dy + fold}",
                    $"L{cx + dx},{cy + dy} L{cx - dx},{cy + dy} Z",
                    $"M{cx + dx - fold},{cy - dy} L{cx + dx - fold},{cy - dy + fold} L{cx + dx},{cy - dy + fold}",
                ]);
            });
        }

        /// <summary>
        /// actionButtonSound icon removed — uses multiPathPresets entry below
        /// actionButtonMovie icon is now rendered via multiPathPresets (see below).
        /// Get the SVG path for the icon overlay of an action button.
        /// </summary>
        /// <param name="shapeType"></param>
        /// <param name="w"></param>
        /// <param name="h"></param>
        /// <returns> Returns null if the shape is not an action button or is actionButtonBlank.</returns>
        public static string GetActionButtonIconPath(string shapeType, double w, double h)
        {
            string key = ActionButtonIcons.FirstOrDefault(item => item.Key.ToLower() == shapeType.ToLower()).Key;

            var generator = key != null ? ActionButtonIcons[key] : null;

            return generator != null ? generator(w, h) : null;
        }

        private static void InitPresetOverlays()
        {
            PresetOverlays.Add("can", (w, h, a) =>
            {
                var ry = h * 0.1;
                var rx = w / 2;

                return new List<PresetOverlayInfo>()
                {
                         new PresetOverlayInfo()
                         {
                            Path= [$"M0,{ry}", $"A{rx},{ry} 0 0,1 {w},{ry}", $"A{rx},{ry} 0 0,1 0,{ry}", "Z"],
                            FillModifier = "lighten"
                         }
                };
            });
        }

        private static void InitMultiPathPresets()
        {
            Func<double, double, string> _rect = (w, h) => $"M0,0 L{w},0 L{w},{h} L0,{h} Z";

            // actionButtonForwardNext (VBA 0130): right-pointing triangle ▶
            MultiPathPresets.Add("actionButtonForwardNext", (w, h, a) =>
                {
                    var info = AbGuides(w, h);
                    var tri = $"M{info.G12},{info.VC} L{info.G11},{info.G9} L{info.G11},{info.G10} Z";

                    return new List<ArrowPathInfo>() {
                     new ArrowPathInfo(){ D = $"{_rect(w, h)} {tri}", Fill= "norm", Stroke= false },
                     new ArrowPathInfo(){ D= tri, Fill= "darken", Stroke= false },
                     new ArrowPathInfo(){ D= tri, Fill= "none", Stroke= true },
                     new ArrowPathInfo(){ D= _rect(w, h), Fill="none", Stroke= true },
                    };
                });

            MultiPathPresets.Add("actionButtonForward", (w, h, a) =>
            {
                var forwardNext = MultiPathPresets["actionButtonForwardNext"];

                return forwardNext != null ? forwardNext(w, h, a) : new List<ArrowPathInfo>();
            });

            // actionButtonBackPrevious (VBA 0129): left-pointing triangle ◀
            MultiPathPresets.Add("actionButtonBackPrevious", (w, h, a) =>
            {
                var info = AbGuides(w, h);
                var tri = $"M{info.G11},{info.VC} L{info.G12},{info.G9} L{info.G12},{info.G10} Z";

                return new List<ArrowPathInfo>() {
                    new ArrowPathInfo() { D= $"{_rect(w, h)} {tri}", Fill= "norm", Stroke= false },
                    new ArrowPathInfo(){ D= tri, Fill= "darken", Stroke= false },
                    new ArrowPathInfo(){ D= tri, Fill= "none", Stroke= true },
                    new ArrowPathInfo() { D= _rect(w, h), Fill= "none", Stroke= true }
                };
            });

            // actionButtonBeginning (VBA 0131): |◀ skip-to-start
            MultiPathPresets.Add("actionButtonBeginning", (w, h, a) =>
            {
                var info = AbGuides(w, h);
                var g9 = info.G9;
                var g10 = info.G10;
                var g11 = info.G11;
                var g12 = info.G12;
                var g13 = info.G13;
                var vc = info.VC;
                double g14 = g13 / 8, g15 = g13 / 4;
                double g16 = g11 + g14, g17 = g11 + g15;
                var tri = $"M{g17},{vc} L{g12},{g9} L{g12},{g10} Z";
                var bar = $"M{g16},{g9} L{g11},{g9} L{g11},{g10} L{g16},{g10} Z";
                var icon = $"{tri} {bar}";

                return new List<ArrowPathInfo>() {
                            new ArrowPathInfo(){ D= $"{_rect(w, h)} {icon}", Fill= "norm", Stroke= false },
                            new ArrowPathInfo(){ D= icon, Fill= "darken", Stroke= false },
                            new ArrowPathInfo(){ D= icon, Fill= "none", Stroke= true },
                            new ArrowPathInfo(){ D= _rect(w, h), Fill= "none", Stroke= true }
                };
            });

            // actionButtonEnd (VBA 0132): ▶| skip-to-end
            MultiPathPresets.Add("actionButtonEnd", (w, h, a) =>
            {
                var info = AbGuides(w, h);
                var g9 = info.G9;
                var g10 = info.G10;
                var g11 = info.G11;
                var g12 = info.G12;
                var g13 = info.G13;
                var hc = info.HC;
                var vc = info.VC;
                double g14 = (g13 * 3) / 4, g15 = (g13 * 7) / 8;
                double g16 = g11 + g14, g17 = g11 + g15;
                var tri = $"M{g16},{vc} L{g11},{g9} L{g11},{g10} Z";
                var bar = $"M{g17},{g9} L{g12},{g9} L{g12},{g10} L{g17},{g10} Z";
                var icon = $"{tri} {bar}";

                return new List<ArrowPathInfo>() {
                            new ArrowPathInfo(){ D= $"{_rect(w, h)} {icon}", Fill= "norm", Stroke= false },
                            new ArrowPathInfo(){ D= icon, Fill= "darken", Stroke= false },
                            new ArrowPathInfo(){ D= icon, Fill= "none", Stroke= true },
                            new ArrowPathInfo(){ D= _rect(w, h), Fill= "none", Stroke= true },
                };
            });

            // actionButtonReturn (VBA 0133): curved return arrow ↩
            // OOXML spec: 4 paths – bg+icon cutout (norm), icon fill (darken), icon outline (stroke), rect outline (stroke)
            // Fill paths use inner arcs curving inward; outline path traces the full shape with reversed arc winding.
            MultiPathPresets.Add("actionButtonReturn", (w, h, a) =>
            {
                var info = AbGuides(w, h);
                var hc = info.HC;
                var vc = info.VC;
                var g9 = info.G9;
                var g10 = info.G10;
                var g11 = info.G11;
                var g12 = info.G12;
                var g13 = info.G13;
                var g14 = (g13 * 7) / 8;
                var g15 = (g13 * 3) / 4;
                var g16 = (g13 * 5) / 8;
                var g17 = (g13 * 3) / 8; // outer arc radius
                var g18 = g13 / 4;
                var g27 = g13 / 8; // inner arc radius
                var g19 = g9 + g15;
                var g20 = g9 + g16;
                var g21 = g9 + g18;
                var g22 = g11 + g14;
                var g23 = g11 + g15;
                var g24 = g11 + g16;
                var g25 = g11 + g17;
                var g26 = g11 + g18;
                // Fill icon path (paths 0 & 1 in OOXML spec — identical geometry)
                // Arc 1: from (g24, g20), wR=g27 hR=g27 stAng=0° swAng=90°
                //   center = (g24-g27, g20), endpoint = (g24-g27, g20+g27) = (g24-g27, g19)
                // Arc 2: from (g25, g19), wR=g27 hR=g27 stAng=90° swAng=90°
                //   center = (g25, g19-g27), endpoint = (g25-g27, g19-g27) = (g26, g20)
                // Arc 3: from (g11, g20), wR=g17 hR=g17 stAng=180° swAng=-90°
                //   center = (g11+g17, g20) = (g25, g20), endpoint = (g25, g20+g17) = (g25, g10)
                // Arc 4: from (hc, g10), wR=g17 hR=g17 stAng=90° swAng=-90°
                //   center = (hc, g10-g17), endpoint = (hc+g17, g10-g17)
                var fillIcon = string.Join(" ", [
                    $"M{g12},{g21}",
                        $"L{g23},{g9}",
                        $"L{info.HC},{g21}",
                        $"L{g24},{g21}",
                        $"L{g24},{g20}",
                        $"A{g27},{g27} 0 0,1 {g24 - g27},{g19}", // arc 1: inner bottom-right corner
                        $"L{g25},{g19}", // across inner bottom
                        $"A{g27},{g27} 0 0,1 {g26},{g20}", // arc 2: inner bottom-left corner
                        $"L{g26},{g21}",
                        $"L{g11},{g21}",
                        $"L{g11},{g20}",
                        $"A{g17},{g17} 0 0,0 {g25},{g10}", // arc 3: outer bottom-left curve
                        $"L{info.HC},{g10}", // across outer bottom
                        $"A{g17},{g17} 0 0,0 {info.HC + g17},{g10 - g17}", // arc 4: outer bottom-right curve
                        $"L{g22},{g21}",
                        "Z",
                    ]);

                // Outline path (path 2 in OOXML spec — traces shape with different arc winding)
                // Starts from right outer edge, traces clockwise: outer right → outer bottom → outer left → inner left → inner bottom → inner right → arrow
                // Arc A: from (g22, g20), wR=g17 hR=g17 stAng=0° swAng=90°
                //   center = (g22-g17, g20) = (g22-g17, g20), endpoint = (g22-g17, g20+g17)
                //   g22-g17 = g11+g14-g17 = g11 + g13*7/8 - g13*3/8 = g11 + g13/2 = g25 + g13/8 = hc? No.
                //   Actually: g22 = g11+g14, g14 = g13*7/8, g17 = g13*3/8
                //   g22 - g17 = g11 + g13*7/8 - g13*3/8 = g11 + g13*4/8 = g11 + g13/2 = hc (since hc = g11 + dx2 = g11 + g13/2)
                //   Hmm wait, dx2 = ss*3/8 and g13 = ss*3/4. So g13/2 = ss*3/8 = dx2. So hc = g11 + dx2 = g11 + g13/2. Yes!
                //   endpoint = (hc, g20+g17) = (hc, g10)? g20+g17 = (g9+g16)+g17 = g9+g13*5/8+g13*3/8 = g9+g13 = g9+ss*3/4
                //   g10 = vc+dx2. g9+g13 = (vc-dx2) + 2*dx2 = vc+dx2 = g10. Yes! endpoint = (hc, g10) ✓ but wait...
                //   Actually stAng=0° means start angle is 0°. center = (g22 - g17*cos(0), g20 - g17*sin(0)) = (g22-g17, g20).
                //   endAng = 0+90 = 90°. endX = center.x + g17*cos(90°) = g22-g17. endY = center.y + g17*sin(90°) = g20+g17.
                //   So endpoint = (g22-g17, g20+g17). var"s verify: g22-g17 = g11+g14-g17 = g11+g13*(7/8-3/8) = g11+g13/2 = g25+g13/8
                //   Hmm, g25 = g11+g17 = g11+g13*3/8. g11+g13/2 = g11+g13*4/8. That"s not g25, it"s g25 + g13/8.
                //   Actually var me just compute: g11+g13/2. g13/2 is not one of the named guides.
                //   OK, the spec says after this arc: lnTo (g25, g10). So endpoint.x must be something, then line to g25.
                //   endpoint.x = g22-g17 = g11+g14-g17 = g11+g13*7/8-g13*3/8 = g11+g13*4/8 = g11+g13/2.
                //   Then lnTo (g25, g10) where g25 = g11+g13*3/8.
                //   endpointY = g20+g17 = g10. So endpoint = (g11+g13/2, g10).
                //   Line from there to (g25, g10) is horizontal. Makes sense.
                // Arc B: from (g25, g10), wR=g17 hR=g17 stAng=90° swAng=90°
                //   center = (g25, g10-g17), endAng=180°
                //   endX = g25+g17*cos(180°) = g25-g17 = g11+g17-g17 = g11
                //   endY = (g10-g17)+g17*sin(180°) = g10-g17 = g20
                //   endpoint = (g11, g20). Then lnTo (g11, g21).
                // Arc C: from (g26, g20), wR=g27 hR=g27 stAng=180° swAng=-90°
                //   center = (g26+g27, g20) = (g26+g27, g20). g26+g27 = g11+g18+g13/8 = g11+g13/4+g13/8 = g11+g13*3/8 = g25
                //   endAng = 180-90 = 90°. endX = g25+g27*cos(90°) = g25. endY = g20+g27*sin(90°) = g20+g27 = g19.
                //   endpoint = (g25, g19). Hmm, but spec says lnTo(hc, g19) after this arc.
                //   Wait: lnTo before spec says "<lnTo><pt x="hc" y="g19"/></lnTo>". So endpoint is (g25, g19), then line to (hc, g19).
                //   Hmm actually spec says: "<lnTo><pt x="hc" y="g19" /></lnTo>".
                //   Wait no: "L(hc, g19)" in the spec.
                // Arc D: from (hc, g19), wR=g27 hR=g27 stAng=90° swAng=-90°
                //   center = (hc, g19-g27), endAng = 0°.
                //   endX = hc+g27*cos(0°) = hc+g27. g19-g27 = g20. endY = g20+g27*sin(0°) = g20.
                //   endpoint = (hc+g27, g20). Hmm, but g24 = g11+g16 = g11+g13*5/8.
                //   hc+g27 = g11+g13/2+g13/8 = g11+g13*5/8 = g24. So endpoint = (g24, g20).
                //   Then lnTo (g24, g21). Then lnTo (hc, g21). Then lnTo (g23, g9). Close.
                var outline = string.Join(" ", [
                    $"M{g12},{g21}",
                        $"L{g22},{g21}",
                        $"L{g22},{g20}",
                        $"A{g17},{g17} 0 0,1 {g11 + g13 / 2},{g10}", // arc A: outer bottom-right (0°→90°)
                        $"L{g25},{g10}", // across outer bottom
                        $"A{g17},{g17} 0 0,1 {g11},{g20}", // arc B: outer bottom-left (90°→180°)
                        $"L{g11},{g21}",
                        $"L{g26},{g21}",
                        $"L{g26},{g20}",
                        $"A{g27},{g27} 0 0,0 {g25},{g19}", // arc C: inner bottom-left (180°→90°, CCW)
                        $"L{hc},{g19}", // across inner bottom
                        $"A{g27},{g27} 0 0,0 {g24},{g20}", // arc D: inner bottom-right (90°→0°, CCW)
                        $"L{g24},{g21}",
                        $"L{hc},{g21}",
                        $"L{g23},{g9}",
                        "Z",
                    ]);

                return new List<ArrowPathInfo>() {
                                new ArrowPathInfo(){ D= $"{_rect(w, h)} {fillIcon}", Fill= "norm", Stroke= false },
                                new ArrowPathInfo(){ D= fillIcon, Fill= "darken", Stroke= false },
                                new ArrowPathInfo(){ D= outline, Fill= "none", Stroke= true },
                                new ArrowPathInfo(){ D= _rect(w, h), Fill= "none", Stroke= true },
                    };
            });

            // actionButtonSound (VBA 0135): speaker icon with 3 sound wave lines
            // OOXML spec: 4 paths – bg+speaker cutout (norm), speaker fill (darken), speaker outline+waves (stroke), rect outline (stroke)
            MultiPathPresets.Add("actionButtonSound", (w, h, a) =>
            {
                var info = AbGuides(w, h);
                var g9 = info.G9;
                var g10 = info.G10;
                var g11 = info.G11;
                var g12 = info.G12;
                var g13 = info.G13;
                var hc = info.HC;
                var vc = info.VC;
                // Guide calculations from OOXML presetShapeDefinitions.xml
                var g14 = g13 / 8;
                var g15 = (g13 * 5) / 16;
                var g16 = (g13 * 5) / 8;
                var g17 = (g13 * 11) / 16;
                var g18 = (g13 * 3) / 4;
                var g19 = (g13 * 7) / 8;
                // Absolute positions
                var g20 = g9 + g14;
                var g21 = g9 + g15;
                var g22 = g9 + g17;
                var g23 = g9 + g19;
                var g24 = g11 + g15;
                var g25 = g11 + g16;
                var g26 = g11 + g18;
                // Speaker shape (pentagon-like)
                var speaker = $"M{g11},{g21} L{g11},{g22} L{g24},{g22} L{g25},{g10} L{g25},{g9} L{g24},{g21} Z";
                // Outline path: speaker outline (different winding) + 3 sound wave lines
                var speakerOutline = $"M{g11},{g21} L{g24},{g21} L{g25},{g9} L{g25},{g10} L{g24},{g22} L{g11},{g22} Z";
                var waveLine1 = $"M{g26},{g21} L{g12},{g20}"; // top-right diagonal
                var waveLine2 = $"M{g26},{vc} L{g12},{vc}"; // middle horizontal
                var waveLine3 = $"M{g26},{g22} L{g12},{g23}"; // bottom-right diagonal
                var outlineWithWaves = $"{speakerOutline} {waveLine1} {waveLine2} {waveLine3}";

                return new List<ArrowPathInfo>() {
                             new ArrowPathInfo(){ D= $"{_rect(w, h)} {speaker}", Fill= "norm", Stroke= false },
                             new ArrowPathInfo(){ D= speaker, Fill= "darken", Stroke= false },
                             new ArrowPathInfo(){ D= outlineWithWaves, Fill= "none", Stroke= true },
                             new ArrowPathInfo(){ D= _rect(w, h), Fill= "none", Stroke= true },
                };
            });

            // actionButtonInformation (VBA 0128): circle with "i" inside
            MultiPathPresets.Add("actionButtonInformation", (w, h, a) =>
            {
                var info = AbGuides(w, h);
                var g9 = info.G9;
                var g10 = info.G10;
                var g11 = info.G11;
                var g12 = info.G12;
                var g13 = info.G13;
                var hc = info.HC;
                var vc = info.VC;
                var dx2 = info.DX2;
                var g14 = g13 / 32;
                var g17v = (g13 * 5) / 16;
                var g18v = (g13 * 3) / 8;
                var g19v = (g13 * 13) / 32;
                var g20v = (g13 * 19) / 32;
                var g22v = (g13 * 11) / 16;
                var g23v = (g13 * 13) / 16;
                var g24v = (g13 * 7) / 8;
                var g38 = (g13 * 3) / 32;
                var y25 = g9 + g14;
                var y28 = g9 + g17v;
                var y29 = g9 + g18v;
                var y30 = g9 + g23v;
                var y31 = g9 + g24v;
                var x32 = g11 + g17v;
                var x34 = g11 + g19v;
                var x35 = g11 + g20v;
                var x37 = g11 + g22v;
                var circle = $"M{hc},{g9} A{dx2},{dx2} 0 1,1 {hc},{g10} A{dx2},{dx2} 0 1,1 {hc},{g9} Z";
                var dot = $"M{hc},{y25} A{g38},{g38} 0 1,1 {hc},{y25 + g38 * 2} A{g38},{g38} 0 1,1 {hc},{y25} Z";
                var iBody = $"M{x32},{y28} L{x37},{y28} L{x37},{y29} L{x35},{y29} L{x35},{y30} L{x37},{y30} L{x37},{y31} L{x32},{y31} L{x32},{y30} L{x34},{y30} L{x34},{y29} L{x32},{y29} Z";
                var iconInner = $"{dot} {iBody}";

                return new List<ArrowPathInfo>() {
                new ArrowPathInfo(){ D= $"{_rect(w, h)} {circle}", Fill= "norm", Stroke= false },
                new ArrowPathInfo(){ D= $"{circle} {iconInner}", Fill= "darken", Stroke= false },
                new ArrowPathInfo(){ D= iconInner, Fill= "lighten", Stroke= false },
                new ArrowPathInfo(){ D= $"{circle} {iconInner}", Fill= "none", Stroke= true },
                new ArrowPathInfo(){ D= _rect(w, h), Fill= "none", Stroke= true },
                };
            });

            // actionButtonHome (VBA 0126): house icon with chimney and door
            // OOXML spec: 5 paths – bg+house cutout (norm), walls+chimney (darkenLess), roof+door (darken),
            // icon outline (stroke), rect outline (stroke)
            MultiPathPresets.Add("actionButtonHome", (w, h, a) =>
            {
                var info = AbGuides(w, h);
                var g9 = info.G9;
                var g10 = info.G10;
                var g11 = info.G11;
                var g12 = info.G12;
                var g13 = info.G13;
                var hc = info.HC;
                var vc = info.VC;
                // Guide calculations from OOXML presetShapeDefinitions.xml
                var g14 = g13 / 16;
                var g15 = g13 / 8;
                var g16 = (g13 * 3) / 16;
                var g17 = (g13 * 5) / 16;
                var g18 = (g13 * 7) / 16;
                var g19 = (g13 * 9) / 16;
                var g20 = (g13 * 11) / 16;
                var g21 = (g13 * 3) / 4;
                var g22 = (g13 * 13) / 16;
                var g23 = (g13 * 7) / 8;
                // Absolute positions
                var g24 = g9 + g14;
                var g25 = g9 + g16;
                var g26 = g9 + g17;
                var g27 = g9 + g21;
                var g28 = g11 + g15;
                var g29 = g11 + g18;
                var g30 = g11 + g19;
                var g31 = g11 + g20;
                var g32 = g11 + g22;
                var g33 = g11 + g23;
                // Path 0: background rect + full house outline cutout (norm, no stroke)
                // House outline: roof triangle → right side → chimney → left side → base
                var houseOutline = $"M{hc},{g9} " +
                    $"L{g11},{vc} L{g28},{vc} L{g28},{g10} L{g33},{g10} L{g33},{vc} L{g12},{vc} " +
                    $"L{g32},{g26} L{g32},{g24} L{g31},{g24} L{g31},{g25} Z";
                // Path 1: walls + chimney (darkenLess, no stroke)
                // Sub-path 1: chimney bar
                var chimney = $"M{g32},{g26} L{g32},{g24} L{g31},{g24} L{g31},{g25} Z";
                // Sub-path 2: house body (walls) with door cutout
                var walls = $"M{g28},{vc} L{g28},{g10} L{g29},{g10} L{g29},{g27} L{g30},{g27} L{g30},{g10} L{g33},{g10} L{g33},{vc} Z";
                // Path 2: roof triangle + door (darken, no stroke)
                var roof = $"M{hc},{g9} L{g11},{vc} L{g12},{vc} Z";
                var door = $"M{g29},{g27} L{g30},{g27} L{g30},{g10} L{g29},{g10} Z";
                // Path 3: icon outline with all detail lines (none fill, stroke)
                var iconOutline = $"M{hc},{g9} " +
                    $"L{g31},{g25} L{g31},{g24} L{g32},{g24} L{g32},{g26} L{g12},{vc} " +
                    $"L{g33},{vc} L{g33},{g10} L{g28},{g10} L{g28},{vc} L{g11},{vc} Z " +
                    // Chimney diagonal line
                    $"M{g31},{g25} L{g32},{g26} " +
                    // Horizontal eave line
                    $"M{g33},{vc} L{g28},{vc} " +
                    // Door outline
                    $"M{g29},{g10} L{g29},{g27} L{g30},{g27} L{g30},{g10}";

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= $"{_rect(w, h)} {houseOutline}", Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= $"{chimney} {walls}", Fill= "darkenLess", Stroke= false },
                        new ArrowPathInfo(){ D= $"{roof} {door}", Fill= "darken", Stroke= false },
                        new ArrowPathInfo(){ D= iconOutline, Fill= "none", Stroke= true },
                        new ArrowPathInfo(){ D= _rect(w, h), Fill= "none", Stroke= true },
                };
            });

            // actionButtonHelp (VBA 0127): question mark "?" inside rectangle
            // OOXML spec: 4 paths – bg+icon cutout (norm), icon fill (darken), icon outline (stroke), rect outline (stroke)
            MultiPathPresets.Add("actionButtonHelp", (w, h, a) =>
            {
                var info = AbGuides(w, h);
                var g9 = info.G9;
                var g10 = info.G10;
                var g11 = info.G11;
                var g12 = info.G12;
                var g13 = info.G13;
                var hc = info.HC;
                var vc = info.VC;
                // Guide calculations from OOXML presetShapeDefinitions.xml
                var g14 = g13 / 7;
                var g15 = (g13 * 3) / 14;
                var g16 = (g13 * 2) / 7;
                var g19 = (g13 * 3) / 7;
                var g20 = (g13 * 4) / 7;
                var g21 = (g13 * 17) / 28;
                var g23 = (g13 * 21) / 28;
                var g24 = (g13 * 11) / 14;
                var g41 = g13 / 14;
                var g42 = (g13 * 3) / 28;
                // Absolute positions
                var g27 = g9 + g16;
                var g29 = g9 + g21;
                var g30 = g9 + g23;
                var g31 = g9 + g24;
                var g33 = g11 + g15;
                var g36 = g11 + g19;
                var g37 = g11 + g20;
                // Helper: OOXML arcTo → SVG arc segment
                // Computes endpoint from center (derived from current point + start angle) and returns SVG A command
                Func<double, double, double, double, double, double, ArcToInfo> arcSeg = (curX, curY, wR, hR, stDeg, swDeg) =>
                {
                    var stRad = (stDeg * Math.PI) / 180;
                    var endRad = ((stDeg + swDeg) * Math.PI) / 180;
                    var cx = curX - wR * Math.Cos(stRad);
                    var cy = curY - hR * Math.Sin(stRad);
                    var endX = cx + wR * Math.Cos(endRad);
                    var endY = cy + hR * Math.Sin(endRad);
                    var largeArc = Math.Abs(swDeg) > 180 ? 1 : 0;
                    var sweep = swDeg > 0 ? 1 : 0;

                    return new ArcToInfo() { EndX = endX, EndY = endY, SVG = $"A{wR},{hR} 0 {largeArc},{sweep} {endX},{endY}" };
                };
                // Build question mark path following OOXML arcTo sequence exactly
                // Start at (g33, g27)
                double cx = g33, cy = g27;
                // Arc 1: wR=g16 hR=g16 stAng=180° swAng=180° (top semicircle, clockwise)
                var a1 = arcSeg(cx, cy, g16, g16, 180, 180);
                cx = a1.EndX;
                cy = a1.EndY;
                // Arc 2: wR=g14 hR=g15 stAng=0° swAng=90° (curve down right)
                var a2 = arcSeg(cx, cy, g14, g15, 0, 90);
                cx = a2.EndX;
                cy = a2.EndY;
                // Arc 3: wR=g41 hR=g42 stAng=270° swAng=-90° (small reverse curve)
                var a3 = arcSeg(cx, cy, g41, g42, 270, -90);
                // After arc 3, lines to stem
                // lnTo (g37, g30), (g36, g30), (g36, g29)
                // then more arcs back up
                // Arc 4: wR=g14 hR=g15 stAng=180° swAng=90° (inner curve going up)
                var a4 = arcSeg(g36, g29, g14, g15, 180, 90);
                // Arc 5: wR=g41 hR=g42 stAng=90° swAng=-90° (small inner reverse curve)
                var a5 = arcSeg(a4.EndX, a4.EndY, g41, g42, 90, -90);
                // Arc 6: wR=g14 hR=g14 stAng=0° swAng=-180° (inner top semicircle, counter-clockwise)
                var a6 = arcSeg(a5.EndX, a5.EndY, g14, g14, 0, -180);
                // Bottom dot circle at (hc, g31) with radius g42
                var dot = $"M{hc},{g31} A{g42},{g42} 0 1,1 {hc},{g31 + g42 * 2} A{g42},{g42} 0 1,1 {hc},{g31} Z";
                // Question mark path (outer shape with arcs + stem + inner cutout arcs)
                var qMark = $"M{g33},{g27} " +
                    $"{a1.SVG} " +
                    $"{a2.SVG} " +
                    $"{a3.SVG} " +
                    $"L{g37},{g30} L{g36},{g30} L{g36},{g29} " +
                    $"{a4.SVG} " +
                    $"{a5.SVG} " +
                    $"{a6.SVG} Z";
                var icon = $"{qMark} {dot}";

                return new List<ArrowPathInfo>() {
                    new ArrowPathInfo(){ D= $"{_rect(w, h)} {icon}", Fill= "norm", Stroke= false }, // Background with icon cutout
                    new ArrowPathInfo(){ D= icon, Fill= "darken", Stroke= false }, // Darkened icon fill
                    new ArrowPathInfo(){ D= icon, Fill= "none", Stroke= true }, // Icon outline
                    new ArrowPathInfo(){ D= _rect(w, h), Fill= "none", Stroke= true }, // Rect outline
                };
            });

            // actionButtonDocument (VBA 0134): document with folded corner
            MultiPathPresets.Add("actionButtonDocument", (w, h, a) =>
            {
                var ss = Math.Min(w, h);
                double hc = w / 2, vc = h / 2;
                var dx2 = (ss * 3) / 8;
                var dx1 = (ss * 9) / 32;
                double g9 = vc - dx2, g10 = vc + dx2;
                double g11 = hc - dx1, g12 = hc + dx1;
                var g13 = (ss * 3) / 16;
                var g14 = g12 - g13;
                var g15 = g9 + g13;
                var doc = $"M{g11},{g9} L{g14},{g9} L{g12},{g15} L{g12},{g10} L{g11},{g10} Z";
                var fold = $"M{g14},{g9} L{g14},{g15} L{g12},{g15} Z";
                var outline = $"{doc} M{g12},{g15} L{g14},{g15} L{g14},{g9}";

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= $"{_rect(w, h)} {doc}", Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= doc, Fill= "darkenLess", Stroke= false },
                        new ArrowPathInfo(){ D= fold, Fill= "darken", Stroke= false },
                        new ArrowPathInfo(){ D= outline, Fill= "none", Stroke= true },
                        new ArrowPathInfo(){ D= _rect(w, h), Fill= "none", Stroke= true },
                };
            });

            // actionButtonMovie (VBA 0136): film strip / camera icon
            MultiPathPresets.Add("actionButtonMovie", (w, h, a) =>
            {
                var info = AbGuides(w, h);
                var g9 = info.G9;
                var g10 = info.G10;
                var g11 = info.G11;
                var g12 = info.G12;
                var g13 = info.G13;
                var hc = info.HC;
                var vc = info.VC;
                // Guide values from OOXML presetShapeDefinitions.xml (fractions of g13 = ss*3/4)
                var g14 = (g13 * 1455) / 21600;
                var g15 = (g13 * 1905) / 21600;
                var g16 = (g13 * 2325) / 21600;
                var g17 = (g13 * 16155) / 21600;
                var g18 = (g13 * 17010) / 21600;
                var g19 = (g13 * 19335) / 21600;
                var g20 = (g13 * 19725) / 21600;
                var g21 = (g13 * 20595) / 21600;
                var g22 = (g13 * 5280) / 21600;
                var g23 = (g13 * 5730) / 21600;
                var g24 = (g13 * 6630) / 21600;
                var g25 = (g13 * 7492) / 21600;
                var g26 = (g13 * 9067) / 21600;
                var g27 = (g13 * 9555) / 21600;
                var g28 = (g13 * 13342) / 21600;
                var g29 = (g13 * 14580) / 21600;
                var g30 = (g13 * 15592) / 21600;
                // Composite guides: x = g11 + gN, y = g9 + gN
                var x31 = g11 + g14;
                var x32 = g11 + g15;
                var x33 = g11 + g16;
                var x34 = g11 + g17;
                var x35 = g11 + g18;
                var x36 = g11 + g19;
                var x37 = g11 + g20;
                var x38 = g11 + g21;
                var y39 = g9 + g22;
                var y40 = g9 + g23;
                var y41 = g9 + g24;
                var y42 = g9 + g25;
                var y43 = g9 + g26;
                var y44 = g9 + g27;
                var y45 = g9 + g28;
                var y46 = g9 + g29;
                var y47 = g9 + g30;
                var icon = string.Join(" ", [
                $"M{g11},{y39}",
                        $"L{g11},{y44}",
                        $"L{x31},{y44}",
                        $"L{x32},{y43}",
                        $"L{x33},{y43}",
                        $"L{x33},{y47}",
                        $"L{x35},{y47}",
                        $"L{x35},{y45}",
                        $"L{x36},{y45}",
                        $"L{x38},{y46}",
                        $"L{g12},{y46}",
                        $"L{g12},{y41}",
                        $"L{x38},{y41}",
                        $"L{x37},{y42}",
                        $"L{x35},{y42}",
                        $"L{x35},{y41}",
                        $"L{x34},{y40}",
                        $"L{x32},{y40}",
                        $"L{x31},{y39}",
                        "Z",
                        ]);

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= $"{_rect(w, h)} {icon}", Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= icon, Fill= "darken", Stroke= false },
                        new ArrowPathInfo(){ D= icon, Fill= "none", Stroke= true },
                        new ArrowPathInfo(){ D= _rect(w, h), Fill= "none", Stroke= true },
            };
            });

            // flowChartOfflineStorage (VBA 0139): inverted triangle with horizontal base line
            MultiPathPresets.Add("flowChartOfflineStorage", (w, h, a) =>
            {
                var tri = $"M0,0 L{w},0 L{w / 2},{h} Z";
                var lineY = (h * 4) / 5;
                var line = $"M{(w * 2) / 5},{lineY} L{(w * 3) / 5},{lineY}";

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= tri, Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= line, Fill= "none", Stroke= true },
                        new ArrowPathInfo(){ D= tri, Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("cube", (w, h, adjustments) =>
            {
                var a = Math.Min(Math.Max(Adjust(adjustments, "adj", 25000), 0), 0.45);
                var depth = Math.Min(w, h) * a;
                var front = string.Join(" ", [
                        $"M0,{depth}",
                    $"L{w - depth},{depth}",
                    $"L{w - depth},{h}",
                    $"L0,{h}",
                    "Z",
                    ]);
                var top = string.Join(" ", [$"M0,{depth}", $"L{depth},0", $"L{w},0", $"L{w - depth},{depth}", "Z"]);
                var right = string.Join(" ", [
                        $"M{w - depth},{depth}",
                    $"L{w},0",
                    $"L{w},{h - depth}",
                    $"L{w - depth},{h}",
                    "Z",
                    ]);

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= front, Fill= "norm", Stroke= true },
                        new ArrowPathInfo(){ D= top, Fill= "lightenLess", Stroke= true },
                        new ArrowPathInfo(){ D= right, Fill= "darkenLess", Stroke= true },
                };
            });

            MultiPathPresets.Add("bevel", (w, h, adjustments) =>
            {
                // OOXML bevel: picture-frame shape with 4 beveled faces + center rect.
                // adj = bevel thickness (default 12500 = 12.5% of min(w,h))
                var a = Math.Min(Math.Max(Adjust(adjustments, "adj", 12500), 0), 0.45);
                var t = Math.Min(w, h) * a;
                var inner = $"M{t},{t} L{w - t},{t} L{w - t},{h - t} L{t},{h - t} Z";
                var top = $"M0,0 L{w},0 L{w - t},{t} L{t},{t} Z";
                var bottom = $"M0,{h} L{t},{h - t} L{w - t},{h - t} L{w},{h} Z";
                var left = $"M0,0 L{t},{t} L{t},{h - t} L0,{h} Z";
                var right = $"M{w},0 L{w},{h} L{w - t},{h - t} L{w - t},{t} Z";

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= inner, Fill= "norm", Stroke= true },
                        new ArrowPathInfo(){ D= top, Fill= "lightenLess", Stroke= true },
                        new ArrowPathInfo(){ D= right, Fill= "darken", Stroke= true },
                        new ArrowPathInfo(){ D= bottom, Fill= "darken", Stroke= true },
                        new ArrowPathInfo(){ D= left, Fill= "lighten", Stroke= true },
                };
            });

            MultiPathPresets.Add("leftRightRibbon", (w, h, adjustments) =>
            {
                // OOXML leftRightRibbon: 3-path shape (body + center fold shadow + stroke outline).
                // adj1=50000 (band height), adj2=50000 (notch width), adj3=16667 (wave amplitude).
                var ss = Math.Min(w, h);
                var wd2 = w / 2;
                var wd32 = w / 32;
                var hc = w / 2;
                var vc = h / 2;
                var a3 = Math.Min(Math.Max((adjustments?["adj3"] ?? 16667) / 100000, 0), 0.33333);
                var maxAdj1 = 1 - a3;
                var a1 = Math.Min(Math.Max((adjustments?["adj1"] ?? 50000) / 100000, 0), maxAdj1);
                var w1 = wd2 - wd32;
                var maxAdj2 = w1 / ss;
                var a2 = Math.Min(Math.Max((adjustments?["adj2"] ?? 50000) / 100000, 0), maxAdj2);
                var x1 = ss * a2;
                var x4 = w - x1;
                var dy1 = (h * a1) / 2;
                var dy2 = (-h * a3) / 2;
                var ly1 = vc + dy2 - dy1;
                var ry4 = vc + dy1 - dy2;
                var ly2 = ly1 + dy1;
                var ry3 = h - ly2;
                var ly4 = ly2 * 2;
                var ry1 = h - ly4;
                var ly3 = ly4 - ly1;
                var ry2 = h - ly3;
                var hR = (a3 * ss) / 4;
                var x2 = hc - wd32;
                var x3 = hc + wd32;
                var y1 = ly1 + hR;
                var y2 = ry2 - hR;
                // Helper: compute OOXML arcTo → SVG arc segment
                Func<double, double, double, double, double, double, ArcToInfo> arcTo = (curX, curY, wR, hRad, stDeg, swDeg) =>
                {
                    var stRad = (stDeg * Math.PI) / 180;
                    var endRad = ((stDeg + swDeg) * Math.PI) / 180;
                    var cx = curX - wR * Math.Cos(stRad);
                    var cy = curY - hRad * Math.Sin(stRad);
                    var endX = cx + wR * Math.Cos(endRad);
                    var endY = cy + hRad * Math.Sin(endRad);
                    var largeArc = Math.Abs(swDeg) > 180 ? 1 : 0;
                    var sweep = swDeg > 0 ? 1 : 0;
                    return new ArcToInfo() { EndX = endX, EndY = endY, SVG = $"A{wR},{hRad} 0 {largeArc},{sweep} {endX},{endY}" };
                };
                // Path 1: Main body (fill, no stroke)
                double cx1 = hc, cy1 = ly1; // after lnTo (hc, ly1)
                var arc1a = arcTo(cx1, cy1, wd32, hR, 270, 180);
                var arc1b = arcTo(arc1a.EndX, arc1a.EndY, wd32, hR, 270, -180);
                double cx1c = hc, cy1c = ry4; // after lnTo (hc, ry4)
                var arc1c = arcTo(cx1c, cy1c, wd32, hR, 90, 90);
                var body = string.Join(" ", [
                    $"M0,{ly2}",
                    $"L{x1},0",
                    $"L{x1},{ly1}",
                    $"L{hc},{ly1}",
                    arc1a.SVG,
                    arc1b.SVG,
                    $"L{x4},{ry2}",
                    $"L{x4},{ry1}",
                    $"L{w},{ry3}",
                    $"L{x4},{h}",
                    $"L{x4},{ry4}",
                    $"L{hc},{ry4}",
                    arc1c.SVG,
                    $"L{x2},{ly3}",
                    $"L{x1},{ly3}",
                    $"L{x1},{ly4}",
                    "Z",
                    ]);
                // Path 2: Center fold shadow (darkenLess, no stroke)
                var arc2a = arcTo(x3, y1, wd32, hR, 0, 90);
                var arc2b = arcTo(arc2a.EndX, arc2a.EndY, wd32, hR, 270, -180);
                var shadow = string.Join(" ", [$"M{x3},{y1}", arc2a.SVG, arc2b.SVG, $"L{x3},{ry2}", "Z"]);
                // Path 3: Stroke outline (no fill) — same as body + interior fold lines
                var outline = string.Join(" ", [body, $"M{x3},{y1} L{x3},{ry2}", $"M{x2},{y2} L{x2},{ly3}"]);

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= body, Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= shadow, Fill= "darkenLess", Stroke= false },
                        new ArrowPathInfo(){ D= outline, Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("ellipseRibbon", (w, h, adjustments) =>
            {
                // OOXML ellipseRibbon: ribbon with parabolic curved bottom edge
                // 3 paths: body (fill=norm), darkenLess shadow folds, outline (fill=none)
                var adj1 = adjustments?["adj1"] ?? 25000;
                var adj2 = adjustments?["adj2"] ?? 50000;
                var adj3 = adjustments?["adj3"] ?? 12500;
                var a1 = Math.Max(0, Math.Min(adj1, 100000));
                var a2 = Math.Max(25000, Math.Min(adj2, 75000));
                var q10 = 100000 - a1;
                var q11 = q10 / 2;
                var q12 = a1 - q11;
                var minAdj3 = Math.Max(0, q12);
                var a3 = Math.Max(minAdj3, Math.Min(adj3, a1));
                var dx2 = (w * a2) / 200000.0;
                var x2 = w / 2 - dx2;
                var x3 = x2 + w / 8;
                var x4 = w - x3;
                var x5 = w - x2;
                var x6 = w - w / 8;
                var dy1 = (h * a3) / 100000.0;
                var f1 = w > 0 ? (4 * dy1) / w : 0;
                // Parabola: p(x) = f1 * x * (1 - x/w)
                Func<double, double> parab = (x) => f1 * (x - (x * x) / w);
                var y1 = parab(x3);
                var cx1 = x3 / 2;
                var cy1 = f1 * cx1; // Bezier control (approximation)
                var cx2 = w - cx1;
                // q1 redefined: total fold height
                var q1 = (h * a1) / 100000.0;
                var dy3 = q1 - dy1;
                var q5 = parab(x2);
                var y3 = q5 + dy3;
                var q6 = dy1 + dy3 - y3;
                var q7 = q6 + dy1;
                var cy3 = q7 + dy3;
                var rh = h - q1;
                var q8 = (dy1 * 14) / 16;
                var y2 = (q8 + rh) / 2;
                var y5 = q5 + rh;
                var y6 = y3 + rh;
                var cx4 = x2 / 2;
                var cy4 = f1 * cx4 + rh;
                var cx5 = w - cx4;
                var cy6 = cy3 + rh;
                var y7 = y1 + dy3;
                var cy7 = q1 + q1 - y7;
                var hc = w / 2;
                var wd8 = w / 8;
                // Path 1: body fill (stroke=false)
                var body = string.Join(" ", [
                    "M0,0",
                    $"Q{cx1},{cy1} {x3},{y1}",
                    $"L{x2},{y3}",
                    $"Q{hc},{cy3} {x5},{y3}",
                    $"L{x4},{y1}",
                    $"Q{cx2},{cy1} {w},0",
                    $"L{x6},{y2}",
                    $"L{w},{rh}",
                    $"Q{cx5},{cy4} {x5},{y5}",
                    $"L{x5},{y6}",
                    $"Q{hc},{cy6} {x2},{y6}",
                    $"L{x2},{y5}",
                    $"Q{cx4},{cy4} 0,{rh}",
                    $"L{wd8},{y2}",
                    "Z",
                    ]);
                // Path 2: darkenLess shadow folds (stroke=false)
                var shadow = string.Join(" ", [
                    $"M{x3},{y7}",
                    $"L{x3},{y1}",
                    $"L{x2},{y3}",
                    $"Q{hc},{cy3} {x5},{y3}",
                    $"L{x4},{y1}",
                    $"L{x4},{y7}",
                    $"Q{hc},{cy7} {x3},{y7}",
                    "Z",
                    ]);
                // Path 3: outline (fill=none)
                var outline = string.Join(" ", [
                    "M0,0",
                    $"Q{cx1},{cy1} {x3},{y1}",
                    $"L{x2},{y3}",
                    $"Q{hc},{cy3} {x5},{y3}",
                    $"L{x4},{y1}",
                    $"Q{cx2},{cy1} {w},0",
                    $"L{x6},{y2}",
                    $"L{w},{rh}",
                    $"Q{cx5},{cy4} {x5},{y5}",
                    $"L{x5},{y6}",
                    $"Q{hc},{cy6} {x2},{y6}",
                    $"L{x2},{y5}",
                    $"Q{cx4},{cy4} 0,{rh}",
                    $"L{wd8},{y2}",
                    "Z",
                    $"M{x2},{y5} L{x2},{y3}",
                    $"M{x5},{y3} L{x5},{y5}",
                    $"M{x3},{y1} L{x3},{y7}",
                    $"M{x4},{y7} L{x4},{y1}",
                    ]);

                return new List<ArrowPathInfo>() {
                            new ArrowPathInfo(){ D= body, Fill= "norm", Stroke= false },
                            new ArrowPathInfo(){ D= shadow, Fill= "darkenLess", Stroke= false },
                            new ArrowPathInfo(){ D= outline, Fill= "none", Stroke= true },
                    };
            });

            MultiPathPresets.Add("ellipseRibbon2", (w, h, adjustments) =>
            {
                // OOXML ellipseRibbon2: inverted ribbon with parabolic curved top edge
                // 3 paths: body (fill=norm), darkenLess shadow folds, outline (fill=none)
                // All y-values computed as b - value (measured from bottom)
                var adj1 = adjustments?["adj1"] ?? 25000;
                var adj2 = adjustments?["adj2"] ?? 50000;
                var adj3 = adjustments?["adj3"] ?? 12500;
                var a1 = Math.Max(0, Math.Min(adj1, 100000));
                var a2 = Math.Max(25000, Math.Min(adj2, 75000));
                var q10 = 100000 - a1;
                var q11 = q10 / 2;
                var q12 = a1 - q11;
                var minAdj3 = Math.Max(0, q12);
                var a3 = Math.Max(minAdj3, Math.Min(adj3, a1));
                var b = h;
                var dx2 = (w * a2) / 200000.0;
                var x2 = w / 2 - dx2;
                var x3 = x2 + w / 8;
                var x4 = w - x3;
                var x5 = w - x2;
                var x6 = w - w / 8;
                var dy1 = (h * a3) / 100000.0;
                var f1 = w > 0 ? (4 * dy1) / w : 0;
                // u1 = parabola at x3
                var u1 = f1 * (x3 - (x3 * x3) / w);
                var y1 = b - u1;
                var cx1 = x3 / 2;
                var cu1 = f1 * cx1;
                var cy1 = b - cu1;
                var cx2 = w - cx1;
                // q1 redefined: total fold height
                var q1 = (h * a1) / 100000.0;
                var dy3 = q1 - dy1;
                var q5 = f1 * (x2 - (x2 * x2) / w);
                var u3 = q5 + dy3;
                var y3 = b - u3;
                var q6 = dy1 + dy3 - u3;
                var q7 = q6 + dy1;
                var cu3 = q7 + dy3;
                var cy3 = b - cu3;
                var rh = b - q1;
                var q8 = (dy1 * 14) / 16;
                var u2 = (q8 + rh) / 2;
                var y2 = b - u2;
                var u5 = q5 + rh;
                var y5 = b - u5;
                var u6 = u3 + rh;
                var y6 = b - u6;
                var cx4 = x2 / 2;
                var cu4 = f1 * cx4 + rh;
                var cy4 = b - cu4;
                var cx5 = w - cx4;
                var cu6 = cu3 + rh;
                var cy6 = b - cu6;
                var u7 = u1 + dy3;
                var y7 = b - u7;
                var cu7 = q1 + q1 - u7;
                var cy7 = b - cu7;
                var hc = w / 2;
                var wd8 = w / 8;
                // Path 1: body fill (stroke=false)
                var body = string.Join(" ", [
                    $"M0,{b}",
                    $"Q{cx1},{cy1} {x3},{y1}",
                    $"L{x2},{y3}",
                    $"Q{hc},{cy3} {x5},{y3}",
                    $"L{x4},{y1}",
                    $"Q{cx2},{cy1} {w},{b}",
                    $"L{x6},{y2}",
                    $"L{w},{q1}",
                    $"Q{cx5},{cy4} {x5},{y5}",
                    $"L{x5},{y6}",
                    $"Q{hc},{cy6} {x2},{y6}",
                    $"L{x2},{y5}",
                    $"Q{cx4},{cy4} 0,{q1}",
                    $"L{wd8},{y2}",
                    "Z",
                    ]);
                // Path 2: darkenLess shadow folds (stroke=false)
                var shadow = string.Join(" ", [
                        $"M{x3},{y7}",
                    $"L{x3},{y1}",
                    $"L{x2},{y3}",
                    $"Q{hc},{cy3} {x5},{y3}",
                    $"L{x4},{y1}",
                    $"L{x4},{y7}",
                    $"Q{hc},{cy7} {x3},{y7}",
                    "Z",
                    ]);
                // Path 3: outline (fill=none)
                var outline = string.Join(" ", [
                    $"M0,{b}",
                    $"L{wd8},{y2}",
                    $"L0,{q1}",
                    $"Q{cx4},{cy4} {x2},{y5}",
                    $"L{x2},{y6}",
                    $"Q{hc},{cy6} {x5},{y6}",
                    $"L{x5},{y5}",
                    $"Q{cx5},{cy4} {w},{q1}",
                    $"L{x6},{y2}",
                    $"L{w},{b}",
                    $"Q{cx2},{cy1} {x4},{y1}",
                    $"L{x5},{y3}",
                    $"Q{hc},{cy3} {x2},{y3}",
                    $"L{x3},{y1}",
                    $"Q{cx1},{cy1} 0,{b}",
                    "Z",
                    $"M{x2},{y3} L{x2},{y5}",
                    $"M{x5},{y5} L{x5},{y3}",
                    $"M{x3},{y7} L{x3},{y1}",
                    $"M{x4},{y1} L{x4},{y7}",
                    ]);

                return new List<ArrowPathInfo>() {
                            new ArrowPathInfo(){ D= body, Fill= "norm", Stroke= false },
                            new ArrowPathInfo(){ D= shadow, Fill= "darkenLess", Stroke= false },
                            new ArrowPathInfo(){ D= outline, Fill= "none", Stroke= true },
                    };
            });

            MultiPathPresets.Add("smileyFace", (w, h, adjustments) =>
            {
                // OOXML smileyFace: 4 paths — face(norm), eyes(darkenLess), smile(none), outline(none+stroke)
                var wd2 = w / 2;
                var hd2 = h / 2;
                var hc = w / 2;
                var vc = h / 2;
                // Adjustment: smile amplitude (default 4653, range -4653..4653)
                var rawAdj = adjustments?["adj"] ?? 4653;
                var a = Math.Max(-4653, Math.Min(rawAdj, 4653));
                // Eye positions (OOXML exact)
                var x2 = (w * 6215) / 21600;
                var x3 = (w * 13135) / 21600;
                var y1 = (h * 7570) / 21600;
                var wR = (w * 1125) / 21600;
                var hR = (h * 1125) / 21600;
                // Smile curve positions (OOXML exact)
                var x1 = (w * 4969) / 21699;
                var x4 = (w * 16640) / 21600;
                var y3 = (h * 16515) / 21600;
                var dy2 = (h * a) / 100000.0;
                var y2 = y3 - dy2;
                var y4 = y3 + dy2;
                var dy3 = (h * a) / 50000.0;
                var y5 = y4 + dy3;
                // Path 1: face ellipse (fill=norm, stroke=false) — two half-arcs for full circle
                var face = $"M{w},{vc} A{wd2},{hd2} 0 1,1 0,{vc} A{wd2},{hd2} 0 1,1 {w},{vc} Z";
                // Path 2: eyes (fill=darkenLess) — two small ellipses at OOXML positions (two half-arcs each)
                var leftEye = $"M{(x2 + wR).ToFixed(2)},{y1.ToFixed(2)} A{wR.ToFixed(2)},{hR.ToFixed(2)} 0 1,1 {(x2 - wR).ToFixed(2)},{y1.ToFixed(2)} A{wR.ToFixed(2)},{hR.ToFixed(2)} 0 1,1 {(x2 + wR).ToFixed(2)},{y1.ToFixed(2)} Z";
                var rightEye = $"M{(x3 + wR).ToFixed(2)},{y1.ToFixed(2)} A{wR.ToFixed(2)},{hR.ToFixed(2)} 0 1,1 {(x3 - wR).ToFixed(2)},{y1.ToFixed(2)} A{wR.ToFixed(2)},{hR.ToFixed(2)} 0 1,1 {(x3 + wR).ToFixed(2)},{y1.ToFixed(2)} Z";
                // Path 3: smile (fill=none) — quadratic Bezier (OOXML quadBezTo)
                var smile = $"M{x1.ToFixed(2)},{y2.ToFixed(2)} Q{hc.ToFixed(2)},{y5.ToFixed(2)} {x4.ToFixed(2)},{y2.ToFixed(2)}";
                // Path 4: face outline (fill=none, stroke=true) — same as path 1
                var outline = $"M{w},{vc} A{wd2},{hd2} 0 1,1 0,{vc} A{wd2},{hd2} 0 1,1 {w},{vc} Z";

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= face, Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= $"{leftEye} {rightEye}", Fill= "darkenLess", Stroke= false },
                        new ArrowPathInfo(){ D= smile, Fill= "none", Stroke= true },
                        new ArrowPathInfo(){ D= outline, Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("foldedCorner", (w, h, adjustments) =>
            {
                var a = Adjust(adjustments, "adj", 16667);
                var fold = Math.Min(w, h) * a * 0.7;
                var body = $"M0,0 L{w},0 L{w},{h - fold} L{w - fold},{h} L0,{h} Z";
                var foldFace = $"M{w - fold},{h} L{w - fold},{h - fold} L{w},{h - fold} Z";
                var crease = $"M{w - fold},{h} L{w - fold},{h - fold}";

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= body, Fill= "norm", Stroke= true },
                        new ArrowPathInfo(){ D= foldFace, Fill= "darkenLess", Stroke= false },
                        new ArrowPathInfo(){ D= crease, Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("can", (w, h, adjustments) =>
            {
                // OOXML: 3 paths — body (norm), top face (lighten), outline (stroke-only)
                var ss = Math.Min(w, h);
                var maxAdj = (50000 * h) / ss;
                var a = Math.Min(Math.Max(adjustments?["adj"] ?? 25000, 0), maxAdj);
                var y1 = (ss * a) / 200000.0;
                var y3 = h - y1;
                var wd2 = w / 2;
                Func<double, double, double, double, double, double, ArcToInfo> arcSeg = (curX, curY, wR, hR, stDeg, swDeg) =>
                {
                    var stRad = (stDeg * Math.PI) / 180;
                    var endRad = ((stDeg + swDeg) * Math.PI) / 180;
                    var cx = curX - wR * Math.Cos(stRad);
                    var cy = curY - hR * Math.Sin(stRad);
                    var endX = cx + wR * Math.Cos(endRad);
                    var endY = cy + hR * Math.Sin(endRad);
                    var largeArc = Math.Abs(swDeg) > 180 ? 1 : 0;
                    var sweep = swDeg > 0 ? 1 : 0;

                    return new ArcToInfo() { EndX = endX, EndY = endY, SVG = $"A{wR},{hR} 0 {largeArc},{sweep} {endX},{endY}" };
                };
                // Path 1: Body (stroke:false, fill:norm)
                var a1 = arcSeg(0, y1, wd2, y1, 180, -180);
                var a2 = arcSeg(w, y3, wd2, y1, 0, 180);
                var body = $"M0,{y1} {a1.SVG} L{w},{y3} {a2.SVG} Z";
                // Path 2: Top face (stroke:false, fill:lighten)
                var a3 = arcSeg(0, y1, wd2, y1, 180, 180);
                var a4 = arcSeg(a3.EndX, a3.EndY, wd2, y1, 0, 180);
                var topFace = $"M0,{y1} {a3.SVG} {a4.SVG} Z";
                // Path 3: Outline (fill:none, stroke:true)
                var a5 = arcSeg(w, y1, wd2, y1, 0, 180);
                var a6 = arcSeg(a5.EndX, a5.EndY, wd2, y1, 180, 180);
                var a7 = arcSeg(w, y3, wd2, y1, 0, 180);
                var outline = $"M{w},{y1} {a5.SVG} {a6.SVG} L{w},{y3} {a7.SVG} L0,{y1}";
                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= body, Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= topFace, Fill= "lighten", Stroke= false },
                        new ArrowPathInfo(){ D= outline, Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("curvedrightarrow", (w, h, adjustments) => BuildCurvedArrowMultiPath("curvedRightArrow", w, h, adjustments));
            MultiPathPresets.Add("curvedleftarrow", (w, h, adjustments) => BuildCurvedArrowMultiPath("curvedLeftArrow", w, h, adjustments));
            MultiPathPresets.Add("curveduparrow", (w, h, adjustments) => buildCurvedVerticalArrowMultiPath("curvedUpArrow", w, h, adjustments));
            MultiPathPresets.Add("curveddownarrow", (w, h, adjustments) => buildCurvedVerticalArrowMultiPath("curvedDownArrow", w, h, adjustments));

            MultiPathPresets.Add("bordercallout1", (w, h, adjustments) =>
            {
                // OOXML: filled+stroked rectangle body + separate leader line (stroke-only).
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 112500)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -38333)) / 100000.0;

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= true },
                        new ArrowPathInfo(){ D= $"M{x1},{y1} L{x2},{y2}", Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("accentcallout1", (w, h, adjustments) =>
            {
                // OOXML: filled rect + accent bar at x1 + 1-segment callout line
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 112500)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -38333)) / 100000.0;

                return new List<ArrowPathInfo>() {
                    new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= false },
                    new ArrowPathInfo(){ D= $"M{x1},0 L{x1},{h}", Fill= "none", Stroke= true },
                    new ArrowPathInfo(){ D= $"M{x1},{y1} L{x2},{y2}", Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("accentcallout2", (w, h, adjustments) =>
            {
                // OOXML: filled rect + accent bar at x1 + 2-segment callout line
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 18750)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -16667)) / 100000.0;
                var y3 = (h * (adjustments?["adj5"] ?? 112500)) / 100000.0;
                var x3 = (w * (adjustments?["adj6"] ?? -46667)) / 100000.0;

                return new List<ArrowPathInfo>() {
                    new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= false },
                    new ArrowPathInfo(){ D= $"M{x1},0 L{x1},{h}", Fill= "none", Stroke= true },
                    new ArrowPathInfo(){ D= $"M{x1},{y1} L{x2},{y2} L{x3},{y3}", Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("accentcallout3", (w, h, adjustments) =>
            {
                // OOXML: filled rect + accent bar at x1 + 3-segment callout line
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 18750)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -16667)) / 100000.0;
                var y3 = (h * (adjustments?["adj5"] ?? 100000)) / 100000.0;
                var x3 = (w * (adjustments?["adj6"] ?? -16667)) / 100000.0;
                var y4 = (h * (adjustments?["adj7"] ?? 112963)) / 100000.0;
                var x4 = (w * (adjustments?["adj8"] ?? -8333)) / 100000.0;

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= $"M{x1},0 L{x1},{h}", Fill= "none", Stroke= true },
                        new ArrowPathInfo(){ D= $"M{x1},{y1} L{x2},{y2} L{x3},{y3} L{x4},{y4}", Fill= "none", Stroke= true },
                };
            });

            // --- callout1/2/3: filled rect (no stroke) + callout line segments ---
            MultiPathPresets.Add("callout1", (w, h, adjustments) =>
            {
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 112500)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -38333)) / 100000.0;

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= $"M{x1},{y1} L{x2},{y2}", Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("callout2", (w, h, adjustments) =>
            {
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 18750)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -16667)) / 100000.0;
                var y3 = (h * (adjustments?["adj5"] ?? 112500)) / 100000.0;
                var x3 = (w * (adjustments?["adj6"] ?? -46667)) / 100000.0;

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= $"M{x1},{y1} L{x2},{y2} L{x3},{y3}", Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("callout3", (w, h, adjustments) =>
            {
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 18750)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -16667)) / 100000.0;
                var y3 = (h * (adjustments?["adj5"] ?? 100000)) / 100000.0;
                var x3 = (w * (adjustments?["adj6"] ?? -16667)) / 100000.0;
                var y4 = (h * (adjustments?["adj7"] ?? 112963)) / 100000.0;
                var x4 = (w * (adjustments?["adj8"] ?? -8333)) / 100000.0;

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= $"M{x1},{y1} L{x2},{y2} L{x3},{y3} L{x4},{y4}", Fill= "none", Stroke= true },
                    };
            });

            // --- borderCallout2/3: filled+stroked rect + callout line segments ---
            MultiPathPresets.Add("bordercallout2", (w, h, adjustments) =>
            {
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 18750)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -16667)) / 100000.0;
                var y3 = (h * (adjustments?["adj5"] ?? 112500)) / 100000.0;
                var x3 = (w * (adjustments?["adj6"] ?? -46667)) / 100000.0;

                return new List<ArrowPathInfo>() {
                    new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= true },
                    new ArrowPathInfo(){ D= $"M{x1},{y1} L{x2},{y2} L{x3},{y3}", Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("bordercallout3", (w, h, adjustments) =>
            {
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 18750)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -16667)) / 100000.0;
                var y3 = (h * (adjustments?["adj5"] ?? 100000)) / 100000.0;
                var x3 = (w * (adjustments?["adj6"] ?? -16667)) / 100000.0;
                var y4 = (h * (adjustments?["adj7"] ?? 112963)) / 100000.0;
                var x4 = (w * (adjustments?["adj8"] ?? -8333)) / 100000.0;

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= true },
                        new ArrowPathInfo(){ D= $"M{x1},{y1} L{x2},{y2} L{x3},{y3} L{x4},{y4}", Fill= "none", Stroke= true },
                };
            });

            // --- accentBorderCallout1/2/3: filled+stroked rect + accent bar + callout line ---
            MultiPathPresets.Add("accentbordercallout1", (w, h, adjustments) =>
            {
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 112500)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -38333)) / 100000.0;

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= true },
                        new ArrowPathInfo(){ D= $"M{x1},0 L{x1},{h}", Fill= "none", Stroke= true },
                        new ArrowPathInfo(){ D= $"M{x1},{y1} L{x2},{y2}", Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("accentbordercallout2", (w, h, adjustments) =>
            {
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 18750)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -16667)) / 100000.0;
                var y3 = (h * (adjustments?["adj5"] ?? 112500)) / 100000.0;
                var x3 = (w * (adjustments?["adj6"] ?? -46667)) / 100000.0;

                return new List<ArrowPathInfo>() {
                    new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= true },
                    new ArrowPathInfo(){ D= $"M{x1},0 L{x1},{h}", Fill= "none", Stroke= true },
                    new ArrowPathInfo(){ D= $"M{x1},{y1} L{x2},{y2} L{x3},{y3}", Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("accentbordercallout3", (w, h, adjustments) =>
            {
                var y1 = (h * (adjustments?["adj1"] ?? 18750)) / 100000.0;
                var x1 = (w * (adjustments?["adj2"] ?? -8333)) / 100000.0;
                var y2 = (h * (adjustments?["adj3"] ?? 18750)) / 100000.0;
                var x2 = (w * (adjustments?["adj4"] ?? -16667)) / 100000.0;
                var y3 = (h * (adjustments?["adj5"] ?? 100000)) / 100000.0;
                var x3 = (w * (adjustments?["adj6"] ?? -16667)) / 100000.0;
                var y4 = (h * (adjustments?["adj7"] ?? 112963)) / 100000.0;
                var x4 = (w * (adjustments?["adj8"] ?? -8333)) / 100000.0;

                return new List<ArrowPathInfo>() {
                       new ArrowPathInfo() { D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= true },
                       new ArrowPathInfo() { D= $"M{x1},0 L{x1},{h}", Fill= "none", Stroke= true },
                       new ArrowPathInfo() { D= $"M{x1},{y1} L{x2},{y2} L{x3},{y3} L{x4},{y4}", Fill= "none", Stroke= true },
                };
            });

            // Chart placeholders: frame + guide lines.
            // PowerPoint uses these as pre-chart placeholders (chartX / chartPlus / chartStar).
            MultiPathPresets.Add("chartx", (w, h, a) =>
            {
                return new List<ArrowPathInfo>() {
                     new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= false },
                     new ArrowPathInfo(){ D= $"M0,0 L{w},{h} M{w},0 L0,{h}", Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("chartplus", (w, h, a) =>
            {
                var cx = w / 2;
                var cy = h / 2;

                return new List<ArrowPathInfo>() {
                    new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= false },
                    new ArrowPathInfo(){ D= $"M{cx},0 L{cx},{h} M0,{cy} L{w},{cy}", Fill= "none", Stroke= true },
                };
            });

            MultiPathPresets.Add("chartstar", (w, h, a) =>
            {
                // OOXML: 3 guide paths — 2 diagonals + 1 vertical (no horizontal center line)
                var cx = w / 2;

                return new List<ArrowPathInfo>() {
                    new ArrowPathInfo(){ D= $"M0,0 L{w},0 L{w},{h} L0,{h} Z", Fill= "norm", Stroke= false },
                    new ArrowPathInfo(){ D= $"M0,0 L{w},{h} M{w},0 L0,{h} M{cx},0 L{cx},{h}", Fill= "none",Stroke= true, },
                };
            });

            // --- ribbon (OOXML spec: 3 paths with arcTo, adj1=16667, adj2=50000) ---
            // Ribbon with tails at top, front panel at bottom. Three paths: body, darkenLess folds, outline.
            MultiPathPresets.Add("ribbon", (w, h, adjustments) =>
            {
                var adj1Raw = adjustments?["adj1"] ?? 16667;
                var adj2Raw = adjustments?["adj2"] ?? 50000;
                var a1 = Math.Min(Math.Max(adj1Raw, 0), 33333);
                var a2 = Math.Min(Math.Max(adj2Raw, 25000), 75000);
                var hc = w / 2;
                var wd8 = w / 8;
                var wd32 = w / 32;
                var x10 = w - wd8;
                var dx2 = (w * a2) / 200000.0;
                var x2 = hc - dx2;
                var x9 = hc + dx2;
                var x3 = x2 + wd32;
                var x8 = x9 - wd32;
                var x5 = x2 + wd8;
                var x6 = x9 - wd8;
                var x4 = x5 - wd32;
                var x7 = x6 + wd32;
                var y1 = (h * a1) / 200000.0;
                var y2 = (h * a1) / 100000.0;
                var y4 = h - y2;
                var y3 = y4 / 2;
                var hR = (h * a1) / 400000;
                var y5 = h - hR;
                var y6 = y2 - hR;
                double cx, cy;
                ArcToInfo arc;
                // Path 1: body fill (stroke=false)
                var p1 = new List<string>();
                cx = 0;
                cy = 0;
                p1.Add($"M{0},{0}");
                p1.Add($"L{x4},{0}");
                cx = x4;
                cy = 0;
                arc = OoArcTo(cx, cy, wd32, hR, 270, 180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x3},{y1}");
                cx = x3;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 270, -180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x8},{y2}");
                cx = x8;
                cy = y2;
                arc = OoArcTo(cx, cy, wd32, hR, 90, -180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x7},{y1}");
                cx = x7;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 90, 180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{w},{0}");
                p1.Add($"L{x10},{y3}");
                p1.Add($"L{w},{y4}");
                p1.Add($"L{x9},{y4}");
                p1.Add($"L{x9},{y5}");
                cx = x9;
                cy = y5;
                arc = OoArcTo(cx, cy, wd32, hR, 0, 90);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x3},{h}");
                cx = x3;
                cy = h;
                arc = OoArcTo(cx, cy, wd32, hR, 90, 90);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x2},{y4}");
                p1.Add($"L{0},{y4}");
                p1.Add($"L{wd8},{y3}");
                p1.Add("Z");
                // Path 2: darkenLess folds (stroke=false)
                var p2 = new List<string>();
                // Left fold
                cx = x5;
                cy = hR;
                p2.Add($"M{cx},{cy}");
                arc = OoArcTo(cx, cy, wd32, hR, 0, 90);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p2.Add($"L{x3},{y1}");
                cx = x3;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 270, -180);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p2.Add($"L{x5},{y2}");
                p2.Add("Z");
                // Right fold
                cx = x6;
                cy = hR;
                p2.Add($"M{cx},{cy}");
                arc = OoArcTo(cx, cy, wd32, hR, 180, -90);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p2.Add($"L{x8},{y1}");
                cx = x8;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 270, 180);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p2.Add($"L{x6},{y2}");
                p2.Add("Z");
                // Path 3: outline (fill=none, includes fold lines)
                var p3 = new List<string>();
                cx = 0;
                cy = 0;
                p3.Add($"M{0},{0}");
                p3.Add($"L{x4},{0}");
                cx = x4;
                cy = 0;
                arc = OoArcTo(cx, cy, wd32, hR, 270, 180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x3},{y1}");
                cx = x3;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 270, -180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x8},{y2}");
                cx = x8;
                cy = y2;
                arc = OoArcTo(cx, cy, wd32, hR, 90, -180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x7},{y1}");
                cx = x7;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 90, 180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{w},{0}");
                p3.Add($"L{x10},{y3}");
                p3.Add($"L{w},{y4}");
                p3.Add($"L{x9},{y4}");
                p3.Add($"L{x9},{y5}");
                cx = x9;
                cy = y5;
                arc = OoArcTo(cx, cy, wd32, hR, 0, 90);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x3},{h}");
                cx = x3;
                cy = h;
                arc = OoArcTo(cx, cy, wd32, hR, 90, 90);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x2},{y4}");
                p3.Add($"L{0},{y4}");
                p3.Add($"L{wd8},{y3}");
                p3.Add("Z");
                // Fold lines
                p3.Add($"M{x5},{hR} L{x5},{y2}");
                p3.Add($"M{x6},{y2} L{x6},{hR}");
                p3.Add($"M{x2},{y4} L{x2},{y6}");
                p3.Add($"M{x9},{y6} L{x9},{y4}");

                return new List<ArrowPathInfo>() {
                        new ArrowPathInfo(){ D= string.Join(" ",p1), Fill= "norm", Stroke= false },
                        new ArrowPathInfo(){ D= string.Join(" ",p2), Fill= "darkenLess", Stroke= false },
                        new ArrowPathInfo(){ D= string.Join(" ",p3), Fill= "none", Stroke= true },
                };
            });

            // --- ribbon2 (OOXML spec: 3 paths, inverted ribbon with tails at bottom) ---
            MultiPathPresets.Add("ribbon2", (w, h, adjustments) =>
            {
                var adj1Raw = adjustments?["adj1"] ?? 16667;
                var adj2Raw = adjustments?["adj2"] ?? 50000;
                var a1 = Math.Min(Math.Max(adj1Raw, 0), 33333);
                var a2 = Math.Min(Math.Max(adj2Raw, 25000), 75000);
                var hc = w / 2;
                var wd8 = w / 8;
                var wd32 = w / 32;
                var x10 = w - wd8;
                var dx2 = (w * a2) / 200000.0;
                var x2 = hc - dx2;
                var x9 = hc + dx2;
                var x3 = x2 + wd32;
                var x8 = x9 - wd32;
                var x5 = x2 + wd8;
                var x6 = x9 - wd8;
                var x4 = x5 - wd32;
                var x7 = x6 + wd32;
                var dy1 = (h * a1) / 200000.0;
                var y1 = h - dy1;
                var dy2 = (h * a1) / 100000.0;
                var y2 = h - dy2;
                var y4 = dy2;
                var y3 = (y4 + h) / 2;
                var hR = (h * a1) / 400000;
                var y6 = h - hR;
                var y7 = y1 - hR;
                double cx, cy;
                ArcToInfo arc;
                // Path 1: body fill (stroke=false)
                var p1 = new List<string>();
                p1.Add($"M{0},{h}");
                p1.Add($"L{x4},{h}");
                cx = x4;
                cy = h;
                arc = OoArcTo(cx, cy, wd32, hR, 90, -180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x3},{y1}");
                cx = x3;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 90, 180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x8},{y2}");
                cx = x8;
                cy = y2;
                arc = OoArcTo(cx, cy, wd32, hR, 270, 180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x7},{y1}");
                cx = x7;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 270, -180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{w},{h}");
                p1.Add($"L{x10},{y3}");
                p1.Add($"L{w},{y4}");
                p1.Add($"L{x9},{y4}");
                p1.Add($"L{x9},{hR}");
                cx = x9;
                cy = hR;
                arc = OoArcTo(cx, cy, wd32, hR, 0, -90);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x3},{0}");
                cx = x3;
                cy = 0;
                arc = OoArcTo(cx, cy, wd32, hR, 270, -90);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x2},{y4}");
                p1.Add($"L{0},{y4}");
                p1.Add($"L{wd8},{y3}");
                p1.Add("Z");
                // Path 2: darkenLess folds (stroke=false)
                var p2 = new List<string>();
                // Left fold
                cx = x5;
                cy = y6;
                p2.Add($"M{cx},{cy}");
                arc = OoArcTo(cx, cy, wd32, hR, 0, -90);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p2.Add($"L{x3},{y1}");
                cx = x3;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 90, 180);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p2.Add($"L{x5},{y2}");
                p2.Add("Z");
                // Right fold
                cx = x6;
                cy = y6;
                p2.Add($"M{cx},{cy}");
                arc = OoArcTo(cx, cy, wd32, hR, 180, 90);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p2.Add($"L{x8},{y1}");
                cx = x8;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 90, -180);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p2.Add($"L{x6},{y2}");
                p2.Add("Z");
                // Path 3: outline (fill=none)
                var p3 = new List<string>();
                p3.Add($"M{0},{h}");
                p3.Add($"L{wd8},{y3}");
                p3.Add($"L{0},{y4}");
                p3.Add($"L{x2},{y4}");
                p3.Add($"L{x2},{hR}");
                cx = x2;
                cy = hR;
                arc = OoArcTo(cx, cy, wd32, hR, 180, 90);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x8},{0}");
                cx = x8;
                cy = 0;
                arc = OoArcTo(cx, cy, wd32, hR, 270, 90);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x9},{y4}");
                p3.Add($"L{w},{y4}");
                p3.Add($"L{x10},{y3}");
                p3.Add($"L{w},{h}");
                p3.Add($"L{x7},{h}");
                cx = x7;
                cy = h;
                arc = OoArcTo(cx, cy, wd32, hR, 90, 180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x8},{y1}");
                cx = x8;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 90, -180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x3},{y2}");
                cx = x3;
                cy = y2;
                arc = OoArcTo(cx, cy, wd32, hR, 270, -180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x4},{y1}");
                cx = x4;
                cy = y1;
                arc = OoArcTo(cx, cy, wd32, hR, 270, 180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add("Z");
                // Fold lines
                p3.Add($"M{x5},{y2} L{x5},{y6}");
                p3.Add($"M{x6},{y6} L{x6},{y2}");
                p3.Add($"M{x2},{y7} L{x2},{y4}");
                p3.Add($"M{x9},{y4} L{x9},{y7}");

                return new List<ArrowPathInfo>() {
                            new ArrowPathInfo(){ D= string.Join(" ",p1), Fill= "norm", Stroke= false },
                            new ArrowPathInfo(){ D= string.Join(" ",p2), Fill= "darkenLess", Stroke= false },
                            new ArrowPathInfo(){ D= string.Join(" ",p3), Fill= "none", Stroke= true },
                    };
            });

            // --- horizontalScroll (OOXML spec: 3 paths with arcTo) ---
            MultiPathPresets.Add("horizontalscroll", (w, h, adjustments) =>
            {
                var adjVal = adjustments?["adj"] ?? 12500;
                var a = Math.Min(Math.Max(adjVal, 0), 25000);
                var ss = Math.Min(w, h);
                var ch = (ss * a) / 100000.0;
                var ch2 = ch / 2;
                var ch4 = ch / 4;
                var y3 = ch + ch2;
                var y4 = ch + ch;
                var y6 = h - ch;
                var y7 = h - ch2;
                var y5 = y6 - ch2;
                var x3 = w - ch;
                var x4 = w - ch2;
                // Path 1: main fill (stroke=false)
                var p1 = new List<string>();
                double cx, cy;
                // moveTo (r, ch2) = (w, ch2)
                cx = w;
                cy = ch2;
                p1.Add($"M{cx},{cy}");
                // arcTo wR=ch2 hR=ch2 stAng=0 swAng=cd4(90°)
                var arc = OoArcTo(cx, cy, ch2, ch2, 0, 90);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                // lnTo (x4, ch2) — but after the arc we should be at (x4, 0)… wait
                // Actually: arcTo from (w, ch2) with stAng=0 swAng=90° → center=(w-ch2, ch2), end=(w-ch2, 0)=x4,0
                // Then lnTo (x4, ch2)... hmm, this goes from top-right curl area
                // var me re-read: lnTo pt x="x4" y="ch2"... that doesn"t match. Wait, the lnTo goes DOWN.
                // After arc: we"re at (x4, 0). lnTo (x4, ch2):
                p1.Add($"L{x4},{ch2}");
                // arcTo wR=ch4 hR=ch4 stAng=0 swAng=cd2(180°)
                arc = OoArcTo(x4, ch2, ch4, ch4, 0, 180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                // lnTo (x3, ch)
                p1.Add($"L{x3},{ch}");
                // lnTo (ch2, ch)
                p1.Add($"L{ch2},{ch}");
                // arcTo wR=ch2 hR=ch2 stAng=3cd4(270°) swAng=-5400000(-90°)
                cx = ch2;
                cy = ch;
                arc = OoArcTo(cx, cy, ch2, ch2, 270, -90);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                // lnTo (0, y7)
                p1.Add($"L{0},{y7}");
                // arcTo wR=ch2 hR=ch2 stAng=cd2(180°) swAng=-10800000(-180°)
                cx = 0;
                cy = y7;
                arc = OoArcTo(cx, cy, ch2, ch2, 180, -180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                // lnTo (ch, y6)
                p1.Add($"L{ch},{y6}");
                // lnTo (x4, y6)
                p1.Add($"L{x4},{y6}");
                // arcTo wR=ch2 hR=ch2 stAng=cd4(90°) swAng=-5400000(-90°)
                cx = x4;
                cy = y6;
                arc = OoArcTo(cx, cy, ch2, ch2, 90, -90);
                p1.Add(arc.SVG);
                p1.Add("Z");
                // Sub-path 2 in Path 1: left bottom curl circle
                cx = ch2;
                cy = y4;
                p1.Add($"M{cx},{cy}");
                // arcTo wR=ch2 hR=ch2 stAng=cd4(90°) swAng=-5400000(-90°)
                arc = OoArcTo(cx, cy, ch2, ch2, 90, -90);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                // arcTo wR=ch4 hR=ch4 stAng=0 swAng=-10800000(-180°)
                arc = OoArcTo(cx, cy, ch4, ch4, 0, -180);
                p1.Add(arc.SVG);
                p1.Add("Z");
                // Path 2: darkenLess fill (stroke=false) — shadow areas
                var p2 = new List<string>();
                // Sub-path 1: same as path1 sub-path2 (left bottom curl)
                cx = ch2;
                cy = y4;
                p2.Add($"M{cx},{cy}");
                arc = OoArcTo(cx, cy, ch2, ch2, 90, -90);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                arc = OoArcTo(cx, cy, ch4, ch4, 0, -180);
                p2.Add(arc.SVG);
                p2.Add("Z");
                // Sub-path 2: right top curl
                cx = x4;
                cy = ch;
                p2.Add($"M{cx},{cy}");
                // arcTo wR=ch2 hR=ch2 stAng=cd4(90°) swAng=-16200000(-270°)
                arc = OoArcTo(cx, cy, ch2, ch2, 90, -270);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                // arcTo wR=ch4 hR=ch4 stAng=cd2(180°) swAng=-10800000(-180°)
                arc = OoArcTo(cx, cy, ch4, ch4, 180, -180);
                p2.Add(arc.SVG);
                p2.Add("Z");
                // Path 3: stroke-only detail lines (fill=none)
                var p3 = new List<string>();
                // Sub-path 1: left side detail
                cx = 0;
                cy = y3;
                p3.Add($"M{cx},{cy}");
                arc = OoArcTo(cx, cy, ch2, ch2, 180, 90);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x3},{ch}");
                p3.Add($"L{x3},{ch2}");
                cx = x3;
                cy = ch2;
                arc = OoArcTo(cx, cy, ch2, ch2, 180, 180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{w},{y5}");
                cx = w;
                cy = y5;
                arc = OoArcTo(cx, cy, ch2, ch2, 0, 90);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{ch},{y6}");
                p3.Add($"L{ch},{y7}");
                cx = ch;
                cy = y7;
                arc = OoArcTo(cx, cy, ch2, ch2, 0, 180);
                p3.Add(arc.SVG);
                p3.Add("Z");
                // Sub-path 2: top-right connector
                p3.Add($"M{x3},{ch}");
                p3.Add($"L{x4},{ch}");
                cx = x4;
                cy = ch;
                arc = OoArcTo(cx, cy, ch2, ch2, 90, -90);
                p3.Add(arc.SVG);
                // Sub-path 3: right curl inner detail
                p3.Add($"M{x4},{ch}");
                p3.Add($"L{x4},{ch2}");
                cx = x4;
                cy = ch2;
                arc = OoArcTo(cx, cy, ch4, ch4, 0, 180);
                p3.Add(arc.SVG);
                // Sub-path 4: left curl inner detail
                p3.Add($"M{ch2},{y4}");
                p3.Add($"L{ch2},{y3}");
                cx = ch2;
                cy = y3;
                arc = OoArcTo(cx, cy, ch4, ch4, 180, 180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                arc = OoArcTo(cx, cy, ch2, ch2, 0, 180);
                p3.Add(arc.SVG);
                // Sub-path 5: vertical divider
                p3.Add($"M{ch},{y3}");
                p3.Add($"L{ch},{y6}");

                return new List<ArrowPathInfo>() {
                    new ArrowPathInfo(){ D= string.Join(" ",p1), Fill= "norm", Stroke= false },
                    new ArrowPathInfo(){ D= string.Join(" ",p2), Fill= "darkenLess", Stroke= false },
                    new ArrowPathInfo(){ D= string.Join(" ",p3), Fill= "none", Stroke= true },
                };
            });

            // --- verticalScroll (OOXML spec: 3 paths with arcTo) ---
            MultiPathPresets.Add("verticalscroll", (w, h, adjustments) =>
            {
                var adjVal = adjustments?["adj"] ?? 12500;
                var a = Math.Min(Math.Max(adjVal, 0), 25000);
                var ss = Math.Min(w, h);
                var ch = (ss * a) / 100000.0;
                var ch2 = ch / 2;
                var ch4 = ch / 4;
                var x3 = ch + ch2;
                var x4 = ch + ch;
                var x6 = w - ch;
                var x7 = w - ch2;
                var _x5 = x6 - ch2;
                var y3 = h - ch;
                var y4 = h - ch2;
                // Path 1: main fill (stroke=false)
                var p1 = new List<string>();
                double cx, cy;
                cx = ch2;
                cy = h;
                p1.Add($"M{cx},{cy}");
                var arc = OoArcTo(cx, cy, ch2, ch2, 90, -90);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{ch2},{y4}");
                cx = ch2;
                cy = y4;
                arc = OoArcTo(cx, cy, ch4, ch4, 90, -180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{ch},{y3}");
                p1.Add($"L{ch},{ch2}");
                cx = ch;
                cy = ch2;
                arc = OoArcTo(cx, cy, ch2, ch2, 180, 90);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x7},{0}");
                cx = x7;
                cy = 0;
                arc = OoArcTo(cx, cy, ch2, ch2, 270, 180);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p1.Add($"L{x6},{ch}");
                p1.Add($"L{x6},{y4}");
                cx = x6;
                cy = y4;
                arc = OoArcTo(cx, cy, ch2, ch2, 0, 90);
                p1.Add(arc.SVG);
                p1.Add("Z");
                // Sub-path 2: top-right curl circle
                cx = x4;
                cy = ch2;
                p1.Add($"M{cx},{cy}");
                arc = OoArcTo(cx, cy, ch2, ch2, 0, 90);
                p1.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                arc = OoArcTo(cx, cy, ch4, ch4, 90, 180);
                p1.Add(arc.SVG);
                p1.Add("Z");
                // Path 2: darkenLess fill (stroke=false)
                var p2 = new List<string>();
                cx = x4;
                cy = ch2;
                p2.Add($"M{cx},{cy}");
                arc = OoArcTo(cx, cy, ch2, ch2, 0, 90);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                arc = OoArcTo(cx, cy, ch4, ch4, 90, 180);
                p2.Add(arc.SVG);
                p2.Add("Z");
                cx = ch;
                cy = y4;
                p2.Add($"M{cx},{cy}");
                arc = OoArcTo(cx, cy, ch2, ch2, 0, 270);
                p2.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                arc = OoArcTo(cx, cy, ch4, ch4, 270, 180);
                p2.Add(arc.SVG);
                p2.Add("Z");
                // Path 3: stroke-only detail lines (fill=none)
                var p3 = new List<string>();
                cx = ch;
                cy = y3;
                p3.Add($"M{cx},{cy}");
                p3.Add($"L{ch},{ch2}");
                cx = ch;
                cy = ch2;
                arc = OoArcTo(cx, cy, ch2, ch2, 180, 90);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x7},{0}");
                cx = x7;
                cy = 0;
                arc = OoArcTo(cx, cy, ch2, ch2, 270, 180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x6},{ch}");
                p3.Add($"L{x6},{y4}");
                cx = x6;
                cy = y4;
                arc = OoArcTo(cx, cy, ch2, ch2, 0, 90);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{ch2},{h}");
                cx = ch2;
                cy = h;
                arc = OoArcTo(cx, cy, ch2, ch2, 90, 180);
                p3.Add(arc.SVG);
                p3.Add("Z");
                // top curl
                p3.Add($"M{x3},{0}");
                cx = x3;
                cy = 0;
                arc = OoArcTo(cx, cy, ch2, ch2, 270, 180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                arc = OoArcTo(cx, cy, ch4, ch4, 90, 180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{x4},{ch2}");
                // horizontal divider
                p3.Add($"M{x6},{ch}");
                p3.Add($"L{x3},{ch}");
                // bottom-left curl detail
                p3.Add($"M{ch2},{y3}");
                cx = ch2;
                cy = y3;
                arc = OoArcTo(cx, cy, ch4, ch4, 270, 180);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{ch},{y4}");
                // bottom curl
                p3.Add($"M{ch2},{h}");
                cx = ch2;
                cy = h;
                arc = OoArcTo(cx, cy, ch2, ch2, 90, -90);
                p3.Add(arc.SVG);
                cx = arc.X;
                cy = arc.Y;
                p3.Add($"L{ch},{y3}");

                return new List<ArrowPathInfo>() {
                    new ArrowPathInfo(){ D= string.Join(" ",p1), Fill= "norm", Stroke= false },
                    new ArrowPathInfo(){ D= string.Join(" ",p2), Fill= "darkenLess", Stroke= false },
                    new ArrowPathInfo(){ D= string.Join(" ",p3), Fill= "none", Stroke= true },
                };
            });
        }

        /// <summary>
        /// Get overlay paths for a preset shape (3D top faces, etc.).
        /// Returns empty array if the shape has no overlays.
        /// </summary>
        /// <param name="shapeType"></param>
        /// <param name="w"></param>
        /// <param name="h"></param>
        /// <param name="adjustments"></param>
        /// <returns></returns>
        public static List<PresetOverlayInfo> GetPresetOverlays(string shapeType, double w, double h, Dictionary<string, int> adjustments)
        {
            var key = shapeType.ToLower();
            var gen = PresetOverlays[key] ?? PresetOverlays[shapeType];

            return gen != null ? gen(w, h, adjustments) : new List<PresetOverlayInfo>();
        }

        /// <summary>
        /// ==== Action Button multi-path presets (OOXML spec-accurate) ====
        /// Common helper: OOXML action button guide values
        /// </summary>
        /// <param name="w"></param>
        /// <param name="h"></param>
        /// <returns></returns>
        public static AbGuidesInfo AbGuides(double w, double h)
        {
            var ss = Math.Min(w, h);
            double hc = w / 2, vc = h / 2;
            var dx2 = (ss * 3) / 8; // icon half-extent

            return new AbGuidesInfo()
            {
                SS = ss,
                HC = hc,
                VC = vc,
                DX2 = dx2,
                G9 = vc - dx2,
                G10 = vc + dx2,
                G11 = hc - dx2,
                G12 = hc + dx2,
                G13 = (ss * 3) / 4,
            };
        }

        /// <summary>
        /// Get multi-path preset sub-paths for a shape type.
        /// Returns null if the shape is not a multi-path preset (use GetPresetShapePath instead).
        /// </summary>
        /// <param name="shapeType"></param>
        /// <param name="w"></param>
        /// <param name="h"></param>
        /// <param name="adjustments"></param>
        /// <returns></returns>
        public static List<ArrowPathInfo> GetMultiPathPreset(string shapeType, double w, double h, Dictionary<string, int> adjustments)
        {
            string key = MultiPathPresets.FirstOrDefault(item => item.Key.ToLower() == shapeType.ToLower()).Key;

            var gen = key != null ? MultiPathPresets[shapeType] : null;

            return gen != null ? gen(w, h, adjustments) : null;
        }

        public static string GetPresetShapePath(string shapeType, double w, double h, Dictionary<string, int> adjustments)
        {
            // <a:prstGeom prst="textNoShape"> means text-only shape without geometry.
            if (shapeType == "textNoShape" || shapeType.ToLower() == "textnoshape")
                return "";

            // OOXML preset names are often camelCase; normalize to lowercase for lookup
            string key = PresetShapes.FirstOrDefault(item => item.Key.ToLower() == shapeType.ToLower()).Key;
            var generator = key != null ? PresetShapes[key] : null;

            if (generator != null)
            {
                return generator(w, h, adjustments);
            }

            return $"M0,0 L{w},0 L{w},{h} L0,{h} Z";
        }

        /// <summary>
        /// Helper: compute OOXML arcTo endpoint and SVG arc command from current position.
        /// OOXML arcTo: center = curPos - radius*dir(stAng), endpoint = center + radius*dir(stAng+swAng)
        ///  Returns { svgArc, endX, endY }
        /// </summary>
        /// <param name="curX"></param>
        /// <param name="curY"></param>
        /// <param name="wR"></param>
        /// <param name="hR"></param>
        /// <param name="stAngDeg"></param>
        /// <param name="swAngDeg"></param>
        /// <returns></returns>
        public static ArcToInfo OoArcTo(double curX, double curY, double wR, double hR, int stAngDeg, int swAngDeg)
        {
            var stRad = (stAngDeg * Math.PI) / 180;
            var cx = curX - wR * Math.Cos(stRad);
            var cy = curY - hR * Math.Sin(stRad);
            var endRad = ((stAngDeg + swAngDeg) * Math.PI) / 180;
            var ex = cx + wR * Math.Cos(endRad);
            var ey = cy + hR * Math.Sin(endRad);
            var absSweep = Math.Abs(swAngDeg);
            var largeArc = absSweep > 180 ? 1 : 0;
            var sweepFlag = swAngDeg >= 0 ? 1 : 0;

            return new ArcToInfo() { SVG = $"A{wR},{hR} 0 {largeArc},{sweepFlag} {ex},{ey}", X = ex, Y = ey };
        }

        public static string GearShape(double w, double h, int teeth, double adj1Raw, double adj2Raw)
        {
            // Gear shape: teeth protrude from inner ellipse by th, narrowed by lFD at tips.
            // Uses per-tooth edge-perpendicular computation for B/C tip direction.
            var cx = w / 2;
            var cy = h / 2;
            var ss = Math.Min(w, h);
            var maxAdj2 = teeth == 6 ? 5358 : 2679;
            var a1v = Math.Min(Math.Max(adj1Raw, 0), 20000);
            var a2v = Math.Min(Math.Max(adj2Raw, 0), maxAdj2);
            var th = (ss * a1v) / 100000.0; // tooth height
            var lFD = (ss * a2v) / 100000.0; // tooth flat distance offset
            var rw = w / 2 - th; // inner ellipse width radius
            var rh = h / 2 - th; // inner ellipse height radius
            if (rw <= 0 || rh <= 0)
                return $"M0,0 L{w},0 L{w},{h} L0,{h} Z";
            // OOXML: ha = at2(maxr, l3) where maxr=min(rw,rh), l3=th/2+lFD/2
            var l3 = th / 2 + lFD / 2;
            var maxr = Math.Min(rw, rh);
            var ha = Math.Atan2(l3, maxr); // half-angle of each tooth on the inner ellipse
            int[] centerDegs = teeth == 6 ? [330, 30, 90, 150, 210, 270] : [310, 350, 30, 70, 110, 150, 190, 230, 270];
            var parts = new List<string>();

            for (var i = 0; i < centerDegs.Length; i++)
            {
                var baseAngle = (centerDegs[i] * Math.PI) / 180;
                var aStart = baseAngle - ha; // tooth base start angle (A point)
                var aEnd = baseAngle + ha; // tooth base end angle (D point)
                                           // A and D: inner ellipse points at tooth base edges
                var ax = cx + rw * Math.Cos(aStart);
                var ay = cy + rh * Math.Sin(aStart);
                var dx = cx + rw * Math.Cos(aEnd);
                var dy = cy + rh * Math.Sin(aEnd);
                // Per-tooth edge-perpendicular tip computation:
                // Edge direction A→D
                var edgeX = dx - ax;
                var edgeY = dy - ay;
                var edgeLen = Math.Sqrt(edgeX * edgeX + edgeY * edgeY);
                // Unit normal perpendicular to edge, pointing outward
                // For clockwise winding (our standard), outward normal is (-edgeY, edgeX) / len
                // Verify with radial dot product and flip if needed
                var radX = Math.Cos(baseAngle);
                var radY = Math.Sin(baseAngle);
                var nx = radX;
                var ny = radY;

                if (edgeLen > 0)
                {
                    nx = -edgeY / edgeLen;
                    ny = edgeX / edgeLen;
                }
                if (nx * radX + ny * radY < 0)
                {
                    nx = -nx;
                    ny = -ny;
                }
                // Narrowing: slide A and D inward along edge by lFD
                var ex = edgeLen > 0 ? edgeX / edgeLen : 0;
                var ey = edgeLen > 0 ? edgeY / edgeLen : 0;
                var axN = ax + ex * lFD; // A narrowed (moved toward D)
                var ayN = ay + ey * lFD;
                var dxN = dx - ex * lFD; // D narrowed (moved toward A)
                var dyN = dy - ey * lFD;
                // B and C: tip points = narrowed base + th * outward normal
                var bx = axN + nx * th;
                var by = ayN + ny * th;
                var _cx = dxN + nx * th;
                var _cy = dyN + ny * th;

                if (i == 0)
                {
                    // Start at the valley before first tooth
                    var prevEnd = (centerDegs[centerDegs.Length - 1] * Math.PI) / 180 + ha;
                    var prevIx = cx + rw * Math.Cos(prevEnd);
                    var prevIy = cy + rh * Math.Sin(prevEnd);

                    parts.Add($"M{prevIx},{prevIy}");
                    parts.Add($"A{rw},{rh} 0 0,1 {ax},{ay}");
                }

                // Tooth: A→B→C→D
                parts.Add($"L{bx},{by}");
                parts.Add($"L{_cx},{_cy}");
                parts.Add($"L{dx},{dy}");

                // Arc along inner ring to next tooth
                if (i < centerDegs.Length - 1)
                {
                    var nextStart = (centerDegs[i + 1] * Math.PI) / 180 - ha;
                    var nx2 = cx + rw * Math.Cos(nextStart);
                    var ny2 = cy + rh * Math.Sin(nextStart);

                    parts.Add($"A{rw},{rh} 0 0,1 {nx2},{ny2}");
                }
            }

            parts.Add("Z");
            return string.Join(" ", parts);
        }
    }
}
