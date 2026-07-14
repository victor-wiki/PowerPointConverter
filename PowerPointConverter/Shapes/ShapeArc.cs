namespace PowerPointConverter.Shapes
{
    public class ShapeArc
    {
        /// <summary>
        /// Convert OOXML arc specification to SVG path arc command.
        /// Based on PPTXjs shapeArc() implementation.
        /// </summary>
        /// <param name="cx">Center X coordinate</param>
        /// <param name="cy">Center Y coordinate</param>
        /// <param name="rx">Horizontal radius</param>
        /// <param name="ry">Vertical radius</param>
        /// <param name="startAngle">Start angle in degrees</param>
        /// <param name="endAngle">End angle in degrees</param>
        /// <param name="isClose">Whether to close the path with Z</param>
        /// <returns>SVG path string for the arc</returns>
        public static string Arc(double cx, double cy, double rx, double ry, double startAngle, double endAngle, bool isClose)
        {
            var startRad = (startAngle * Math.PI) / 180;
            var endRad = (endAngle * Math.PI) / 180;
            var x1 = cx + rx * Math.Cos(startRad);
            var y1 = cy + ry * Math.Sin(startRad);
            var x2 = cx + rx * Math.Cos(endRad);
            var y2 = cy + ry * Math.Sin(endRad);

            // OOXML convention: always sweep clockwise from startAngle to endAngle.
            // Compute the clockwise sweep in degrees, handling angle wrapping.
            var sweepDeg = (((endAngle - startAngle) % 360) + 360) % 360;

            if (sweepDeg == 0 && startAngle != endAngle)
                sweepDeg = 360;

            var largeArc = sweepDeg > 180 ? 1 : 0;
            var sweep = 1; // always clockwise

            var d = $"M{x1},{y1} A{rx},{ry} 0 {largeArc},{sweep} {x2},{y2}";

            if (isClose)
            {
                d += " Z";
            }

            return d;
        }
    }
}
