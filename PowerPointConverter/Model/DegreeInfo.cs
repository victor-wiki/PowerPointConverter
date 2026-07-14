namespace PowerPointConverter.Model
{
    public class DegreeInfo
    {
        public static readonly DegreeInfo Degree270 = FromDouble(270.0);

        public static readonly DegreeInfo Degree90 = FromDouble(90.0);

        public static readonly DegreeInfo Degree180 = FromDouble(180.0);

        private int integerValue;

        private const int MaxDegreeValue = 21600000;

        private const double Precision = 60000.0;

        public int IntValue
        {
            get
            {
                return integerValue;
            }

            private set
            {
                int num = value % 21600000;

                if (num < 0)
                {
                    num += 21600000;
                }

                integerValue = num;
            }
        }

        public double DoubleValue => IntValue / 60000.0;

        public DegreeInfo(int value)
        {
            IntValue = value;
        }

        public double ToRadiansValue()
        {
            return DoubleValue / 180.0 * Math.PI;
        }

        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }

            DegreeInfo degree = (DegreeInfo)obj;

            return IntValue == degree.IntValue;
        }

        public override int GetHashCode()
        {
            return IntValue.GetHashCode();
        }

        public static DegreeInfo operator +(DegreeInfo a)
        {
            return a;
        }

        public static DegreeInfo operator -(DegreeInfo a)
        {
            return new DegreeInfo(-a.IntValue);
        }

        public static DegreeInfo operator +(DegreeInfo a, DegreeInfo b)
        {
            return new DegreeInfo(a.IntValue + b.IntValue);
        }

        public static DegreeInfo operator -(DegreeInfo a, DegreeInfo b)
        {
            return a + -b;
        }

        public static DegreeInfo operator *(DegreeInfo a, double b)
        {
            return new DegreeInfo((int)(a.IntValue * b));
        }

        public static DegreeInfo operator *(double a, DegreeInfo b)
        {
            return new DegreeInfo((int)(a * b.IntValue));
        }

        public static DegreeInfo operator /(DegreeInfo a, double b)
        {
            if (b == 0.0)
            {
                throw new DivideByZeroException();
            }

            return new DegreeInfo((int)(a.IntValue / b));
        }

        public static bool operator >(DegreeInfo a, DegreeInfo b)
        {
            return a.IntValue > b.IntValue;
        }

        public static bool operator <(DegreeInfo a, DegreeInfo b)
        {
            return b > a;
        }

        public static bool operator ==(DegreeInfo a, DegreeInfo b)
        {
            return a.IntValue == b.IntValue;
        }

        public static bool operator !=(DegreeInfo a, DegreeInfo b)
        {
            return !(a == b);
        }

        public static DegreeInfo FromDouble(double value)
        {
            return new DegreeInfo((int)(value * 60000.0));
        }
    }
}
