namespace DocSharp.Wmf2Svg.Gdi;

public sealed class Point
{
    public int X { get; set; }

    public int Y { get; set; }

    public Point(int x, int y)
    {
        X = x;
        Y = y;
    }

    public override int GetHashCode()
    {
        const int prime = 31;
        var result = 1;
        result = prime * result + X;
        result = prime * result + Y;
        return result;
    }

    public override bool Equals(object? obj)
    {
        if (this == obj)
        {
            return true;
        }

        if (obj == null)
        {
            return false;
        }

        if (GetType() != obj.GetType())
        {
            return false;
        }

        var other = (Point)obj;
        if (X != other.X)
        {
            return false;
        }

        if (Y != other.Y)
        {
            return false;
        }

        return true;
    }

    public override string ToString()
    {
        return $"Point [X={X}, Y={Y}]";
    }
}