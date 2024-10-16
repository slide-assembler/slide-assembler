public class Series
{
    public string name { get; set; }
    public double[] values { get; set; }
    public Series(string name, double[] values)
    {
        this.name = name;
        this.values = values;
    }

}

