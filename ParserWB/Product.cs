namespace ParserWB
{
    class ProductWB
    {
        public Metadata metadata { get; set; }
        public int state { get; set; }
        public int version { get; set; }
        public Data data { get; set; }
    }

    public class Metadata
    {
        public string name { get; set; }
        public string catalog_type { get; set; }
        public string catalog_value { get; set; }
    }

    public class Data
    {
        public Product[] products { get; set; }
    }

    public class Product
    {
        public int time1 { get; set; }
        public int time2 { get; set; }
        public int id { get; set; }
        public int root { get; set; }
        public int kindId { get; set; }
        public int subjectId { get; set; }
        public int subjectParentId { get; set; }
        public string name { get; set; }
        public string brand { get; set; }
        public int brandId { get; set; }
        public int siteBrandId { get; set; }
        public int sale { get; set; }
        public int priceU { get; set; }
        public int salePriceU { get; set; }
        public int pics { get; set; }
        public int rating { get; set; }
        public int feedbacks { get; set; }
        public Color[] colors { get; set; }
        public Size[] sizes { get; set; }
        public bool diffPrice { get; set; }
        public int panelPromoId { get; set; }
        public string promoTextCat { get; set; }
    }

    public class Color
    {
        public string name { get; set; }
        public int id { get; set; }
    }

    public class Size
    {
        public string name { get; set; }
        public string origName { get; set; }
        public int rank { get; set; }
        public int optionId { get; set; }
    }

}

