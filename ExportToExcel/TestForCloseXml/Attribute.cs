namespace TestForCloseXml
{
    public class MyAttribute
    {
        [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
        public class ExcelColumnAttribute : Attribute
        {
            private readonly string name;

            public ExcelColumnAttribute(string name)
            {
                this.name = name;
            }
            public string Name { get { return name; } }
        }

        public interface IExcelExporterModel { }
    }
}
