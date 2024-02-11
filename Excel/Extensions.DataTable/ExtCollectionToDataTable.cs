using System.Data;

public static class ExtCollectionToDataTable
{
    public static DataTable ToDataTable<T>(this IEnumerable<T> collection)
    {
        DataTable dt = new DataTable();

        var properties = typeof(T).GetProperties();
        foreach (var property in properties)
        {
            dt.Columns.Add(property.Name, property.PropertyType);
        }

        foreach (T item in collection)
        {
            var newRow = dt.NewRow();
            foreach (var property in properties)
            {
                var value = property.GetValue(item);
                newRow[property.Name] = value;
            }
            dt.Rows.Add(newRow);
        }
        return dt;
    }
}
