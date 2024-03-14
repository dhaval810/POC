using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

internal class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Load the Word template
            Document doc = new Document(@"C:\Project\POC\POC\test.docx");

            // Sample data (you can replace this with actual data retrieved from your API)
            // string receiptNumber = "123456";

            List<Customer> customers = new List<Customer>
                {
                    new Customer("Thomas Hardy", "120 Hanover Sq., London"),
                    new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"),
                    new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"),
                    new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"),
                    new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"),
                    new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"),
                    new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino")
                };

            // Create a custom mail merge data source
            CustomerMailMergeDataSource dataSource = new CustomerMailMergeDataSource(customers);

            // Perform mail merge
            doc.MailMerge.UseNonMergeFields = true;
            // doc.MailMerge.Execute(dataSource);
            //string[] fieldNames = {
            //    "ReceiptNumber"
            //};
            //string[] fieldValues = {
            //    "12345677"
            //};
            //doc.MailMerge.UseNonMergeFields = true;

            //doc.MailMerge.Execute(fieldNames, fieldValues);

            doc.MailMerge.ExecuteWithRegions(dataSource);

            // Save the merged document as PDF
            doc.Save(@"C:\Project\POC\POC\Receipts.pdf", SaveFormat.Pdf);

            Console.WriteLine("Word document merged and saved successfully as PDF.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

public class Customer
{
    public Customer(string aFullName, string anAddress)
    {
        FullName = aFullName;
        Address = anAddress;
    }

    public string FullName { get; set; }
    public string Address { get; set; }
}
/// <summary>
/// A custom mail merge data source that you implement to allow Aspose.Words 
/// to mail merge data from your Customer objects into Microsoft Word documents.
/// </summary>
public class CustomerMailMergeDataSource : IMailMergeDataSource
{
    public CustomerMailMergeDataSource(List<Customer> customers)
    {
        mCustomers = customers;

        // When we initialize the data source, its position must be before the first record.
        mRecordIndex = -1;
    }

    /// <summary>
    /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
    /// </summary>
    public string TableName
    {
        get { return "customers"; }
    }

    /// <summary>
    /// Aspose.Words calls this method to get a value for every data field.
    /// </summary>
    public bool GetValue(string fieldName, out object fieldValue)
    {
        switch (fieldName)
        {
            case "FullName":
                fieldValue = mCustomers[mRecordIndex].FullName;
                return true;
            case "Address":
                fieldValue = mCustomers[mRecordIndex].Address;
                return true;
            default:
                // Return "false" to the Aspose.Words mail merge engine to signify
                // that we could not find a field with this name.
                fieldValue = null;
                return false;
        }
    }

    /// <summary>
    /// A standard implementation for moving to a next record in a collection.
    /// </summary>
    public bool MoveNext()
    {
        if (!IsEof)
            mRecordIndex++;

        return !IsEof;
    }

    public IMailMergeDataSource GetChildDataSource(string tableName)
    {
        return null;
    }

    private bool IsEof
    {
        get { return (mRecordIndex >= mCustomers.Count); }
    }

    private readonly List<Customer> mCustomers;
    private int mRecordIndex;
}