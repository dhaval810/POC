using Aspose.Words;
using Aspose.Words.MailMerging;
using Org.BouncyCastle.Utilities;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;


//using System.Net.Mail;
//using System.Net;
using System.Security.Cryptography.X509Certificates;

class Program
{
    static void Main(string[] args)
    {
        IWebProxy proxy = WebRequest.GetSystemWebProxy();

        proxy.Credentials = CredentialCache.DefaultCredentials;

        HttpClient.DefaultProxy = proxy;


        try
        {
            // Load the Word template
            Document doc = new Document(@"C:\Project\POC\POC\Receipt - Copy.docx");

            // Sample data
            PaymentDetails paymentDetails = new PaymentDetails
            {
                InstanceId = 37253,
                Date = "05/03/2024 07:01 PM",
                Total = 64798.17,
                GstTotal = 5,
                MerchantFee = 1295.96,
                TotalPayment = 66090.13,
                ImageUrl = "https://www.sketchgroup.com.au/wp-content/uploads/2022/07/logo-project-01.png",
                Accounts = new List<Account>
            {
                new Account
                {
                   // ModuleReference = "RT",
                    AccountLabel = "Rates - 69889",
                    Details = new List<string> { "6 Lansell Road West PELICAN NSW 2281" },
                    AccountTotal = 64041.12,
                   // AccountGstTotal = 0,
                    ReceiptTypes = new List<ReceiptType>
                    {
                        new ReceiptType
                        {
                            ReceiptTypeLabel = "Rates receipt type",
                            ReceiptTypeTotal = 64041.12,
                            ReceiptTypeGstTotal = 0,
                            PaymentLines = new List<PaymentLinesData>
                            {
                                new PaymentLinesData
                                {
                                    LineLabel = "Line 1 test",
                                    Amount = 100,
                                    //GstAmount = 10
                                },
                                new PaymentLinesData
                                {
                                    LineLabel = "Line 2 test test",
                                    Amount = 200,
                                   // GstAmount = 20
                                },
                                new PaymentLinesData
                                {
                                    LineLabel = "Line 3 test test test",
                                    Amount = 200,
                                   // GstAmount = 20
                                }
                            }
                        }
                    }
                },
                new Account
                {
                   // ModuleReference = "IN",
                    AccountLabel = "Infringements - 296",
                    Details = new List<string>(),
                    AccountTotal = 85,
                    AccountGstTotal = 0,
                    ReceiptTypes = new List<ReceiptType>
                    {
                        new ReceiptType
                        {
                            ReceiptTypeLabel = "Infringement Notice",
                            ReceiptTypeTotal = 85,
                            ReceiptTypeGstTotal = 0,
                            PaymentLines = new List<PaymentLinesData>()
                        }
                    }
                },
                new Account
                {
                   // ModuleReference = "WB",
                    AccountLabel = "Water Billing - 701529",
                    Details = new List<string> { "test test 6 Lansell Road West PELICAN NSW 2281" },
                    AccountTotal = 672.05,
                    AccountGstTotal = 0,
                    ReceiptTypes = new List<ReceiptType>
                    {
                        new ReceiptType
                        {
                            ReceiptTypeLabel = "Water Billing",
                            ReceiptTypeTotal = 672.05,
                           // ReceiptTypeGstTotal = 0,
                           PaymentLines = new List<PaymentLinesData>
                            {
                                new PaymentLinesData
                                {
                                    LineLabel = "Line 1",
                                    Amount = 100,
                                    //GstAmount = 10
                                },
                                new PaymentLinesData
                                {
                                    LineLabel = "Line 2",
                                    Amount = 200,
                                   // GstAmount = 20
                                }
                            }
                        }
                    }
                }
            }
            };

            // Create a custom mail merge data source
            PaymentDetailsMailMergeDataSource dataSource = new PaymentDetailsMailMergeDataSource(paymentDetails);
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions;

            // Perform mail merge
            doc.MailMerge.UseNonMergeFields = true;
            doc.FieldOptions.LegacyNumberFormat = true;

            doc.MailMerge.FieldMergingCallback = new HandleMergeImageField(paymentDetails.ImageUrl);

            doc.MailMerge.ExecuteWithRegions(dataSource);
             
            // Save the merged document as PDF
            doc.Save($@"C:\Project\POC\POC\Receipts-{paymentDetails.InstanceId}.pdf", SaveFormat.Pdf);

            Console.WriteLine("Word document merged and saved successfully as PDF.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

public class HandleMergeImageField : IFieldMergingCallback
{
    private readonly string imageUrl;
    public HandleMergeImageField(string url)
    {
        imageUrl = url;
    }
    void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
    {
        // Do nothing.
    }
    static byte[] DownloadImage(string imageUrl)
    {
        using (WebClient client = new WebClient())
        {
            return client.DownloadData(imageUrl);
        }
    }

    /// <summary>
    /// This is called when a mail merge encounters a MERGEFIELD in the document with an "Image:" tag in its name.
    /// </summary>
    void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
    {
        byte[] imageBytes = DownloadImage(imageUrl);

        e.ImageStream = new MemoryStream(imageBytes);
    }
}

public class PaymentDetailsMailMergeDataSource : IMailMergeDataSource
{
    private readonly PaymentDetails paymentDetails;
    private int recordIndex;

    public PaymentDetailsMailMergeDataSource(PaymentDetails paymentDetails)
    {
        this.paymentDetails = paymentDetails;
        this.recordIndex = -1;
    }

    public string TableName => "paymentDetails";

    public bool GetValue(string fieldName, out object? fieldValue)
    {
        fieldValue = null;

        switch (fieldName)
        {
            case "InstanceId":
                fieldValue = paymentDetails.InstanceId;
                return true;
            case "Date":
                fieldValue = paymentDetails.Date;
                return true;
            case "Total":
                fieldValue = paymentDetails.Total;
                return true;
            case "GstTotal":
                fieldValue = paymentDetails.GstTotal;
                return true;
            case "MerchantFee":
                fieldValue = paymentDetails.MerchantFee;
                return true;
            case "TotalPayment":
                fieldValue = paymentDetails.TotalPayment;
                return true;
            case "Accounts":
                if (recordIndex < paymentDetails.Accounts.Count)
                {
                    fieldValue = paymentDetails.Accounts[recordIndex];
                    return true;
                }
                else
                {
                    return false;
                }
            default:
                return false;
        }
    }

    public bool MoveNext()
    {
        recordIndex++;
        return recordIndex < 1;
    }

    public IMailMergeDataSource? GetChildDataSource(string tableName)
    {
        if (tableName == "Accounts" && recordIndex < paymentDetails.Accounts.Count)
        {
            return new AccountMailMergeDataSource(paymentDetails.Accounts);
        }
        else
        {
            return null;
        }
    }
}

public class AccountMailMergeDataSource : IMailMergeDataSource
{
    private readonly List<Account> accounts;
    private int recordIndex;

    public AccountMailMergeDataSource(List<Account> accounts)
    {
        this.accounts = accounts;
        this.recordIndex = -1;
    }

    public string TableName => "Accounts";

    public bool GetValue(string fieldName, out object? fieldValue)
    {
        fieldValue = null;

        switch (fieldName)
        {
            case "ModuleReference":
                fieldValue = accounts[recordIndex].ModuleReference;
                return true;
            case "AccountLabel":
                fieldValue = accounts[recordIndex].AccountLabel;
                return true;
            case "AccountTotal":
                fieldValue = accounts[recordIndex].AccountTotal;
                return true;
            case "AccountGstTotal":
                fieldValue = accounts[recordIndex].AccountGstTotal;
                return true;

            case "ReceiptTypes":
                if (recordIndex < accounts[recordIndex].ReceiptTypes.Count)
                {
                    fieldValue = accounts[recordIndex];
                    return true;
                }
                else
                {
                    return false;
                }
            default:
                return false;
        }
    }

    public bool MoveNext()
    {
        recordIndex++;
        return recordIndex < accounts.Count;
    }

    public IMailMergeDataSource? GetChildDataSource(string tableName)
    {
        if (tableName == "ReceiptTypes")
        {
            return new ReceiptTypeMailMergeDataSource(accounts[recordIndex].ReceiptTypes);
        }
        else if (tableName == "Details")
        {
            return new DetailsMailMergeDataSource(accounts[recordIndex].Details);
        }
        else
        {
            return null;
        }
    }
}

public class ReceiptTypeMailMergeDataSource : IMailMergeDataSource
{
    private readonly List<ReceiptType> receiptTypes;
    private int recordIndex;
    public ReceiptTypeMailMergeDataSource(List<ReceiptType> receiptTypes)
    {
        this.receiptTypes = receiptTypes;
        this.recordIndex = -1;
    }

    public string TableName => "ReceiptTypes";

    public bool GetValue(string fieldName, out object? fieldValue)
    {
        fieldValue = null;

        switch (fieldName)
        {
            case "HasPaymentLines":
                fieldValue = receiptTypes[recordIndex].PaymentLines.Count > 0;
                return true;
            case "ReceiptTypeLabel":
                fieldValue = receiptTypes[recordIndex].ReceiptTypeLabel;
                return true;
            case "ReceiptTypeTotal":
                fieldValue = receiptTypes[recordIndex].ReceiptTypeTotal;
                return true;
            case "ReceiptTypeGstTotal":
                fieldValue = receiptTypes[recordIndex].ReceiptTypeGstTotal;
                return true;
            default:
                return false;
        }
    }

    public bool MoveNext()
    {
        recordIndex++;
        // ReceiptType doesn't have child data source
        return recordIndex < receiptTypes.Count;
    }

    public IMailMergeDataSource? GetChildDataSource(string tableName)
    {
        if (tableName == "PaymentLines")
        {
            if (recordIndex >= 0 && recordIndex < receiptTypes.Count)
            {
                return new PaymentLinesDataMailMergeDataSource(receiptTypes[recordIndex].PaymentLines);
            }
        }
        return null;
    }
}

public class PaymentLinesDataMailMergeDataSource : IMailMergeDataSource
{
    private readonly List<PaymentLinesData> paymentLines;
    private int recordIndex;

    public PaymentLinesDataMailMergeDataSource(List<PaymentLinesData> paymentLines)
    {
        this.paymentLines = paymentLines;
        this.recordIndex = -1;
    }

    public string TableName => "PaymentLines";

    public bool GetValue(string fieldName, out object? fieldValue)
    {
        fieldValue = null;

        if (recordIndex >= 0 && recordIndex < paymentLines.Count)
        {
            switch (fieldName)
            {
                case "LineLabel":
                    fieldValue = paymentLines[recordIndex].LineLabel;
                    return true;
                case "Amount":
                    fieldValue = paymentLines[recordIndex].Amount;
                    return true;
                default:
                    return false;
            }
        }
        return false;
    }

    public bool MoveNext()
    {
        recordIndex++;
        return recordIndex < paymentLines.Count;
    }

    public IMailMergeDataSource? GetChildDataSource(string tableName)
    {
        // PaymentLinesData doesn't have child data source
        return null;
    }
}

public class DetailsMailMergeDataSource : IMailMergeDataSource
{
    private readonly List<string> details;
    private int recordIndex;
    public DetailsMailMergeDataSource(List<string> details)
    {
        this.details = details;
        this.recordIndex = -1;
    }

    public string TableName => "Details";

    public bool GetValue(string fieldName, out object fieldValue)
    {
        fieldValue = null;

        if (details.Count > 0 && recordIndex < details.Count)
        {
            fieldValue = details[recordIndex];
            return true;
        }
        else
        {
            fieldValue = " ";
            return false; // Return false when there are no details available
        }
    }

    public bool MoveNext()
    {
        recordIndex++;
        // ReceiptType doesn't have child data source
        return recordIndex < details.Count;
    }

    public IMailMergeDataSource? GetChildDataSource(string tableName)
    {
        // ReceiptType doesn't have child data source
        return null;
    }
}



public class PaymentDetails
{
    public int InstanceId { get; set; }
    public string Date { get; set; }
    public double Total { get; set; }
    public double GstTotal { get; set; }
    public double MerchantFee { get; set; }
    public double TotalPayment { get; set; }
    public string ImageUrl { get; set; }
    public List<Account> Accounts { get; set; }
}

public class Account
{
    public string ModuleReference { get; set; }
    public string AccountLabel { get; set; }
    public List<string> Details { get; set; }
    public double AccountTotal { get; set; }
    public double AccountGstTotal { get; set; }
    public List<ReceiptType> ReceiptTypes { get; set; }
}

public class ReceiptType
{
    public string ReceiptTypeLabel { get; set; }
    public double ReceiptTypeTotal { get; set; }
    public double ReceiptTypeGstTotal { get; set; }
    public List<PaymentLinesData> PaymentLines { get; set; }
}

public class PaymentLinesData
{
    public string LineLabel { get; set; }

    public decimal Amount { get; set; }

    public decimal GstAmount { get; set; }
}