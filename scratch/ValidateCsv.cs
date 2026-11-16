using System;
using System.Collections.Generic;
using System.IO;

namespace JustGo.PreImportValidation
{
    class Program
    {
        // Sample CSV input as provided in Exhibit Q6
        private static readonly string CsvData = 
@"MemberID,EventID,PaymentStatus
1,101,Paid
2,102,Pending
3,101,Cancelled
4,0,Paid
-5,102,Paid
6,abc,Pending
7,101,paid
8,103,Refunded
9,104
10 , 102 , Pending
0,0,Unknown";

        // Allowed payment statuses per business rules and database schema
        private static readonly HashSet<string> AllowedStatuses = new HashSet<string>
        {
            "Paid",
            "Pending",
            "Cancelled"
        };

        static void Main(string[] args)
        {
            Console.WriteLine("=== JustGo Pre-Import Registration Validation Tool ===");
            Console.WriteLine();

            int totalRows = 0;
            int validRows = 0;
            List<string> invalidRowErrors = new List<string>();

            using (StringReader reader = new StringReader(CsvData))
            {
                string header = reader.ReadLine(); // Read and skip header row
                string line;
                int rowNumber = 0;

                while ((line = reader.ReadLine()) != null)
                {
                    rowNumber++;
                    totalRows++;

                    string errorMessage;
                    bool isValid = ValidateRow(line, out errorMessage);

                    if (isValid)
                    {
                        validRows++;
                    }
                    else
                    {
                        invalidRowErrors.Add(string.Format("Row {0}: {1}", rowNumber, errorMessage));
                    }
                }
            }

            // Summary Output
            Console.WriteLine("--- Validation Summary ---");
            Console.WriteLine(string.Format("Total rows processed: {0}", totalRows));
            Console.WriteLine(string.Format("Number of valid rows:  {0}", validRows));
            Console.WriteLine(string.Format("Number of invalid rows:{0}", invalidRowErrors.Count));
            Console.WriteLine();

            Console.WriteLine("--- Invalid Rows Details ---");
            foreach (string error in invalidRowErrors)
            {
                Console.WriteLine(error);
            }
        }

        /// <summary>
        /// Validates a single CSV record against business rules:
        /// 1. Expected 3 columns (MemberID, EventID, PaymentStatus).
        /// 2. MemberID and EventID must be positive integers (> 0).
        /// 3. PaymentStatus must be exactly 'Paid', 'Pending', or 'Cancelled'.
        /// </summary>
        public static bool ValidateRow(string line, out string errorMessage)
        {
            List<string> errors = new List<string>();

            if (string.IsNullOrWhiteSpace(line))
            {
                errorMessage = "Empty or whitespace line.";
                return false;
            }

            string[] parts = line.Split(',');
            if (parts.Length != 3)
            {
                errorMessage = string.Format("Malformed row. Expected 3 columns, found {0}", parts.Length);
                return false;
            }

            string memberIdStr = parts[0].Trim();
            string eventIdStr = parts[1].Trim();
            string paymentStatus = parts[2].Trim();

            // Validate MemberID: must be a positive integer (> 0)
            int memberId;
            if (!int.TryParse(memberIdStr, out memberId) || memberId <= 0)
            {
                errors.Add(string.Format("Invalid MemberID '{0}' (must be a positive integer)", memberIdStr));
            }

            // Validate EventID: must be a positive integer (> 0)
            int eventId;
            if (!int.TryParse(eventIdStr, out eventId) || eventId <= 0)
            {
                errors.Add(string.Format("Invalid EventID '{0}' (must be a positive integer)", eventIdStr));
            }

            // Validate PaymentStatus: strict matching against allowed statuses
            // [Assumption]: Case sensitivity is enforced to adhere strictly to database constraints.
            if (!AllowedStatuses.Contains(paymentStatus))
            {
                errors.Add(string.Format("Invalid PaymentStatus '{0}' (expected 'Paid', 'Pending', or 'Cancelled')", paymentStatus));
            }

            if (errors.Count > 0)
            {
                errorMessage = string.Join("; ", errors.ToArray());
                return false;
            }

            errorMessage = null;
            return true;
        }

        /*
        ===================================================================================
        PERSONALIZATION & REAL-WORLD EXTENSION COMMENT:
        -----------------------------------------------------------------------------------
        In my data engineering and machine learning projects, preprocessing large CSV datasets 
        or live streaming transactions cannot rely on in-memory buffers or basic console printing. 
        To adapt this validation utility for production-scale migrations:
        1. Memory-Efficient Streaming: Replace StringReader/ReadAllLines with an asynchronous 
           streaming pipeline (e.g. CsvHelper or async stream readers), reading records in chunks 
           with constant memory overhead.
        2. Referential & Relational Integrity Checks: Beyond syntactic checks, real migrations 
           must verify foreign key existence (verifying MemberID exists in Members and EventID in Events). 
           To prevent N+1 query bottlenecks against the database, I load valid ID key-sets into an 
           in-memory Bloom filter or Redis set, or pass chunked Table-Valued Parameters (TVPs) to SQL Server.
        3. Dead-Letter Quarantine & Audit Logging: Invalid rows should not halt the ingestion batch; 
           instead, they are routed to a 'quarantine_registrations' staging table or error CSV with 
           precise row indexes, raw source lines, and failure reasons for automated client remediation.
        ===================================================================================
        */
    }
}
