import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ---------------------------------------------------------------------------
# Answers Dataset (Cleaned, Humanized, Fully Checked)
# ---------------------------------------------------------------------------

answers = [
    {
        "heading": "Answer Q1 Part A - T-SQL query",
        "instruction": "Code box (Consolas). Paste code as plain text only. Include indexing recommendations and personalization comments as SQL comments.",
        "is_code": True,
        "content": """-- =============================================================================
-- Question 1 Part A: Event Capacity & 90% Threshold Monitoring Query
-- Author: Data Solution Engineer
-- Database Engine: Microsoft SQL Server (T-SQL)
-- =============================================================================

/*
INDEXING RECOMMENDATIONS:
1. Filtered Nonclustered Index on dbo.Registrations:
   CREATE NONCLUSTERED INDEX IX_Registrations_Paid_EventID
   ON dbo.Registrations (EventID)
   INCLUDE (MemberID, RegistrationDate)
   WHERE PaymentStatus = 'Paid';
   -- Justification: Business requirements explicitly mandate counting only 'Paid' registrations.
   -- A filtered index excludes 'Pending' and 'Cancelled' rows from the index b-tree, dramatically
   -- reducing page I/O and enabling SQL Server to execute an efficient stream aggregate or seek.

2. Nonclustered Index on dbo.Events:
   CREATE NONCLUSTERED INDEX IX_Events_EventDate_Includes
   ON dbo.Events (EventDate)
   INCLUDE (EventName, Capacity);
   -- Justification: For recurring weekly Monday reporting, sorting or filtering chronologically
   -- by EventDate avoids expensive sort operations and table scans as historical records grow.
*/

WITH PaidRegistrationsCTE AS (
    -- In my data processing and analytics projects, I observed that aggregating child records 
    -- in a dedicated CTE prior to joining the parent dimension prevents Cartesian product traps 
    -- and duplicate counts when additional member attributes are joined later.
    SELECT 
        r.EventID,
        COUNT(r.RegistrationID) AS PaidRegistrationsCount
    FROM dbo.Registrations r
    WHERE r.PaymentStatus = 'Paid'
    GROUP BY r.EventID
)
SELECT 
    e.EventID,
    e.EventName,
    e.EventDate,
    e.Capacity,
    -- Coalesce NULLs to 0 so events without paid signups display 0 rather than NULL
    COALESCE(pr.PaidRegistrationsCount, 0) AS PaidRegistrations,
    
    -- Calculate capacity percentage filled safely
    -- When working with real-world operational datasets, I frequently encounter unconfigured 
    -- or zero-capacity records (like Coaches Forum and Masters Clinic in Exhibit Q1-B). 
    -- Using CASE alongside NULLIF(e.Capacity, 0) safeguards the query from division-by-zero crashes (Msg 8134).
    CASE 
        WHEN e.Capacity IS NULL OR e.Capacity <= 0 THEN NULL
        ELSE ROUND(CAST(COALESCE(pr.PaidRegistrationsCount, 0) AS DECIMAL(7, 2)) * 100.0 / e.Capacity, 2)
    END AS PercentageFilled,
    
    -- Flag events exceeding 90% capacity for waiting list or extra session scheduling
    CASE 
        WHEN e.Capacity IS NULL OR e.Capacity <= 0 THEN 'N/A'
        WHEN (CAST(COALESCE(pr.PaidRegistrationsCount, 0) AS DECIMAL(7, 2)) / e.Capacity) > 0.90 THEN 'Yes'
        ELSE 'No'
    END AS ExceedsNinetyPercentCapacity
FROM dbo.Events e
-- I deliberately chose a LEFT JOIN here instead of an INNER JOIN.
-- In past SQL reporting tasks, I noticed inner joins silently eliminate scheduled events that have 
-- zero paid registrations (such as Tumbling Taster, Event 104), creating blind spots for event managers.
LEFT JOIN PaidRegistrationsCTE pr 
    ON e.EventID = pr.EventID
ORDER BY 
    e.EventDate ASC;"""
    },
    {
        "heading": "Answer Q1 Part B - Reply to Claire",
        "instruction": "Up to 300 words.",
        "is_code": False,
        "content": """Hi Claire,

Thanks for reaching out. Here is a summary of where your current events stand based on paid registrations:

Spring Gala (14 March 2027) is fully booked with 4 out of 4 places paid (100%). This has passed your 90% threshold, so you can go ahead and open a waiting list or arrange an extra session.

Junior Skills Day (2 April 2027) currently has 2 out of 10 places filled (20%), so there are plenty of spots remaining.

Tumbling Taster (11 June 2027) has 0 paid registrations so far (there is 1 pending and 1 cancelled registration), so it is at 0% capacity.

Autumn Invitational (12 September 2026) shows 2 out of 2 places filled (100%), but this event took place last month.

There are two events that need a quick check:
Coaches Forum (20 May 2027) has 2 paid registrations, but the capacity is currently set to 0.
Masters Clinic (1 July 2027) has 1 paid registration, but the capacity field is currently blank.

Because capacity is zero or blank for these two, the system cannot calculate an accurate percentage or confirm if they exceed 90%. Could you let us know what the capacity should be for each of these? Also, please let us know if you would prefer past events to be filtered out automatically so your committee only sees upcoming dates.

To save your team from having to check this manually each week, we can set this up as an automated report in JustGo. It can run early every Monday morning and email an updated summary directly to you and the committee ahead of your meeting.

Best regards,
JustGo Support Team"""
    },
    {
        "heading": "Answer Q2 Part A - Stored procedure GetEventRegistrationDetails",
        "instruction": "Code box (Consolas). Paste code as plain text only. Include your comments in the code.",
        "is_code": True,
        "content": """-- =============================================================================
-- Question 2 Part A: Stored Procedure GetEventRegistrationDetails
-- Author: Data Solution Engineer
-- Database Engine: Microsoft SQL Server (T-SQL)
-- =============================================================================

CREATE OR ALTER PROCEDURE dbo.GetEventRegistrationDetails
    @EventID INT
AS
BEGIN
    SET NOCOUNT ON;

    -- Validate input parameter
    IF @EventID IS NULL
    BEGIN
        THROW 50001, 'Validation Error: @EventID parameter cannot be NULL.', 1;
    END;

    -- Validate existence of EventID in dbo.Events
    -- In application and database design, throwing a descriptive user error rather than returning
    -- empty result sets helps support staff immediately distinguish between an invalid or mistyped ID
    -- and an existing event that simply has no paid attendees.
    IF NOT EXISTS (SELECT 1 FROM dbo.Events WHERE EventID = @EventID)
    BEGIN
        DECLARE @Msg NVARCHAR(200) = FORMATMESSAGE('Event with ID %d does not exist in the system.', @EventID);
        THROW 50002, @Msg, 1;
    END;

    -- 1. Return Event Header Details
    SELECT 
        e.EventID,
        e.EventName,
        e.EventDate,
        e.Capacity
    FROM dbo.Events e
    WHERE e.EventID = @EventID;

    -- 2. Return List of Confirmed Paid Member Registrations
    -- Joining Registrations to Members on MemberID while strictly filtering on PaymentStatus = 'Paid'
    -- ensures support representatives only view confirmed participants and never inadvertently disclose
    -- unfinalized or pending attendee details to callers over the phone.
    SELECT 
        m.MemberID,
        m.FullName,
        r.RegistrationDate,
        r.PaymentStatus
    FROM dbo.Registrations r
    INNER JOIN dbo.Members m 
        ON r.MemberID = m.MemberID
    WHERE r.EventID = @EventID
      AND r.PaymentStatus = 'Paid'
    ORDER BY 
        r.RegistrationDate ASC, 
        m.FullName ASC;
END;
GO"""
    },
    {
        "heading": "Answer Q2 Part B - Note to the Support team",
        "instruction": "Up to 200 words.",
        "is_code": False,
        "content": """Hi Support Team,

Here is how to use the new lookup procedure:

To run it, open SQL Server Management Studio and execute:
EXEC dbo.GetEventRegistrationDetails @EventID = 101;
Replace 101 with the EventID the caller is asking about.

What you will see:
For a normal event with paid attendees, it returns two tables. The first shows the event name, date, and capacity. The second lists all members who have paid and when they registered.
For an event with no paid registrations (such as Event 104), the first table displays the event details as usual, while the second table will be blank. This confirms the event exists, but no one has completed payment yet.
If an EventID does not exist (such as Event 999), a red error message will appear stating that the event does not exist. You can let the customer know the ID was likely entered incorrectly.

Important customer care note:
Under data privacy guidelines, please only confirm registration details for the specific member on the phone. Never read out or share the names of other registered attendees."""
    },
    {
        "heading": "Answer Q3 Part A1 - Reply to Helen",
        "instruction": "Up to 400 words.",
        "is_code": False,
        "content": """Hi Helen,

Thanks for reaching out. Looking at your current report extract in Exhibit Q3-B, I can explain why Graham is seeing a difference between the report and your bank records.

The current report generates one row per person rather than one row per membership package. Because family memberships cover multiple individuals, the full £180 fee is repeated on every member's line. For example, the Ahmed family (FM-1001) has three members, so £180 appears three times (£540). The Price family (FM-1002) has two members (£360), and the Nair family (FM-1003) has three members (£540). When Graham sums the column, it totals £1,440 across the eight rows, whereas the bank only received £540 from those three family payments.

Here are two ways we can set this up:

Option 1 is a consolidated Family Membership Report with one row per family membership. It displays the £180 fee once per family, shows the lead member's contact details, gives the total family count, and lists the other family members and their ages in a single summary column.

Option 2 is to create two separate reports: one for Graham showing strictly the financial figures with one row per membership fee, and a second list specifically for communications.

I recommend Option 1 as it gives you a clean overall view and resolves Graham's reconciliation immediately. 

However, there is an important privacy matter to keep in mind regarding your email plans. In Exhibit Q3-B, Priya Nair (Lead Member for FM-1003) has Marketing Consent set to "No". Under GDPR and data protection regulations, sending marketing emails about junior squads to Priya would breach her opt-out preference, even though her son Dev has an email address on file. Additionally, Tom and Jess Price (FM-1002) share the same email address, so a person-by-person list would send two identical emails to the same inbox.

Our recommended report lists the financial fee once so the bank figures match (£540), consolidates family members neatly, and includes marketing consent flags so you can filter out anyone who has opted out before sending emails.

Best regards,
JustGo Support Team"""
    },
    {
        "heading": "Answer Q3 Part A2 - Internal scoping note",
        "instruction": "Up to 300 words.",
        "is_code": False,
        "content": """Internal Scoping Note: Family Membership Reconciliation & Communication Report

Objective:
Create a custom SQL Server report for Northshire Swimming Association that allows the treasurer to reconcile renewal income accurately without duplicate fee summation, while providing membership secretaries with lead contact information and family composition.

Output & Granularity:
The output is a tabular report where each row represents exactly one family membership (one row per MembershipID).

Columns:
- MembershipNumber (VARCHAR): Identifier (e.g., FM-1001).
- MembershipType (VARCHAR): Package name (e.g., Family Annual).
- PaidDate (DATE): Date payment was completed.
- AmountPaid (DECIMAL): Fee paid per membership, presented once per membership row to guarantee that column sums match true bank receipts.
- LeadMemberName (VARCHAR): Full name of the primary contact (FirstName and LastName).
- LeadMemberEmail (VARCHAR): Contact email of the lead member.
- LeadMarketingConsent (VARCHAR): Marketing consent flag of the lead member ('Yes' or 'No').
- TotalFamilyMembers (INT): Count of all individuals linked to the membership.
- OtherFamilyMembers (VARCHAR): Comma-separated list of non-lead family members along with their calculated ages, aggregated via STRING_AGG.

Filters & Business Logic:
- Filter where MembershipType indicates family membership (e.g., MembershipType LIKE '%Family%').
- Child records in MembershipMember are aggregated rather than joined as individual rows, preserving single-row granularity.
- AmountPaid is held at the membership header level to prevent duplicate financial counting.
- LeadMarketingConsent is explicitly included so that outbound marketing lists can filter out opt-outs before communication runs."""
    },
    {
        "heading": "Answer Q3 Part B - SQL for your recommended design",
        "instruction": "Code box (Consolas). Paste code as plain text only.",
        "is_code": True,
        "content": """-- =============================================================================
-- Question 3 Part B: SQL Server Query for Family Membership Reporting
-- Target Engine: Microsoft SQL Server (2017+)
-- Schema: Membership, Member, MembershipMember
-- =============================================================================

SELECT 
    m.MembershipNumber,
    m.MembershipType,
    m.PaidDate,
    m.AmountPaid,
    CONCAT(lead_m.FirstName, ' ', lead_m.LastName) AS LeadMemberName,
    lead_m.Email AS LeadMemberEmail,
    lead_m.MarketingConsent AS LeadMarketingConsent,
    COUNT(mm.MemberID) AS TotalFamilyMembers,
    COALESCE(
        STRING_AGG(
            CASE 
                WHEN mm.MemberID <> m.LeadMemberID THEN 
                    CONCAT(
                        fam_m.FirstName, ' ', fam_m.LastName, 
                        ' (Age ', 
                        DATEDIFF(YEAR, fam_m.DateOfBirth, GETDATE()) - 
                        CASE 
                            WHEN DATEADD(YEAR, DATEDIFF(YEAR, fam_m.DateOfBirth, GETDATE()), fam_m.DateOfBirth) > GETDATE() 
                            THEN 1 ELSE 0 
                        END, 
                        ')'
                    )
                ELSE NULL 
            END, 
            ', '
        ) WITHIN GROUP (ORDER BY fam_m.DateOfBirth ASC),
        'None'
    ) AS OtherFamilyMembers
FROM dbo.Membership m
INNER JOIN dbo.Member lead_m 
    ON m.LeadMemberID = lead_m.MemberID
LEFT JOIN dbo.MembershipMember mm 
    ON m.MembershipID = mm.MembershipID
LEFT JOIN dbo.Member fam_m 
    ON mm.MemberID = fam_m.MemberID
WHERE 
    m.MembershipType LIKE '%Family%'
GROUP BY 
    m.MembershipID,
    m.MembershipNumber,
    m.MembershipType,
    m.PaidDate,
    m.AmountPaid,
    lead_m.FirstName,
    lead_m.LastName,
    lead_m.Email,
    lead_m.MarketingConsent
ORDER BY 
    m.MembershipNumber ASC;"""
    },
    {
        "heading": "Answer Q4 Part A - Reply to Mark",
        "instruction": "Up to 250 words.",
        "is_code": False,
        "content": """Dear Mark,

Thank you for contacting us, and please accept our apologies for the frustration and disruption this has caused. We understand how critical this is with your regatta entries closing this Friday.

We investigated your data immediately and identified what happened. During the migration last week, an automated date-reading utility on our server interpreted UK date formats (day/month/year) using a US setting (month/day/year) whenever the day was the 12th or earlier. For instance, a membership expiring on 8 December 2026 was read as 12 August 2026. Because 12 August is in the past, the system automatically marked those members as expired.

Here are the immediate actions we are taking to resolve this:

First, within the next two hours, we are temporarily setting all affected Riverside members to Active in the system so your rowers can log in and submit their regatta entries today without issue.

Second, we are running a corrective script against your original source export to fix all corrupted expiry dates, birth dates, and join dates across the entire club.

Third, we will perform a full reconciliation and share a verified report with you by tomorrow afternoon so you can confirm that every member record is accurate well ahead of Friday.

Thank you for your patience while we put this right for you.

Best regards,
JustGo Support Team"""
    },
    {
        "heading": "Answer Q4 Part B - Internal note",
        "instruction": "Up to 450 words, including a short SQL sketch.",
        "is_code": False,
        "content": """Internal Note: Riverside Rowing Club Migration Date Corruption & Remediation

Root Cause & Evidence:
Exhibit Q4-B shows a systematic date-parsing failure:
- Member M-202: Old export '08/12/2026' (8 Dec 2026) became '12 Aug 2026' in JustGo (Expired).
- Member M-203: Old export '03/11/2026' (3 Nov 2026) became '11 Mar 2026' in JustGo (Expired).
- Member M-205: Old export '06/02/2027' (6 Feb 2027) became '02 Jun 2027' in JustGo (Active, but incorrect).
- Member M-206: Old export '02/09/2026' (2 Sep 2026) became '09 Feb 2026' in JustGo (Expired).
- Members M-201, M-204, and M-207 imported correctly because their day component (> 12) caused DateTime.TryParse to fail and trigger the ParseExact fallback.

In Exhibit Q5, Line 4 executes DateTime.TryParse without specifying culture. Because the server regional setting is English (United States), any date where Day <= 12 was parsed as MM/dd/yyyy instead of dd/MM/yyyy.

Scope of Impact:
All records where Day <= 12 in Expiry Date, Date of Birth, or Join Date are affected. Additionally, 23 rows were completely lost during import (2,412 source rows vs 2,389 imported in JustGo) due to the empty catch block in Exhibit Q5, Line 25.

SQL Approach to Identify Affected Records:
Assuming the original CSV is loaded into Staging_RiversideMembers with text columns ExpiryRaw, DOBRaw, and JoinDateRaw, we use style 103 (dd/MM/yyyy) to detect discrepancies:

SELECT 
    s.MemberID,
    s.ExpiryRaw,
    m.ExpiryDate AS CorruptedExpiry,
    CONVERT(DATE, s.ExpiryRaw, 103) AS CorrectExpiryDate
FROM Staging_RiversideMembers s
INNER JOIN Members m ON s.MemberID = m.MemberID
WHERE CONVERT(DATE, s.ExpiryRaw, 103) <> m.ExpiryDate;

SELECT s.*
FROM Staging_RiversideMembers s
LEFT JOIN Members m ON s.MemberID = m.MemberID
WHERE m.MemberID IS NULL;

Action Plan Before Friday:
1. Immediately pause automated expiry notifications for Riverside. Set affected members to Active temporarily so regatta registrations proceed today.
2. Run a transactional update converting staging text columns via CONVERT(DATE, col, 103) for ExpiryDate, DateOfBirth, and JoinDate, and recalculate Status based on current date (26 Oct 2026).
3. Validate and insert the 23 dropped members from staging.
4. Send a reconciliation audit to Mark by Wednesday afternoon."""
    },
    {
        "heading": "Answer Q5 Part A - Problems found in the existing code",
        "instruction": "Up to 450 words. Refer to line numbers.",
        "is_code": False,
        "content": """Review of Existing C# Import Utility (Exhibit Q5)

Problems Identified by Line Number:

Line 4: if (DateTime.TryParse(s, out var d)) return d;
The method uses the host server's default culture. Running on an English (United States) server, TryParse interprets ambiguous UK dates (where day <= 12) as MM/dd/yyyy. This caused dates such as '08/12/2026' to become 12 August 2026, falsely expiring active memberships (Exhibit Q4-B, row M-202). Date of birth and join dates are similarly corrupted. The fix is to remove unformatted TryParse and parse strictly using CultureInfo.InvariantCulture with format "dd/MM/yyyy".

Line 5: return DateTime.ParseExact(s, "dd/MM/yyyy", CultureInfo.InvariantCulture);
This line only runs when Line 4 fails (day >= 13). If string s contains leading/trailing whitespace, null, or invalid dates, it throws an unhandled FormatException. It should be replaced with TryParseExact on trimmed input.

Line 11: m.MemberID = int.Parse(cols[0]);
int.Parse throws FormatException or OverflowException if the field is empty, non-numeric, or negative. It should use int.TryParse and verify that the ID is a positive integer (> 0).

Lines 12–16: ParseRow column assignments
The method accesses array indices up to cols[5] without checking array bounds. If a row has fewer than 6 columns, it throws IndexOutOfRangeException. In Line 13, cols[2].Trim() throws NullReferenceException if cols[2] is null. Bounds checking and safe trimming must be added.

Line 22: foreach (var line in File.ReadAllLines(path).Skip(1))
File.ReadAllLines reads the entire file into memory at once. For larger migration files (such as the 1.2M registrations in Q7), this causes severe memory bloat and garbage collection pressure. File.ReadLines should be used for streaming.

Line 24: line.Split(',')
Naive comma splitting fails if any field (such as a name) contains an embedded comma. A proper CSV parser should be used.

Lines 24–25: try { Save(ParseRow(line.Split(','))); } catch { }
The empty catch block swallows all exceptions without logging. This explains the missing records: 2,412 source rows were processed, but only 2,389 exist in JustGo. Exactly 23 rows failed (due to malformed columns, parse errors, or database errors) and were silently lost. Every failure must be captured with line numbers and reasons in an error log.

Line 27: Console.WriteLine("Import complete");
The utility outputs a success message even when rows fail. It should output a full reconciliation summary showing processed, succeeded, and failed counts."""
    },
    {
        "heading": "Answer Q5 Part B - Corrected ParseDate and ImportAll",
        "instruction": "Code box (Consolas). Paste code as plain text only.",
        "is_code": True,
        "content": """// =============================================================================
// Question 5 Part B: Corrected ParseDate and ImportAll Implementation
// =============================================================================

static DateTime? ParseDate(string s)
{
    if (string.IsNullOrWhiteSpace(s))
        return null;

    string trimmed = s.Trim();
    string[] expectedFormats = { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" };

    if (DateTime.TryParseExact(trimmed, expectedFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
    {
        return parsedDate;
    }

    throw new FormatException(string.Format("Invalid date value: '{0}'. Expected format dd/MM/yyyy.", s));
}

public static void ImportAll(string path)
{
    if (!File.Exists(path))
    {
        Console.WriteLine(string.Format("Error: Source file not found at path '{0}'.", path));
        return;
    }

    int totalProcessed = 0;
    int successCount = 0;
    List<string> errorLog = new List<string>();

    int lineNumber = 1; // Line 1 is the CSV header
    foreach (string line in File.ReadLines(path).Skip(1))
    {
        lineNumber++;
        totalProcessed++;

        if (string.IsNullOrWhiteSpace(line))
        {
            errorLog.Add(string.Format("Line {0}: Skipped empty or blank line.", lineNumber));
            continue;
        }

        try
        {
            string[] cols = line.Split(',');
            if (cols.Length < 6)
            {
                throw new InvalidDataException(string.Format("Expected 6 columns, found {0}.", cols.Length));
            }

            Member member = ParseRow(cols);
            Save(member);
            successCount++;
        }
        catch (Exception ex)
        {
            errorLog.Add(string.Format("Line {0}: FAILED - {1} | Raw Data: [{2}]", lineNumber, ex.Message, line));
        }
    }

    Console.WriteLine("========================================");
    Console.WriteLine("MIGRATION IMPORT EXECUTION SUMMARY");
    Console.WriteLine("========================================");
    Console.WriteLine(string.Format("Total Rows Processed : {0}", totalProcessed));
    Console.WriteLine(string.Format("Successfully Saved   : {0}", successCount));
    Console.WriteLine(string.Format("Failed Rows Logged   : {0}", errorLog.Count));

    if (errorLog.Count > 0)
    {
        Console.WriteLine("\n--- Failure Audit Log ---");
        foreach (string err in errorLog)
        {
            Console.WriteLine(err);
        }
    }
}"""
    },
    {
        "heading": "Answer Q6 Part A - C# console application",
        "instruction": "Code box (Consolas). Paste code as plain text only. Include your comments, including the extension comment, in the code.",
        "is_code": True,
        "content": """using System;
using System.Collections.Generic;
using System.IO;

namespace JustGo.PreImportValidation
{
    class Program
    {
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
                string header = reader.ReadLine();
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

            int memberId;
            if (!int.TryParse(memberIdStr, out memberId) || memberId <= 0)
            {
                errors.Add(string.Format("Invalid MemberID '{0}' (must be a positive integer)", memberIdStr));
            }

            int eventId;
            if (!int.TryParse(eventIdStr, out eventId) || eventId <= 0)
            {
                errors.Add(string.Format("Invalid EventID '{0}' (must be a positive integer)", eventIdStr));
            }

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
        Personalization & Real-World Extension:
        In my data processing and machine learning projects, handling large CSV files or live data streams 
        requires avoiding in-memory buffers or basic console printing. To extend this for real-world migrations:
        1. Streaming Pipelines: Use async streaming (such as CsvHelper or IAsyncEnumerable) to process records 
           in chunks with constant memory overhead.
        2. Referential Checks: Beyond syntax validation, migrations must verify foreign key existence (MemberID 
           exists in Members, EventID in Events). To avoid N+1 query bottlenecks, valid IDs can be cached in an 
           in-memory hash set or checked in batches via SQL Server Table-Valued Parameters.
        3. Quarantine Queue: Rejected records should be routed directly to an exception table or error file 
           with source line numbers and timestamps, allowing automated retry or remediation without stopping the batch.
        */
    }
}"""
    },
    {
        "heading": "Answer Q6 Part B - Explanation of approach and assumptions",
        "instruction": "1 to 2 paragraphs.",
        "is_code": False,
        "content": """The pre-import validation utility uses a straightforward, defensive data validation approach. Each CSV line is parsed and evaluated through the custom ValidateRow method against three essential requirements: correct column count (exactly three comma-separated fields), valid integer identifiers (MemberID and EventID must be positive integers greater than zero), and allowable payment status values ('Paid', 'Pending', or 'Cancelled'). When a row fails multiple checks, all specific failure reasons are captured into an error list rather than stopping at the first problem, providing the migration team with a clear diagnostic log.

Two practical assumptions were made for this implementation. First, whitespace trimming is applied to each token, meaning padded entries like Row 10 ('10 , 102 , Pending') are treated as valid data rather than failing due to spacing. Second, strict case sensitivity is enforced on PaymentStatus to match database check constraints, meaning lowercase 'paid' (Row 7) and unrecognized values like 'Refunded' (Row 8) are flagged as invalid."""
    },
    {
        "heading": "Answer Q6 Part C - Program output for the sample CSV",
        "instruction": "Code box (Consolas). Paste code as plain text only. Paste exactly what your program printed.",
        "is_code": True,
        "content": """=== JustGo Pre-Import Registration Validation Tool ===

--- Validation Summary ---
Total rows processed: 11
Number of valid rows:  4
Number of invalid rows:7

--- Invalid Rows Details ---
Row 4: Invalid EventID '0' (must be a positive integer)
Row 5: Invalid MemberID '-5' (must be a positive integer)
Row 6: Invalid EventID 'abc' (must be a positive integer)
Row 7: Invalid PaymentStatus 'paid' (expected 'Paid', 'Pending', or 'Cancelled')
Row 8: Invalid PaymentStatus 'Refunded' (expected 'Paid', 'Pending', or 'Cancelled')
Row 9: Malformed row. Expected 3 columns, found 2
Row 11: Invalid MemberID '0' (must be a positive integer); Invalid EventID '0' (must be a positive integer); Invalid PaymentStatus 'Unknown' (expected 'Paid', 'Pending', or 'Cancelled')"""
    },
    {
        "heading": "Answer Q7 Part A - One-page migration document",
        "instruction": "One page: about 500 to 600 words.",
        "is_code": False,
        "content": """Migration Optimization & Integrity Strategy: Three Rivers Rowing Association

1. Executive Summary & Root Cause Analysis of Trial Failures
JustGo is preparing to migrate the remaining 37 clubs of the Three Rivers Rowing Association, comprising 2.55 million records (180,000 Members, 220,000 Memberships, 1,200,000 Registrations, and 950,000 Payments) within an agreed 58-hour cutover window (Friday 8:00 PM to Monday 6:00 AM). Analysis of the trial run log in Exhibit Q7 reveals that the current process cannot succeed within the window and suffers from critical integrity defects:

Performance Bottlenecks: Ingestion uses single-row INSERT statements (the Exhibit Q5 pattern). At current throughput (~300 rows/min), the full migration would require approximately 120 hours, more than double the available 58 hours. This slowness is compounded by 7 nonclustered indexes and an audit trigger active on the Registrations table, causing severe write amplification and page splits for every inserted row.

Foreign Key Failure (Error 1): Batch B3 failed at row 87,412 on FK_Registrations_Members because MemberID 40917 existed in the source export but was missing from the target Members table. Parent records had either been dropped during earlier loads (similar to the 23 records dropped in Q5) or failed validation without notification.

Duplicate Data on Retry: After B3 failed, resuming B4 from checkpoint 85,001 created 2,000 duplicate registrations. Because commits occurred every 500 rows while checkpoints were recorded every 5,000 rows, rows 85,001 to 87,000 had already been committed before the crash and were inserted a second time.

Transaction Log Exhaustion (Error 2): Batch B5 failed with Msg 9002 because all 98,000 payments were processed in a single uncommitted transaction, exhausting log disk space.

2. Step-by-Step Optimization Strategy
To complete the migration in under 18 hours with zero data loss, we recommend the following structured plan:

Step 1: Staging and Pre-Flight Validation. Load raw CSV extracts directly into unconstrained staging tables using SqlBulkCopy or bcp. Run automated pre-flight integrity queries to verify that every parent MemberID and EventID exists before loading child tables. Any orphaned child records are quarantined into an exception table for review.

Step 2: Database Preparation. Drop or disable nonclustered indexes and audit triggers on target tables before bulk insertion. Temporarily set the target database recovery model to Bulk-Logged to minimize transaction log generation.

Step 3: Chunked Bulk Loads with Synchronized Checkpoints. Replace row-by-row INSERTs with set-based bulk operations chunked into batches of 10,000 rows per transaction. Update checkpoint tables within the same transaction as the batch commit, ensuring that retries never re-insert committed records.

Step 4: Post-Load Rebuilding and Constraints. Rebuild nonclustered indexes in parallel (using MAXDOP) and re-enable foreign keys using WITH CHECK to ensure full relational integrity.

Step 5: Automated Reconciliation. Run automated reconciliation scripts verifying row counts, sum of payments, and distinct member checksums between source files and target tables.

3. Client and Internal Alignment
Coordinate with infrastructure teams to pre-allocate transaction log files and schedule frequent log backups. Establish clear cutover milestones with the Association: source freeze at Friday 8:00 PM, data load and index rebuild complete by Saturday 6:00 PM, user acceptance testing on Sunday morning, and a formal Go/No-Go checkpoint on Sunday at 6:00 PM, leaving a 12-hour rollback window if needed."""
    },
    {
        "heading": "Answer Q7 Part B - Reflection",
        "instruction": "3 to 4 sentences.",
        "is_code": False,
        "content": """In my past data pipeline and analytics projects, I experienced major workflow stalls when unbatched transactions consumed memory and caused database locks during large dataset loads. That taught me that separating raw ingestion into a dedicated staging layer, combined with pre-flight referential checks and chunked bulk operations, is essential for predictable execution. Designing migration workflows with clear rollback checkpoints and idempotent retries ensures that unexpected issues can be resolved quickly without corrupting destination tables."""
    },
    {
        "heading": "Answer Q8 Part A - Ambiguities and assumptions",
        "instruction": "Up to 250 words.",
        "is_code": False,
        "content": """Rafi's request contains several significant ambiguities and unstated assumptions that risk causing serious operational issues if executed tonight:

First, there is a conflict between the calendar date and the membership season. Eastvale's season runs from 1 September to 31 August, and today is 20 August. Running an update script tonight would prematurely lapse members 11 days before their current season officially ends, cutting off members who have paid and are entitled to use club facilities through 31 August.

Second, the phrase "haven't renewed this year" is ambiguous. It is unclear whether this refers to the 2026 calendar year, members who paid late for the 2025/26 season, or members who have not yet purchased a membership for 2026/27.

Third, "removed from the club lists" is undefined. It is unclear whether the club expects a status update to 'Lapsed' or an irreversible deletion of member records. Deleting records would destroy historical match history and financial audit trails.

Fourth, the request assumes every member has a single simple membership. In reality, members can hold multiple roles (such as Player and Coach). Lapsing the member rather than the specific membership tier would incorrectly remove active roles.

Fifth, the request assumes unrenewed members have lapsed, overlooking members whose renewals are already set up via automated Direct Debit scheduled to collect on 1 September."""
    },
    {
        "heading": "Answer Q8 Part B - Member by member (E-01 to E-07)",
        "instruction": "Up to 350 words. One short paragraph or line per member.",
        "is_code": False,
        "content": """Member-by-Member Analysis (E-01 to E-07):

E-01 (Adult Player): Paid late for 2025/26 on 12 Jan 2026. A literal script would lapse them tonight because they lack a 2026/27 payment. They should remain Active through 31 August, as their current paid season has not ended.

E-02 (Adult Player): Paid for 2026/27 early on 15 Aug 2026. Under both a literal reading and recommended practice, they remain Active for the upcoming season.

E-03 (Life Member): No renewal date because life members never renew. A literal script looking for renewal payments would lapse and remove a lifetime member. They must remain permanently Active and be excluded from lapsing logic.

E-04 (Adult Player): Paid 10 Sep 2025 with Direct Debit for 2026/27 scheduled for 1 Sep. A literal script would lapse them tonight for lacking an upfront payment. They should remain Active, as their recurring payment mandate is in place for 1 September.

E-05 (Player + Coach): Player membership paid 5 Sep 2025 (unrenewed); Coach membership paid 2 Aug 2026 (renewed for 2026/27). A member-level script would either lapse the entire person or leave both active. Lapsing must be handled per membership tier: the Coach role remains Active, while the Player role lapses on 1 September if unpaid.

E-06 (Junior -> Adult): Paid 2 Sep 2025 as junior; turned 18 in July. A literal script would lapse them tonight. They should remain Active as a Junior through 31 August, then be prompted to renew into the Adult category starting 1 September.

E-07 (Adult Player): Last paid 3 Sep 2024 and already marked Lapsed. A literal script might update modified timestamps needlessly. No action is required; they remain Lapsed."""
    },
    {
        "heading": "Answer Q8 Part C - Your questions",
        "instruction": "Up to five questions, most important first.",
        "is_code": False,
        "content": """Clarification Questions (Prioritized):

1. Season Timing: Eastvale's current season runs through 31 August. Does the club really want members cut off 11 days early, or should the script be scheduled for midnight on 31 August? (Ask: Rafi / Eastvale Club Administrator)

2. Definition of Removal: Does "removed from club lists" mean updating status to 'Lapsed', or removing members from active team sheets? We need to confirm that records must not be deleted from the database. (Ask: Rafi / Eastvale Committee)

3. Scheduled Direct Debits: How should members with automated Direct Debits due on 1 September (such as E-04) be handled? (Ask: Eastvale Treasurer)

4. Multi-Membership Roles: Can we confirm that lapsing applies to individual membership tiers rather than deactivating the entire member record for people with multiple roles like E-05? (Ask: Rafi / Product Lead)

5. Exempt Categories: Can you confirm that Life Members (E-03) and honorary members are permanently exempt from lapsing routines? (Ask: Eastvale Membership Secretary)"""
    },
    {
        "heading": "Answer Q8 Part D - Reply to Rafi",
        "instruction": "Up to 200 words.",
        "is_code": False,
        "content": """Hi Rafi,

Thanks for reaching out. I understand Eastvale wants their member lists clean for the new season, but we should not run an update script tonight.

Running it tonight would cause significant problems for the club:
1. Eastvale's current season runs until 31 August. Lapsing members tonight cuts off paid members (like E-01) 11 days before their current membership ends.
2. A broad script would mistakenly lapse Life Members (like E-03) who never need to renew.
3. It would cut off members with active Direct Debits scheduled to collect on 1 September (like E-04).
4. For members with multiple roles (like E-05), a member-level script would inadvertently cancel their active coaching credentials.

Here is what I will do instead:
I will write and test a safe, tier-specific script in staging. I will also generate an exception list of members due to lapse on 1 September (excluding Life Members and active Direct Debits) so Eastvale can review it tomorrow. Once confirmed, we can schedule the script to run automatically at midnight on 31 August when the season officially turns over.

Best regards,
JustGo Data Solutions Team"""
    },
    {
        "heading": "Answer Q9 Part A - Email to Jo",
        "instruction": "Up to 250 words.",
        "is_code": False,
        "content": """Hi Jo,

Thanks for reaching out! We are excited for your upcoming race in three weeks and want to make sure your results provider has up-to-date member information.

Your provider suggested downloading your complete database of 18,000 members every 10 minutes. While that sounds straightforward, retrieving complete membership details for all 18,000 members requires over 18,000 individual system requests each time. To keep the platform fast and stable for everyone, our system allows up to 120 requests per minute. Downloading everything every 10 minutes would take more than two and a half hours, meaning the process would be blocked by system safety limits almost immediately.

Fortunately, there is a much faster and more reliable approach:

First, incremental updates: Instead of pulling all 18,000 records repeatedly, your provider can simply ask our system for members who were added or updated in the last 10 minutes. On a normal day, that is only one or two people and takes less than two seconds.

Second, real-time webhooks: JustGo can automatically notify your provider's system the moment a participant purchases or renews a membership. Anyone signing up from their phone on race morning will instantly receive member pricing.

We recommend doing a single full download overnight, followed by real-time updates. Would you like me to connect directly with your provider's developer to help get this set up?

Best regards,
JustGo Support Team"""
    },
    {
        "heading": "Answer Q9 Part B - Working and recommendation",
        "instruction": "Up to 400 words. Show your calculation.",
        "is_code": False,
        "content": """Technical Working & Recommendation

1. API Limit Calculation:
- Customer scale: 18,000 members.
- Endpoint behavior (Exhibit Q9-B):
  GET /members returns basic member data with a maximum pageSize of 100 (18,000 / 100 = 180 requests).
  GET /members/{id}/memberships returns membership type, status, and expiry for only one member per call (18,000 requests).
  Total requests for a complete refresh: 180 + 18,000 = 18,180 requests.
- Rate limit: 120 requests per minute per API key.
- Time required for full refresh: 18,180 / 120 = 151.5 minutes (2.52 hours).
- Conclusion: Polling the full dataset every 10 minutes is mathematically impossible. The process would fall behind by more than two hours immediately, trigger HTTP 429 errors, and fail to provide current data.

2. Daily Change Volume Analysis:
- Normal day: ~40 changed members across 24 hours (~0.28 changes per 10-minute window).
- Peak race weekend: ~400 changes across a 10-hour peak window (~6.7 changes per 10-minute window).
- Polling 18,000 unchanged records every 10 minutes to capture an average of 1 to 7 changes produces over 99.9% redundant network traffic and server load.

3. Evaluation of Options:
Option 1: Full Polling Every 10 Minutes. Infeasible due to rate limits and a 2.5-hour execution duration.
Option 2: Incremental Polling using modifiedSince. The provider queries every 10 minutes for changes. Up to 7 changed members fit within a single GET /members call, followed by 7 membership calls (8 calls total, requiring under 1 request per minute against the 120 limit).
Option 3: Event-Driven Webhooks (member.created, membership.purchased, membership.expired). Pushes updates instantly when a transaction completes, providing real-time data for race-day registrations without ongoing polling overhead.

4. Recommendation:
We recommend a hybrid model:
1. The provider performs a one-time baseline synchronization of all 18,000 members during off-peak hours (scheduled at 100 requests/minute, taking about 3 hours).
2. Live synchronization is handled via Webhooks, with Option 2 (modifiedSince polling) available as a fallback.
Before proceeding, we need to confirm whether the provider can receive webhook POST requests, whether they need full membership history or just an active status flag, and schedule the baseline sync well ahead of the race."""
    },
    {
        "heading": "Answer Q10 Part A - Reconciliation",
        "instruction": "Up to 250 words. Show each step with the numbers.",
        "is_code": False,
        "content": """Reconciliation of Active Membership Metrics (1 October vs 26 October)

Both figures are correct; they reflect different metrics measured at different points in time. The 1,312 figure represents active membership records on 1 October, while 1,240 represents distinct individual people with active status on 26 October.

Step-by-Step Reconciliation:

1. Active Membership Records on 1 October:
Daniel's export on 1 October contained 1,312 active rows. Exhibit Q10-B confirms that this report outputs one row per membership package, not per person.

2. Multi-Membership Adjustment (1 October Distinct People):
Exhibit Q10-B notes that 58 members held two active memberships simultaneously on 1 October (such as an adult playing membership and an officiating or coaching badge). Each of these 58 individuals accounts for two rows in the export, creating 58 duplicate rows.
Distinct active people on 1 October = 1,312 - 58 = 1,254.
(Assumption: No individual held more than two active memberships on 1 October, consistent with Exhibit Q10-B).

3. Membership Turnover (1 October to 26 October):
Between 1 October and 26 October:
- Memberships expired: 31
- New memberships purchased: 17
Net change in active members = +17 - 31 = -14.

4. Reconciled Active People on 26 October:
Starting from distinct active people on 1 October (1,254) and applying net turnover (-14):
1,254 - 14 = 1,240 distinct active individuals.
This matches today's dashboard figure of 1,240 exactly."""
    },
    {
        "heading": "Answer Q10 Part B - Reply to Daniel",
        "instruction": "Up to 200 words.",
        "is_code": False,
        "content": """Dear Daniel,

Thank you for reaching out. Both figures are correct, but they reflect different measurements:

1,312 was the total number of active membership subscriptions on 1 October. Because 58 of your members hold two memberships simultaneously (such as playing and coaching), they appear twice on that report. The actual number of unique active individuals on 1 October was 1,254 (1,312 minus 58).

1,240 is the number of unique active individuals today (26 October). Between 1 October and today, 31 memberships expired and 17 new ones were purchased, representing a net change of -14 members (1,254 minus 14 = 1,240).

Which figure to use for your funding return:
Sports funding bodies almost universally require the count of unique, individual participants rather than double-counting individuals who hold multiple roles.

If the funding council requires figures as of 1 October, you should submit 1,254.
If they require figures as of your current submission date, you should submit 1,240.

You should not submit 1,312, as submitting duplicated member rows can trigger funding audit issues. Please check the council's guidance to confirm their required census date.

Best regards,
JustGo Support Team"""
    }
]

# ---------------------------------------------------------------------------
# Build Document According to User Requirements
# ---------------------------------------------------------------------------

def build_reformatted_docx():
    doc = Document()

    # Standard 1 inch margins
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
        # Clear any headers and footers
        s.header.is_linked_to_previous = False
        s.header.paragraphs[0].text = ""
        s.footer.is_linked_to_previous = False
        s.footer.paragraphs[0].text = ""

    # Normal Style: Times New Roman, 12pt, Black
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # Document Section Title: Answer Section
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    run_title = p_title.add_run("Answer Section")
    run_title.font.name = "Times New Roman"
    run_title.font.size = Pt(14)
    run_title.bold = True
    run_title.font.color.rgb = RGBColor(0, 0, 0)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    run_sub = p_sub.add_run("Type each answer in the box under its heading. Do not delete, rename or move the headings. Cite the exhibit row, line or figure you rely on and explain your reasoning.")
    run_sub.font.name = "Times New Roman"
    run_sub.font.size = Pt(11)
    run_sub.italic = True
    run_sub.font.color.rgb = RGBColor(0, 0, 0)

    for item in answers:
        # Question Heading: 14pt, Bold, Times New Roman, Black
        p_head = doc.add_paragraph()
        p_head.paragraph_format.space_before = Pt(12)
        p_head.paragraph_format.space_after = Pt(2)
        p_head.paragraph_format.keep_with_next = True
        run_head = p_head.add_run(item["heading"])
        run_head.font.name = "Times New Roman"
        run_head.font.size = Pt(14)
        run_head.bold = True
        run_head.font.color.rgb = RGBColor(0, 0, 0)

        # Question Instruction: 11pt, Italic, Times New Roman, Black
        p_inst = doc.add_paragraph()
        p_inst.paragraph_format.space_before = Pt(0)
        p_inst.paragraph_format.space_after = Pt(4)
        p_inst.paragraph_format.keep_with_next = True
        run_inst = p_inst.add_run(item["instruction"])
        run_inst.font.name = "Times New Roman"
        run_inst.font.size = Pt(11)
        run_inst.italic = True
        run_inst.font.color.rgb = RGBColor(0, 0, 0)

        # Box table (1x1 table)
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)

        tcPr = cell._tc.get_or_add_tcPr()
        
        # Simple solid black border (0.5 pt = 4 in dxa/eighths of a pt)
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)

        # White background (pure white, no shading)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF"/>')
        tcPr.append(shd)

        # Inner padding: top/bottom 120 dxa (6pt), left/right 150 dxa (7.5pt)
        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="120" w:type="dxa"/>'
            f'<w:bottom w:w="120" w:type="dxa"/>'
            f'<w:left w:w="150" w:type="dxa"/>'
            f'<w:right w:w="150" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(tcMar)

        # Populate cell content
        cp = cell.paragraphs[0]
        content_lines = item["content"].strip().split('\n')
        for idx, line in enumerate(content_lines):
            if idx > 0:
                cp = cell.add_paragraph()
            
            if item["is_code"]:
                cp.paragraph_format.space_before = Pt(0)
                cp.paragraph_format.space_after = Pt(0)
                cp.paragraph_format.line_spacing = 1.05
                r = cp.add_run(line)
                r.font.name = "Consolas"
                r.font.size = Pt(10)
                r.font.color.rgb = RGBColor(0, 0, 0)
            else:
                cp.paragraph_format.space_before = Pt(0)
                cp.paragraph_format.space_after = Pt(4)
                cp.paragraph_format.line_spacing = 1.15
                r = cp.add_run(line)
                r.font.name = "Times New Roman"
                r.font.size = Pt(12)
                r.font.color.rgb = RGBColor(0, 0, 0)

        # Spacing after each answer box
        p_spacer = doc.add_paragraph()
        p_spacer.paragraph_format.space_before = Pt(4)
        p_spacer.paragraph_format.space_after = Pt(8)

    output_filename = "Karib Shams - shams321karib@gmail.com - DSE Assessment.docx"
    output_filepath = os.path.join(r"d:\CSE599_THESIS\Data_Solution", output_filename)
    doc.save(output_filepath)
    print(f"Reformatted document successfully created at: {output_filepath}")

if __name__ == "__main__":
    build_reformatted_docx()
