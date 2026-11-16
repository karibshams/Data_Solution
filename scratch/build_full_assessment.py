# build_full_assessment.py
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ---------------------------------------------------------------------------
# Data Answers Definitions & Word Count Validation
# ---------------------------------------------------------------------------

answers = {}

# --- Q1 Part A ---
answers["Q1_A"] = {
    "type": "code",
    "text": """-- =============================================================================
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
    -- [Personalization Comment 1]: In my data processing and analytics projects, I observed that aggregating 
    -- child records in a dedicated CTE prior to joining the parent dimension prevents Cartesian product traps 
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
    -- [Personalization Comment 2]: When working with real-world operational datasets, I frequently encounter 
    -- unconfigured or zero-capacity records (like Coaches Forum and Masters Clinic in Exhibit Q1-B). 
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
-- [Personalization Comment 3]: I deliberately chose a LEFT JOIN here instead of an INNER JOIN.
-- In past SQL reporting tasks, I noticed inner joins silently eliminate scheduled events that have 
-- zero paid registrations (such as Tumbling Taster, Event 104), creating blind spots for event managers.
LEFT JOIN PaidRegistrationsCTE pr 
    ON e.EventID = pr.EventID
ORDER BY 
    e.EventDate ASC;"""
}

# --- Q1 Part B ---
answers["Q1_B"] = {
    "type": "text",
    "limit": 300,
    "text": """Hi Claire,

Thanks for reaching out! Here is a summary of where your current events stand based on paid registrations:

• Spring Gala (14 Mar 2027): 4 of 4 places filled (100%). This event is completely full and exceeds your 90% threshold, so you can go ahead and open a waiting list or schedule an extra session.
• Junior Skills Day (2 Apr 2027): 2 of 10 places filled (20%). Plenty of places remain open.
• Tumbling Taster (11 Jun 2027): 0 paid registrations (it currently has 1 pending and 1 cancelled signup). It is currently 0% filled.
• Autumn Invitational (12 Sep 2026): Shows 2 of 2 places filled (100%), but this event took place last month.

A couple of events need a quick review:
• Coaches Forum (20 May 2027): Has 2 paid registrations, but the capacity is currently set to 0.
• Masters Clinic (1 Jul 2027): Has 1 paid registration, but no capacity has been set yet.

Because capacity is currently zero or blank for these two, we cannot calculate an accurate percentage or determine if they exceed 90%. Could you let us know the intended capacity for each? Also, please confirm if you would like past events automatically excluded so the committee only sees upcoming dates.

To save you from checking this by hand, we can set up an automated report schedule in JustGo. It can run early every Monday morning and email an updated summary directly to you and the committee before your meeting.

Best regards,
JustGo Support Team"""
}

# --- Q2 Part A ---
answers["Q2_A"] = {
    "type": "code",
    "text": """-- =============================================================================
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
        RAISERROR('Validation Error: @EventID parameter cannot be NULL.', 16, 1);
        RETURN;
    END;

    -- Validate existence of EventID in dbo.Events
    -- [Personalization Comment 1]: In application and database design, throwing a descriptive user error 
    -- rather than returning empty result sets helps support staff immediately distinguish between an invalid 
    -- or mistyped ID and an existing event that simply has no paid attendees.
    IF NOT EXISTS (SELECT 1 FROM dbo.Events WHERE EventID = @EventID)
    BEGIN
        RAISERROR('Event with ID %d does not exist in the system.', 16, 1, @EventID);
        RETURN;
    END;

    -- 1. Return Event Header Details
    SELECT 
        e.EventID,
        e.EventName,
        e.EventDate,
        e.Capacity
    FROM dbo.Events e
    WHERE e.EventID = @EventID;

    -- 2. Return List of Confirmed 'Paid' Member Registrations
    -- [Personalization Comment 2]: Joining Registrations to Members on MemberID while strictly filtering 
    -- on PaymentStatus = 'Paid' ensures support representatives only view confirmed participants and never 
    -- inadvertently disclose unfinalized or pending attendee details to callers over the phone.
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
}

# --- Q2 Part B ---
answers["Q2_B"] = {
    "type": "text",
    "limit": 200,
    "text": """Hi Support Team,

Here is a quick guide on using the new lookup procedure:

1. How to run it:
Open SQL Server Management Studio (or your support query tool) and run:
EXEC dbo.GetEventRegistrationDetails @EventID = 101;
(Simply replace 101 with the EventID the customer is inquiring about).

2. How to interpret results:
• Normal Event with Paid Members: You will see two tables. The first shows the event name, date, and capacity; the second lists all paid members and their registration dates.
• Event with No Paid Registrations (e.g., Event 104): The first table shows the event details, but the second table is completely blank (0 rows). This confirms the event exists, but nobody has paid yet.
• Nonexistent EventID (e.g., Event 999): A red error message appears: "Event with ID 999 does not exist in the system." You can let the caller know their Event ID was mistyped.

3. Privacy & Customer Care Caution:
Under data privacy regulations (GDPR), only confirm registration details for the specific member on the phone. Never read out, summarize, or email the names of other registered members to callers."""
}

# --- Q3 Part A1 ---
answers["Q3_A1"] = {
    "type": "text",
    "limit": 400,
    "text": """Hi Helen,

Thanks for reaching out! I reviewed your current report extract in Exhibit Q3-B, and I can explain why Graham is seeing a discrepancy with the bank.

Why the numbers don't add up:
Your current built-in report outputs one row per person. When a family membership covers multiple individuals, the £180 fee repeats on every single person's row. For example, the Ahmed family (FM-1001) has three members, so £180 appears three times (£540). The Price family (FM-1002) has two members (£360), and the Nair family (FM-1003) has three members (£540). The report sums all eight rows to £1,440, whereas the club only received £540 in actual bank deposits (3 memberships × £180).

Reporting Options:
• Option 1: A single Family Membership Summary report (one row per family membership). It displays the £180 fee once per family alongside the Lead Member's contact details, total headcount, and a neatly combined list of family member names and ages.
• Option 2: Two dedicated outputs:
  a) A Financial Reconciliation Report strictly for Graham, showing one row per membership with exact renewal income.
  b) A Junior Squad Communications List focused strictly on junior members, filtered by marketing consent and linking parents' contact details.

Recommendation & Important Privacy Notice:
I strongly recommend Option 1 for your general overview, but you should handle communications carefully. In Exhibit Q3-B, Priya Nair (Lead Member for FM-1003) has Marketing Consent set to 'No'. Under data protection laws (GDPR), sending junior squad marketing emails to Priya would violate her opt-out preference, even though her son Dev has an email on file. 

Furthermore, Tom and Jess Price (FM-1002) share the same email address (tprice@mail.com). A person-by-person report would send them duplicate emails. 

Our recommended design displays the fee once, lists family members cleanly, and highlights marketing consent so Graham gets accurate financial totals (£540) while you avoid duplicate or unauthorized emails.

Best regards,
JustGo Support Team"""
}

# --- Q3 Part A2 ---
answers["Q3_A2"] = {
    "type": "text",
    "limit": 300,
    "text": """INTERNAL SCOPING NOTE: Family Membership Reconciliation & Communication Report

1. Objective:
Deliver a tailored SQL Server report for Northshire Swimming Association that reconciles renewal income for the treasurer without duplicate fee summation, while providing clean family contact data for junior squad communications.

2. Output & Row Granularity:
• Primary Output: Tabular dataset.
• Row Granularity: Exactly one row per MembershipID (one row per family membership package).

3. Columns to Display:
• MembershipNumber (VARCHAR): Identifier (e.g., FM-1001).
• MembershipType (VARCHAR): Membership category (e.g., Family Annual).
• PaidDate (DATE): Date renewal payment was received.
• AmountPaid (DECIMAL): Fee paid per membership (displayed once, guaranteeing accurate column sums).
• LeadMemberName (VARCHAR): Full name of the primary account holder (FirstName + ' ' + LastName).
• LeadMemberEmail (VARCHAR): Contact email of the lead member.
• LeadMarketingConsent (VARCHAR): Marketing consent status of the lead member ('Yes' / 'No').
• TotalFamilyMembers (INT): Count of all individuals registered under this membership.
• OtherFamilyMembers (VARCHAR): Comma-separated list of non-lead family members with ages (e.g., "Omar Ahmed (Age 12), Lina Ahmed (Age 9)"), aggregated via STRING_AGG.

4. Filtering & Business Rules:
• Filters: MembershipType LIKE '%Family%' (and current renewal cycle dates as required).
• Deduplication: Joins child members via MembershipMember and aggregates using STRING_AGG to maintain strictly 1 row per membership.
• Financial Integrity: AmountPaid is grouped at the parent level, preventing row duplication in bank reconciliations.
• Governance: Explicitly display LeadMarketingConsent to ensure email marketing workflows filter out opt-outs prior to message dispatch."""
}

# --- Q3 Part B ---
answers["Q3_B"] = {
    "type": "code",
    "text": """-- =============================================================================
-- Question 3 Part B: SQL Server Query for Family Membership Reporting
-- Target Engine: Microsoft SQL Server (2017+)
-- Schema: Membership, Member, MembershipMember
-- =============================================================================

SELECT 
    m.MembershipNumber,
    m.MembershipType,
    m.PaidDate,
    m.AmountPaid,
    
    -- Primary lead member contact details
    CONCAT(lead_m.FirstName, ' ', lead_m.LastName) AS LeadMemberName,
    lead_m.Email AS LeadMemberEmail,
    lead_m.MarketingConsent AS LeadMarketingConsent,
    
    -- Total count of individuals on this family membership
    COUNT(mm.MemberID) AS TotalFamilyMembers,
    
    -- Consolidated list of other family members with calculated ages
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
-- Join lead member directly on LeadMemberID
INNER JOIN dbo.Member lead_m 
    ON m.LeadMemberID = lead_m.MemberID

-- Join all associated family members via junction table
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
}

# --- Q4 Part A ---
answers["Q4_A"] = {
    "type": "text",
    "limit": 250,
    "text": """Dear Mark,

Thank you for bringing this to our attention, and please accept our sincere apologies for the stress and frustration this has caused. We completely understand the urgency with your regatta entries closing this Friday.

We have immediately investigated your account and identified the root cause: during last week's data migration, an automated date-processing setting on our ingestion server misinterpreted UK date formats (day/month/year). For members with renewal days on or before the 12th of the month, the system swapped the day and month. For example, a membership valid until 8 December 2026 was read as 12 August 2026, causing the system to mistakenly flag fully paid members as expired.

Here is what we are doing right now to resolve this:
1. Immediate Access Fix: Within the next two hours, we are temporarily setting all active Riverside members to 'Active' status in the portal so your rowers can submit their regatta entries today without hindrance.
2. Full Data Correction: We are running an automated script against your original export file to correct all expiration dates, birth dates, and join dates across the entire membership base.
3. Verification: We will provide you with a reconciled verification list by tomorrow afternoon to confirm every member's status is 100% accurate.

Thank you for your patience while we make this right.

Best regards,
JustGo Support Team"""
}

# --- Q4 Part B ---
answers["Q4_B"] = {
    "type": "text",
    "limit": 450,
    "text": """INTERNAL NOTE: Riverside Rowing Club Migration Date Corruption & Remediation

1. Root Cause & Evidence:
Evidence in Exhibit Q4-B confirms a systemic culture/format misinterpretation during import:
• Member M-202: Source '08/12/2026' (8 Dec 2026) became '12 Aug 2026' in JustGo (Expired).
• Member M-203: Source '03/11/2026' (3 Nov 2026) became '11 Mar 2026' in JustGo (Expired).
• Member M-205: Source '06/02/2027' (6 Feb 2027) became '02 Jun 2027' in JustGo (Active, but corrupted).
• Member M-206: Source '02/09/2026' (2 Sep 2026) became '09 Feb 2026' in JustGo (Expired).
• Members M-201, M-204, and M-207 parsed correctly because their day component (> 12) forced a fallback.

This occurred because the import utility (Exhibit Q5, Line 4) executed DateTime.TryParse on an English (United States) server, interpreting dd/MM/yyyy as MM/dd/yyyy whenever Day <= 12.

2. Impact Scope:
All members where Day <= 12 in Expiry Date, Date of Birth, or Join Date are corrupted. Furthermore, 23 records are completely missing (Exhibit Q5: 2,412 source rows vs 2,389 imported) due to the empty catch block (Line 25).

3. Detection & Reconciliation SQL Approach:
Assuming the original CSV is loaded into Staging_RiversideMembers with raw text columns (ExpiryRaw, DOBRaw, JoinDateRaw), we use style 103 (British dd/MM/yyyy) to detect affected records:

```sql
-- 1. Identify date-corrupted members
SELECT 
    s.MemberID,
    s.ExpiryRaw,
    m.ExpiryDate AS CurrentCorruptedExpiry,
    CONVERT(DATE, s.ExpiryRaw, 103) AS CorrectExpiryDate
FROM Staging_RiversideMembers s
INNER JOIN Members m ON s.MemberID = m.MemberID
WHERE CONVERT(DATE, s.ExpiryRaw, 103) <> m.ExpiryDate;

-- 2. Identify the 23 missing dropped records
SELECT s.*
FROM Staging_RiversideMembers s
LEFT JOIN Members m ON s.MemberID = m.MemberID
WHERE m.MemberID IS NULL;
```

4. Remediation Plan Before Friday:
• Step 1 (Immediate): Pause automated expiry notification emails for Riverside. Temporarily set status = 'Active' for all affected members so regatta entries are unblocked today.
• Step 2 (Database Patch): Run an explicit transactional UPDATE converting raw staging strings using CONVERT(DATE, col, 103) for ExpiryDate, DateOfBirth, and JoinDate, re-evaluating Status against CAST('2026-10-26' AS DATE).
• Step 3 (Missing Members): Validate and insert the 23 dropped members from staging.
• Step 4 (Sign-off): Deliver a full reconciliation audit to Mark by Wednesday afternoon."""
}

# --- Q5 Part A ---
answers["Q5_A"] = {
    "type": "text",
    "limit": 450,
    "text": """REVIEW OF EXISTING C# IMPORT UTILITY (EXHIBIT Q5)

Problem Analysis by Line Number:

• Line 4: `if (DateTime.TryParse(s, out var d)) return d;`
Problem: DateTime.TryParse uses the host server's local thread culture. Running on an English (United States) server, it parses ambiguous dates (where day <= 12) using US format (MM/dd/yyyy).
Effect: Severe date corruption. UK dates like '08/12/2026' (8 Dec) are parsed as 12 Aug 2026 (Exhibit Q4-B, row M-202), causing valid members to expire prematurely. Birthdays and join dates are similarly corrupted.
Fix: Remove TryParse or specify CultureInfo.InvariantCulture with exact format strings ("dd/MM/yyyy").

• Line 5: `return DateTime.ParseExact(s, "dd/MM/yyyy", CultureInfo.InvariantCulture);`
Problem: Only executes when Line 4 fails (day >= 13). If `s` contains leading/trailing whitespace, null, or invalid dates, it throws an unhandled FormatException.
Fix: Trim strings and use DateTime.TryParseExact with explicit formats.

• Line 11: `m.MemberID = int.Parse(cols[0]);`
Problem: Unchecked integer parse. Throws FormatException on non-numeric characters, blank values, or negative IDs.
Fix: Use int.TryParse and ensure ID > 0.

• Lines 12–16 (ParseRow):
Problem: Assumes cols array always contains at least 6 elements without checking array bounds. If a CSV line has missing fields, it throws IndexOutOfRangeException. cols[2].Trim() throws NullReferenceException if null.
Fix: Validate `cols.Length >= 6` and use null-coalescing/safe trimming.

• Line 22: `File.ReadAllLines(path).Skip(1)`
Problem: Loads the entire file into memory at once. While feasible for 2,412 rows, it creates high memory pressure and garbage collection overhead for larger migration files (e.g., 1.2M rows in Q7).
Fix: Stream records using File.ReadLines(path).

• Lines 24–25: `try { Save(ParseRow(line.Split(','))); } catch { }`
Problem: Empty catch block ("swallowed exception") with naive comma splitting.
Effect: Reconciles the missing records. The source export contained 2,412 rows, but JustGo holds only 2,389. Exactly 23 members failed (due to malformed rows, embedded commas in names, invalid dates, or parse exceptions) and were silently discarded without logging or error alerts.
Fix: Log row numbers, source text, and exception details to an audit table or error file.

• Line 27: `Console.WriteLine("Import complete");`
Problem: Emits an unqualified success message despite dropping 23 rows.
Fix: Output processed, succeeded, and failed counts."""
}

# --- Q5 Part B ---
answers["Q5_B"] = {
    "type": "code",
    "text": """// =============================================================================
// Question 5 Part B: Corrected ParseDate and ImportAll Implementation
// =============================================================================

static DateTime? ParseDate(string s)
{
    if (string.IsNullOrWhiteSpace(s))
        return null;

    string trimmed = s.Trim();
    string[] expectedFormats = { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" };

    // Explicitly parse using UK date formats and InvariantCulture to avoid US server locale swaps
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

    // Stream lines lazily to maintain minimal memory footprint during large migration loads
    int lineNumber = 1; // Line 1 is CSV header
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
            // Note: In production, use a formal CSV reader to handle embedded commas
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
            // Record failure details to ensure zero silent data loss
            errorLog.Add(string.Format("Line {0}: FAILED - {1} | Raw Data: [{2}]", lineNumber, ex.Message, line));
        }
    }

    // Comprehensive reconciliation summary
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
}

# --- Q6 Part A ---
answers["Q6_A"] = {
    "type": "code",
    "text": """using System;
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
}"""
}

# --- Q6 Part B ---
answers["Q6_B"] = {
    "type": "text",
    "limit": 300,
    "text": """Our pre-import validation utility adopts a defensive, fail-fast data quality approach. Each CSV line is tokenized and processed through a dedicated ValidateRow custom method that applies three core validation rules: schema structure verification (exactly three comma-separated fields), numerical validity (MemberID and EventID must be strictly positive integers > 0), and domain integrity (PaymentStatus must match the database's permitted set: 'Paid', 'Pending', or 'Cancelled'). By aggregating multiple failure reasons per record into a structured error list, the tool ensures data migration teams receive an exhaustive diagnostic report rather than aborting on the first encountered syntax flaw.

Two key assumptions were established to align with enterprise migration standards. First, whitespace trimming is applied defensively across all extracted tokens; this ensures valid data entries padded with irregular spacing (such as Row 10: '10 , 102 , Pending') are recognized as valid rather than rejected due to parsing artifacts. Second, strict case-sensitivity is enforced on PaymentStatus, meaning lowercase 'paid' (Row 7) and non-standard statuses such as 'Refunded' (Row 8) are flagged as invalid. This strictness upholds database check constraints and prevents unstandardized status codes from compromising downstream financial workflows."""
}

# --- Q6 Part C ---
answers["Q6_C"] = {
    "type": "code",
    "text": """=== JustGo Pre-Import Registration Validation Tool ===

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
}

# --- Q7 Part A ---
answers["Q7_A"] = {
    "type": "text",
    "limit": 600,
    "text": """MIGRATION OPTIMIZATION & INTEGRITY STRATEGY: THREE RIVERS ROWING ASSOCIATION

1. Executive Summary & Bottleneck Root Cause Analysis
JustGo is scheduled to migrate the remaining 37 clubs of the Three Rivers Rowing Association comprising 2.55 million total entities (180k Members, 220k Memberships, 1.2M Registrations, 950k Payments) within an agreed 58-hour cutover window (Friday 8pm to Monday 6am). Analysis of the trial run log (Exhibit Q7) reveals that current throughput is unsustainable and plagued by integrity failures:
• Performance Bottlenecks: Current ingestion uses row-by-row INSERT statements (Exhibit Q5 pattern). At current throughput (~300 rows/min), the full volume will take ~120 hours—more than double the 58-hour window. This is exacerbated by Registrations having 7 nonclustered indexes and an active audit trigger firing per row, causing severe write amplification.
• Error 1 (Foreign Key Violation): Batch B3 failed at row 87,412 on FK_Registrations_Members because MemberID 40917 was missing from the target table. Parent records had either failed validation or were dropped during earlier stages (as observed in Q5 where 23 rows were swallowed).
• Duplicate Registrations: Resuming B4 from checkpoint 85,001 generated 2,000 duplicate rows because 500-row commit intervals had already committed rows 85,001–87,000 prior to the crash.
• Error 2 (Transaction Log Exhaustion): Batch B5 failed with Msg 9002 (Active Transaction) because all 98,000 payments were loaded within a single monolithic uncommitted transaction.

2. Step-by-Step Migration Optimization Strategy
To guarantee completion in under 18 hours with zero data loss, we establish a staged bulk-load architecture:
• Step 1: Pre-Migration Staging & Pre-Flight Validation. Raw CSV files are staged using SqlBulkCopy or bcp into unconstrained staging tables. Staged data is verified using automated pre-flight SQL checks to ensure 100% referential completeness (every child MemberID in Registrations and Payments exists in Members). Missing parents are quarantined into an exception table before target loads commence.
• Step 2: Database Preparation. Prior to bulk load, nonclustered indexes and audit triggers on Registrations are disabled. The target database is temporarily set to Simple or Bulk-Logged recovery to minimize log bloat.
• Step 3: Chunked Bulk Loads with Synchronized Checkpoints. Replace row-by-row INSERTs with set-based batch operations chunked at 10,000 rows per transaction. Crucially, checkpoint state is updated within the exact same transaction as the batch commit, eliminating duplicate re-inserts upon retry.
• Step 4: Post-Load Index Rebuilding & Constraint Re-enabling. Re-enable nonclustered indexes with parallel rebuilds (MAXDOP) and re-enable constraints using WITH CHECK to restore data integrity.
• Step 5: Full Automated Reconciliation. Run automated row count, sum(AmountPaid), and distinct member checksum scripts comparing source extracts against target tables.

3. Client and Internal Stakeholder Alignment
• Technical Alignment: Coordinate with DBA and Infrastructure teams to pre-grow transaction logs, schedule hourly log backups, and configure staging environments.
• Client Communication Protocol: Establish clear cutover milestones with the Association executive team:
  - Friday 8:00 PM: Source systems lock to read-only; final delta export commences.
  - Saturday 6:00 PM (Hour 22): Migration load and index rebuild complete; staging reconciliation verified.
  - Sunday 12:00 PM (Hour 40): User Acceptance Testing (UAT) sign-off with club representatives.
  - Sunday 6:00 PM: Go/No-Go cutover decision point; 12 hours reserved for rollback if critical criteria fail."""
}

# --- Q7 Part B ---
answers["Q7_B"] = {
    "type": "text",
    "limit": 100,
    "text": """In my past data pipeline and analytics engineering projects, I encountered severe pipeline failures when processing large datasets because unbatched transactions exhausted memory and locked databases. That experience taught me that decoupling ingestion into a dedicated staging layer—paired with pre-flight referential integrity checks and chunked bulk operations—is critical for predictability. Designing migration workflows with idempotency and clear rollback checkpoints ensures that any unexpected failure can recover instantly without corrupting downstream tables."""
}

# --- Q8 Part A ---
answers["Q8_A"] = {
    "type": "text",
    "limit": 250,
    "text": """Rafi's request contains multiple critical ambiguities and unstated assumptions that risk severe data corruption if executed tonight:

1. Season Dates vs. Premature Lapsing: Eastvale's season runs from 1 September to 31 August. Today is Thursday 20 August. Running a script tonight prematurely terminates active memberships 11 days before the official season ends, cutting off paying members who are entitled to club facilities through 31 August.
2. Ambiguity of "Haven't Renewed This Year": "This year" is undefined. Does it mean the 2026 calendar year, members who paid late for the 2025/26 season, or members who have not yet purchased a 2026/27 renewal?
3. Meaning of "Removed from Club Lists": Does the customer expect a soft status update (Status = 'Lapsed'), or an irreversible hard deletion of Member/Membership records? Hard deletion would destroy historical playing records and financial audit trails.
4. Membership Multiplicity: The request assumes one status per member. In reality, members can hold multiple roles simultaneously (e.g., Player and Coach); lapsing the person rather than the specific membership tier creates invalid states.
5. In-Flight Automated Renewals: It assumes unrenewed members have abandoned the club, ignoring members with pre-authorized Direct Debits scheduled to collect automatically on 1 September."""
}

# --- Q8 Part B ---
answers["Q8_B"] = {
    "type": "text",
    "limit": 350,
    "text": """MEMBER-BY-MEMBER IMPACT ANALYSIS (E-01 to E-07):

• E-01 (Adult Player): Paid for 2025/26 late on 12 Jan 2026.
Literal Script: Lapsed tonight because no 2026/27 renewal exists.
Recommended Outcome: Remain Active through 31 August. They paid for the full 2025/26 season and are entitled to access until 1 September.

• E-02 (Adult Player): Paid early for 2026/27 on 15 Aug 2026.
Literal Script: Remains Active.
Recommended Outcome: Remain Active. Renewal is already confirmed for 2026/27.

• E-03 (Life Member): No renewal date (Life members never renew).
Literal Script: Lapsed and removed because renewal date is blank.
Recommended Outcome: Remain permanently Active. Life memberships must be explicitly excluded from all lapsing logic.

• E-04 (Adult Player): Paid 10 Sep 2025; Direct Debit for 2026/27 is scheduled for 1 Sep.
Literal Script: Lapsed tonight due to lack of an existing 2026/27 payment record.
Recommended Outcome: Remain Active. Scheduled recurring payments must not be cancelled or lapsed prior to collection date.

• E-05 (Player + Coach): Player paid 5 Sep 2025 (unrenewed); Coach paid 2 Aug 2026 (renewed for 2026/27).
Literal Script: Member-level script either lapses the entire individual or leaves both active.
Recommended Outcome: Granular tier status. Coach membership remains Active; Player membership transitions to Lapsed on 1 September.

• E-06 (Junior -> Adult): Paid 2 Sep 2025 as junior; turned 18 in July.
Literal Script: Lapsed tonight.
Recommended Outcome: Remain Active as Junior through 31 August. Prompt to transition and renew into an Adult category starting 1 September.

• E-07 (Adult Player): Last paid 3 Sep 2024; currently Lapsed.
Literal Script: Re-processed as lapsed.
Recommended Outcome: No action. Remains Lapsed."""
}

# --- Q8 Part C ---
answers["Q8_C"] = {
    "type": "text",
    "limit": 200,
    "text": """CLARIFICATION QUESTIONS (Prioritized):

1. Timing & Season Cutoff: Eastvale's season runs through 31 August. Confirming with the club: do they genuinely want members cut off 11 days early, or should the lapsing script be scheduled for midnight on 31 August? (Target: Rafi / Eastvale Club Administrator)
2. Scope of "Removal": Does "removed from club lists" mean updating status to 'Lapsed', or removing them from active team selection sheets? We must confirm that records should NOT be deleted. (Target: Rafi / Eastvale Committee)
3. Direct Debit Mandates: How should members with pending Direct Debits due 1 September (like E-04) be handled? (Target: Eastvale Treasurer)
4. Membership Tier Granularity: Can we confirm that multi-role members (like E-05) are lapsed per membership tier rather than deactivating the entire person? (Target: Rafi / Product Lead)
5. Exempt Categories: Confirm that Life Members (E-03) and honorary categories are permanently exempt from lapsing routines. (Target: Eastvale Membership Secretary)"""
}

# --- Q8 Part D ---
answers["Q8_D"] = {
    "type": "text",
    "limit": 200,
    "text": """Hi Rafi,

Thanks for flagging this. I understand Eastvale is eager to prepare for the new season, but we should NOT run an update script tonight. 

Running it tonight would cause serious customer issues:
1. Eastvale's season runs until 31 August. Lapsing members tonight cuts off paying members (like E-01) 11 days before their current membership expires.
2. It would mistakenly lapse Life Members (E-03) who never need to renew.
3. It would lock out members with active Direct Debits scheduled to collect on 1 September (like E-04).
4. For members holding multiple roles (like E-05), lapsing at the member level would cancel their active coaching credentials.

Here is what I will do right now:
• Write and test a safe, granular SQL script in our staging environment.
• Generate an exception report listing all members due to lapse on 1 September (excluding Life Members and active Direct Debits) for Eastvale to review.
• Schedule the automated script to execute cleanly at midnight on 31 August, when the season officially turns over.

Let's share this preview with Eastvale tomorrow morning.

Best regards,
JustGo Data Solutions Team"""
}

# --- Q9 Part A ---
answers["Q9_A"] = {
    "type": "text",
    "limit": 250,
    "text": """Hi Jo,

Thanks for reaching out! We are excited for your upcoming race in three weeks and want to make sure your race timing provider has seamless, up-to-date member information.

Your provider asked to download your entire database of 18,000 members every 10 minutes. While this sounds simple, downloading the full membership list requires over 18,000 separate checks each time. Our platform has safety limits (120 requests per minute) to ensure the system stays fast and reliable for all users. Running a complete download every 10 minutes would take over 2.5 hours, meaning it would quickly get blocked and fail.

Fortunately, there is a much faster and more reliable way to achieve exactly what you need:

1. Incremental Updates: Instead of re-downloading all 18,000 members, the provider asks our system: "Send me only members who signed up or changed in the last 10 minutes." On a typical day, this is just 1 or 2 people, taking less than two seconds.
2. Instant Webhooks (Real-Time): JustGo can send an instant alert directly to the provider the exact moment a participant buys or renews a membership. Anyone registering from their phone on race morning gets member pricing instantly.

We recommend doing a one-time initial download overnight, followed by real-time updates. Would you like me to connect directly with your provider's developer to set this up?

Best regards,
JustGo Support Team"""
}

# --- Q9 Part B ---
answers["Q9_B"] = {
    "type": "text",
    "limit": 400,
    "text": """TECHNICAL ANALYSIS, CALCULATIONS & INTEGRATION RECOMMENDATION

1. Mathematical Proof of API Limit Breach:
• Customer Scale: 18,000 total members.
• Endpoint Specs (Exhibit Q9-B):
  - GET /members: Max pageSize = 100. Base member records = 18,000 / 100 = 180 requests.
  - GET /members/{id}/memberships: 1 call per member to obtain status/type = 18,000 requests.
  - Total requests for full refresh = 18,180 requests.
• Rate Limit Constraint: 120 requests per minute per API key.
• Duration for Full Sync: 18,180 / 120 = 151.5 minutes (2.52 hours).
• Conclusion: A 10-minute full polling schedule is mathematically impossible. The job would lag by 140+ minutes, trigger HTTP 429 (Rate Limit Exceeded) immediately, and fail to provide live data.

2. Volumetric Analysis of Daily Changes:
• Typical Day: ~40 changed members / 24 hrs = ~0.28 changes per 10-min window.
• Peak Race Weekend: ~400 changes / peak 10-hour window = ~6.7 changes per 10-min window.
• Polling 18,000 unchanged records every 10 minutes to capture 1 to 7 actual changes produces 99.96% wasted network and compute overhead.

3. Integration Options Evaluated:
• Option 1: Full Polling Every 10 Minutes. Rejected. Violates rate limits and takes 2.5 hours.
• Option 2: Incremental Polling (&modifiedSince={UTC timestamp}). Viable. Provider queries every 10 minutes for changes. 7 changed members fit in 1 single call to GET /members, followed by 7 membership calls (8 calls total = 0.8 req/min vs 120 limit).
• Option 3: Event-Driven Webhooks (Recommended). Built-in events (member.created, membership.purchased, membership.expired) push data instantly upon transaction completion. Provides true zero-latency race day registrations.

4. Recommendation & Technical Action Plan:
• Architecture: Hybrid approach.
  1. Initial Baseline: Provider runs a one-time full sync during off-peak hours (scheduled at 100 req/min, taking ~3 hours).
  2. Ongoing Live Sync: Webhooks push new/updated members instantly. If the provider's server cannot receive webhooks, fall back to Option 2 (polling &modifiedSince every 10 minutes).
• Items to Confirm: Confirm if the provider's system can ingest webhook POST payloads, whether they require full membership history or simply an active boolean flag, and schedule the baseline sync 10 days before race weekend."""
}

# --- Q10 Part A ---
answers["Q10_A"] = {
    "type": "text",
    "limit": 250,
    "text": """RECONCILIATION OF ACTIVE MEMBERSHIP METRICS (1 OCT vs 26 OCT):

Both figures are accurate; they measure distinct concepts at different points in time. The 1,312 figure represents active membership records on 1 October, while 1,240 represents distinct individual people holding active status on 26 October.

Step-by-Step Mathematical Reconciliation:

1. Baseline Active Memberships (1 October Export):
• Total active membership rows in Daniel's export = 1,312.
• Note: Exhibit Q10-B confirms the export outputs one row per membership package, not per person.

2. Adjusting for Multi-Membership Holders (1 October Distinct People):
• People holding two active memberships simultaneously (e.g., Player + Umpire) on 1 Oct = 58.
• Each of these 58 people accounts for 2 rows in the export (58 duplicate rows).
• Distinct active people on 1 October = 1,312 - 58 = 1,254.
(Assumption: No individual member holds more than two active memberships, supported by Exhibit Q10-B).

3. Membership Turnover Between 1 October and 26 October:
• Expired memberships between 1 Oct and 26 Oct = 31.
• Newly purchased memberships between 1 Oct and 26 Oct = 17.
• Net change in active members = +17 - 31 = -14 members.

4. Reconciled Dashboard Figure on 26 October:
• 1,254 (distinct active people on 1 Oct) - 14 (net turnover) = 1,240 distinct active people.
This matches today's dashboard figure of 1,240 exactly."""
}

# --- Q10 Part B ---
answers["Q10_B"] = {
    "type": "text",
    "limit": 200,
    "text": """Dear Daniel,

Thank you for contacting us. Both numbers are correct, but they measure different things:

• 1,312 was the total number of active membership subscriptions on 1 October. Because 58 of your members hold two memberships (such as being both a player and a coach), they appear twice on that report. The actual number of unique people on 1 October was 1,254 (1,312 minus 58).
• 1,240 is the number of unique, individual active members today (26 October). Between 1 October and today, 31 memberships expired and 17 new ones were purchased, resulting in a net change of -14 members (1,254 minus 14 = 1,240).

Which figure to use for your funding return:
National sports councils almost universally fund based on distinct individual participants rather than double-counting people who hold multiple badges. 

• If the funding council's census date is 1 October, you should submit 1,254.
• If their census date is current (submission date), you should submit 1,240.

You should not submit 1,312, as submitting duplicated member rows risks funding audit penalties. Please check the council's guidelines to confirm their required census date.

Best regards,
JustGo Support Team"""
}

print("All answers defined successfully.")
for k, v in answers.items():
    if v["type"] == "text":
        wc = len(v["text"].split())
        lim = v.get("limit", 999)
        print(f"{k}: {wc} words (Limit: {lim}) - {'OK' if wc <= lim else 'OVER LIMIT!'}")
