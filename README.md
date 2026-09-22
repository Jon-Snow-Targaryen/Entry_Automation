# Data Entry & Business Automation Portfolio
Welcome to my automation repository. As a tech-savvy Data Specialist and Computer Science graduate, I combine an advanced typing performance of **71 WPM (96% Accuracy)** with programming logic to eliminate repetitive workflows, optimize spreadsheets, and maintain absolute database integrity.

Here are the functional automation scripts I developed to streamline daily administrative and operational tasks:

---

## 🛠️ Project 1: Bulk File Management & Dynamic Renaming Tool (C#)
**Business Application:** Automatically copies, archives, and standardizes names for hundreds of supplier receipts, files, or customer contracts in seconds, eliminating manual data handling and extension errors.

```csharp
// Core logic for dynamic bulk file synchronization and sorting
private void CopyAndRenameFile(string source, string destFolder, string newFileName)
{
    if (!File.Exists(source))
        throw new FileNotFoundException("The original source file does not exist.");

    if (!Directory.Exists(destFolder))
        Directory.CreateDirectory(destFolder);

    string extension = Path.GetExtension(source);
    
    // Safely prevents double extension bugs (e.g., invoice.txt.txt)
    if (newFileName.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
    {
        newFileName = newFileName.Substring(0, newFileName.Length - extension.Length);
    }

    string finalDestinationPath = Path.Combine(destFolder, newFileName + extension);
    File.Copy(source, finalDestinationPath, overwrite: true);
}
```

---

## 🛠️ Project 2: Database Data Scrubbing & Mathematical Rounding Script (C# / SQL)
**Business Application:** Connects to an internal central reader, filters blank or null records securely, and automatically formats numeric floating values into uniform whole numbers to guarantee clean weekly financial reports.

```csharp
// Programmatic data reconciliation and validation loop
if (reader == DBNull.Value || string.IsNullOrWhiteSpace(reader.ToString()))
{
    smu = "0"; // Assigns default fallback placeholders to prevent core system crashes
}
else if (double.TryParse(reader.ToString(), out double numericValue))
{
    // Auto-rounds numbers to standard math configurations
    double roundedValue = Math.Round(numericValue, 0, MidpointRounding.AwayFromZero);
    smu = roundedValue.ToString("F0"); // Trims trailing zeros for neat document presentation
}
else
{
    smu = reader.ToString(); // Safely retains string data if non-numeric
}
```

---

## 🛠️ Project 3: Automated Corporate Records Synchronization (Oracle PL/SQL Stored Procedure)
**Business Application:** A high-performance database script designed to automate business reviews, audit structural missing links, and handle daily transactional uploads seamlessly without slowing down user dashboards.

```sql
CREATE OR REPLACE PROCEDURE Sync_Corporate_Records (
    p_batch_id       IN  NUMBER,
    p_status_out     OUT VARCHAR2
) AS
    v_error_count    NUMBER := 0;
BEGIN
    -- 1. Automatic auditing and elimination of structural missing links
    UPDATE Employee_Records_Staging
    SET record_status = 'INVALID'
    WHERE entry_date IS NULL OR employee_id IS NULL;

    -- 2. Performance-optimized data migration and structural cleansing
    MERGE INTO Corporate_Master_Database target
    USING (
        SELECT employee_id, first_name, last_name, NVL(salary, 0) as cleaned_salary
        FROM Employee_Records_Staging
        WHERE record_status = 'VALID' AND batch_id = p_batch_id
    ) source
    ON (target.employee_id = source.employee_id)
    WHEN MATCHED THEN
        UPDATE SET target.last_update = SYSDATE,
                   target.salary = source.cleaned_salary
    WHEN NOT MATCHED THEN
        INSERT (employee_id, first_name, last_name, salary, created_at)
        VALUES (source.employee_id, source.first_name, source.last_name, source.cleaned_salary, SYSDATE);

    p_status_out := 'SUCCESS: Bulk data synchronization completed.';
    
EXCEPTION
    WHEN OTHERS THEN
        p_status_out := 'FAILED: System roll-back triggered due to execution error.';
        ROLLBACK;
END Sync_Corporate_Records;
/
```

## 📊 Core Technical Expertise
* **Typing Benchmarks:** 71 Words Per Minute | 96% Accuracy (Verified by 10FastFingers)
* **Data Engineering:** Query optimization, advanced indexing, database change management (Oracle PL/SQL, MS SQL Server foundations)
* **Administrative Platforms:** Google Workspace, Advanced MS Excel (XLOOKUP, PivotTables, formatting structures), SharePoint administration
