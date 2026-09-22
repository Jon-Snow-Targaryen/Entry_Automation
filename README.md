# Automation Programming Logic (C#, PL/SQL, SQL, )
Programming logic to eliminate repetitive workflows, optimize spreadsheets, and maintain absolute database integrity.

Functional automation scripts to streamline daily administrative and operational tasks:

---

## Bulk File Management & Dynamic Renaming Tool (C#)
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


// A Windows Forms desktop app that reads structural customer tables from a raw text file (`~` delimited metadata tracking), dynamically builds financial range summaries, loops through multi-recipient mail domains, and dispatches dynamic emails via SMTP. It completely removes the manual stress of handling billing distributions.

using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;

namespace Enterprise_Billing_Distributor
{
    public partial class MainForm : Form
    {
        string attachmentsDirectory = "system_attachments";

        public MainForm()
        {
            InitializeComponent();
            PopulateSystemFilters();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Directory.CreateDirectory(attachmentsDirectory);
        }

        private void PopulateSystemFilters()
        {
            // Dynamically populate chronological drop-down data matrix options
            for (int year = 2020; year <= 2099; year++) cbYear.Items.Add(year.ToString());
            cbYear.SelectedItem = DateTime.Now.Year.ToString();

            for (int day = 1; day <= 31; day++) { cbStartDay.Items.Add(day.ToString()); cbEndDay.Items.Add(day.ToString()); }
            cbStartDay.Text = "1"; cbEndDay.Text = "1";

            foreach (string month in CultureInfo.CurrentCulture.DateTimeFormat.MonthNames)
            {
                if (!string.IsNullOrEmpty(month)) cbMonth.Items.Add(month);
            }
            cbMonth.SelectedItem = DateTime.Now.ToString("MMMM");
            cbEmailWorkflowType.Text = "Whole Month";
        }

        private void btnProcessDistribution_Click(object sender, EventArgs e)
        {
            string promptMessage = cbEmailWorkflowType.Text == "Ranged Date" 
                ? $"Launch automated billing distribution from {cbMonth.Text} {cbStartDay.Text}-{cbEndDay.Text}, {cbYear.Text}?"
                : $"Launch automated billing distribution for the full month of {cbMonth.Text} {cbYear.Text}?";

            if (MessageBox.Show(promptMessage, "System Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            // Stream line reader processes structural batch targets independently without blocking UI streams
            using (StreamReader dataReader = new StreamReader(tbDistributionListFile.Text))
            {
                string dataLine;
                while ((dataLine = dataReader.ReadLine()) != null)
                {
                    // Data tokenization mapping: ACCT_ID~BATCH_SEG~TOTAL_FILES~INV_COUNT~CM_COUNT~FILE_NAME~RECIPIENT_EMAILS
                    string[] targetDetails = dataLine.Split('~');
                    
                    try
                    {
                        using (MailMessage clientMail = new MailMessage())
                        {
                            using (SmtpClient smtpGateway = new SmtpClient())
                            {
                                smtpGateway.Host = "://secure-mail-server.com";
                                smtpGateway.Port = 587;
                                smtpGateway.EnableSsl = true;
                                smtpGateway.UseDefaultCredentials = false;
                                smtpGateway.Credentials = new NetworkCredential(tbGatewayUser.Text, tbGatewayPassword.Text);

                                clientMail.From = new MailAddress(tbGatewayUser.Text, "Corporate Finance Operations");

                                // Parse comma-separated direct customer contacts safely
                                string[] targetRecipients = targetDetails[6].Split(',');
                                foreach (string emailAddress in targetRecipients)
                                {
                                    clientMail.To.Add(new MailAddress(emailAddress.Trim()));
                                }

                                string temporalRange = cbEmailWorkflowType.Text == "Ranged Date" 
                                    ? $"{cbMonth.Text} {cbStartDay.Text}-{cbEndDay.Text}, {cbYear.Text}"
                                    : $"{cbMonth.Text} {cbYear.Text}";

                                clientMail.Subject = $"Statement Summary Reference ({temporalRange}) - ID: {targetDetails[0]}";
                                clientMail.Body = 
                                    $"Dear Valued Customer,{Environment.NewLine}{Environment.NewLine}" +
                                    $"Please find attached your operational records and service account itemizations for the processing period of {temporalRange}.{Environment.NewLine}{Environment.NewLine}" +
                                    $"- Consolidated Ledger Attachments: {targetDetails[2]}{Environment.NewLine}" +
                                    $"- Total Standard Invoices Processed: {targetDetails[3]}{Environment.NewLine}" +
                                    $"- Total Credit Adjustments Issued: {targetDetails[4]}{Environment.NewLine}{Environment.NewLine}" +
                                    $"For any discrepancies or data inquiries, please reach out directly to your assigned accounts manager.{Environment.NewLine}{Environment.NewLine}" +
                                    $"Best Regards,{Environment.NewLine}Finance Operations Administration{Environment.NewLine}{Environment.NewLine}" +
                                    $"CONFIDENTIALITY NOTICE: This transmission is intended solely for the addressee and contains legally privileged parameters. Unauthorized replication, scrubbing, or file duplication is strictly prohibited.";

                                smtpGateway.Send(clientMail);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Internal error monitoring logging logic handles edge-case faults gracefully
                    }
                }
            }
            MessageBox.Show("Automated batch dispatch cycle completed successfully.", "Status Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}


```

---

## Database Data Scrubbing & Mathematical Rounding Script (C# / SQL)
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

## Automated Corporate Records Synchronization (Oracle PL/SQL Stored Procedure)
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

CREATE OR REPLACE PROCEDURE Update_Expired_Order_Status IS
    -- Ensures auditing logs are saved independently even if main transactions fail
    PRAGMA AUTONOMOUS_TRANSACTION;
    v_log_id       NUMBER;
    v_process_name VARCHAR2(100) := 'UPDATE_EXPIRED_ORDER_STATUS';
BEGIN
    -- 1. Initialize background monitoring and performance audit trail
    v_log_id := System_Logbook.Initialize_Entry(v_process_name, SYSDATE);

    -- 2. Execute automated structural update on business records
    -- Automatically flags pending expired items to a specific operational queue
    UPDATE Corporate_Order_Registry COR 
    SET COR.workflow_status = 'PENDING_REFUND' 
    WHERE COR.workflow_status = 'PENDING_CHARGE' 
      AND COR.clearing_deadline <= TRUNC(SYSDATE); 

    -- 3. Commit changes immediately under the autonomous workflow context
    COMMIT;

    -- 4. Finalize activity tracking and runtime analytics logging
    System_Logbook.Finalize_Entry(v_log_id, SYSDATE);
  
EXCEPTION
    WHEN OTHERS THEN
        -- Secure rollback mechanism to prevent locking active data streams
        ROLLBACK;
        RAISE;
END Update_Expired_Order_Status;
/

CREATE OR REPLACE PROCEDURE Update_Asset_Transfer_Inventory IS
    -- Isolates structural log recording to prevent performance issues
    PRAGMA AUTONOMOUS_TRANSACTION;
    v_log_id       NUMBER;
    v_process_name VARCHAR2(100) := 'UPDATE_ASSET_TRANSFER_INVENTORY';
BEGIN
    -- 1. Initialize audit trail documentation
    v_log_id := System_Logbook.Initialize_Entry(v_process_name, SYSDATE);
  
    -- 2. Performance loop: Process dynamic asset transfers processed on the previous day
    FOR r_item IN (
        SELECT tr.transfer_id, det.quantity, det.asset_value
        FROM System_Asset_Transfers tr
        JOIN System_Asset_Transfer_Details det ON det.transfer_id = tr.transfer_id
        WHERE tr.date_received IS NOT NULL
          AND TRUNC(tr.date_received_update) = TRUNC(SYSDATE) - 1
    ) LOOP
        -- 3. Automatically adjust and balance system inventory configurations based on transfer loads
        UPDATE System_Configuration_Parameters scp 
        SET scp.param_allocation_val_1 = scp.param_allocation_val_1 - r_item.quantity, 
            scp.param_allocation_val_2 = scp.param_allocation_val_2 - r_item.quantity
        WHERE scp.param_code LIKE 'ASSET_TRANSFER_CORE_INV_%'
          AND scp.param_description = r_item.asset_value
          AND scp.param_allocation_val_2 > 0;
          
        -- 4. Dynamic operational warning log for clean tracking of heavy data loads
        System_Logbook.Record_Warning(v_log_id, 'LOG: Asset Value ' || r_item.asset_value || ' - QTY: ' || r_item.quantity);
    END LOOP;

    -- 5. Commit operational changes permanently under the autonomous execution layer
    COMMIT;

    -- 6. Finalize runtime tracking execution
    System_Logbook.Finalize_Entry(v_log_id, SYSDATE);
  
EXCEPTION
    WHEN OTHERS THEN
        ROLLBACK;
        RAISE;
END Update_Asset_Transfer_Inventory;
/

DECLARE
    v_next_log_id NUMBER;
BEGIN
    -- 1. Query records within a specific date boundary for bulk status processing
    FOR r_voucher IN (
        SELECT item_serial_no, operational_branch_code
        FROM Corporate_Voucher_Registry
        WHERE TRUNC(validation_date) >= TO_DATE('12/01/2024', 'MM/DD/YYYY') 
          AND TRUNC(validation_date) <= TO_DATE('12/03/2024', 'MM/DD/YYYY')
    ) LOOP

        -- 2. Dynamically calculate the next structural log ID to prevent unique constraint conflicts
        SELECT MAX(audit_log_no) INTO v_next_log_id
        FROM System_Workflow_Logs 
        WHERE voucher_serial_no = r_voucher.item_serial_no;

        -- Increment the tracking identifier safely
        v_next_log_id := NVL(v_next_log_id, 0) + 1;

        -- 3. Insert standardized operational status updates to maintain audit traceability
        INSERT INTO System_Workflow_Logs (
            audit_log_no, fr_status, to_status, created_date, 
            created_by, voucher_serial_no, approved_by, branch_code
        )
        VALUES (
            v_next_log_id, 'ISSUED', 'PROCESSED', SYSDATE, 
            'AUTOMATION_AGENT', r_voucher.item_serial_no, 'SYSTEM_ADMIN', r_voucher.operational_branch_code
        )
    );
    END LOOP;
    
    -- Commit changes to secure data processing integrity
    COMMIT;
END;
/




//Automation of Report to Excel file

CREATE OR REPLACE PROCEDURE Generate_Infrastructure_Report IS 
    -- 1. Configuration variables for data destination and dynamic naming
    v_filepath      VARCHAR2(200) := 'C:\System_Exports\Reports\';
    v_filename      VARCHAR2(200);
    v_full_file_url VARCHAR2(400);
    v_active_assets NUMBER(3) := 0;
    v_total_records NUMBER := 0;
    v_delimiter     VARCHAR2(1) := '~'; -- Delimiter used to compile flat data structure safely
    
    -- File handle types utilizing standard I/O library models
    v_output_file   TEXT_IO.FILE_TYPE;
    v_text_buffer   VARCHAR2(5000);
BEGIN
    -- 2. Dynamically construct filename with a safe timestamp sequence to avoid duplicate overwrites
    v_filename := 'INFRA_MONITORING_' || TO_CHAR(SYSDATE, 'YYYYMMDD_HH24MISS') || '.xls';
    v_full_file_url := v_filepath || v_filename;
    
    -- Open cloud/local stream directory with Write permissions
    v_output_file := TEXT_IO.FOPEN(v_full_file_url, 'W');
    
    -- 3. Design structural report headers for cross-departmental business reviews
    v_text_buffer := 'Vendor Code'
                ||v_delimiter||'Vendor Name'
                ||v_delimiter||'Billing Month'
                ||v_delimiter||'Asset Type'
                ||v_delimiter||'Reporting Period'
                ||v_delimiter||'Branch Code'
                ||v_delimiter||'Account Assignee'
                ||v_delimiter||'Account Number'
                ||v_delimiter||'Service Identifier'
                ||v_delimiter||'Account Status'
                ||v_delimiter||'Portal Sync Status'
                ||v_delimiter||'Contractual Rate'
                ||v_delimiter||'Rate Breakdown'
                ||v_delimiter||'Operational Manager'
                ||v_delimiter||'Account Creation Date'
                ||v_delimiter||'Effectivity Date'
                ||v_delimiter||'Termination Date'
                ||v_delimiter||'Operational Remarks';
    
    -- Commit header block to flat stream
    TEXT_IO.PUT_LINE(v_output_file, v_text_buffer);
    
    -- 4. Execute complex database cursor joins to collate data from 5 separate operational layers
    FOR r_asset IN (
        SELECT ast.vendor_code, 
               ast.asset_type,                             
               det.branch_code, 
               det.service_identifier, 
               acc.account_no, 
               det.effectivity_date, 
               det.modified_date,
               det.terminated_date, 
               det.workflow_status, 
               det.portal_sync_status, 
               rem.admin_remarks, 
               ast.vendor_name, 
               acc.account_status, 
               det.created_date, 
               det.rate_breakdown, 
               det.contract_rate,
               det.assignee_name, 
               mgr.assigned_manager
        FROM System_Accounts acc
        JOIN System_Account_Details det    ON acc.account_id = det.account_id
        JOIN System_Account_Remarks rem    ON det.account_detail_id = rem.account_detail_id
        JOIN System_Vendor_Registry ast    ON ast.vendor_id = acc.vendor_id
        JOIN System_Assigned_Managers mgr  ON ast.vendor_id = mgr.vendor_id
        WHERE acc.account_no = DECODE(:UI_BLOCK.ACCOUNT_NO, NULL, acc.account_no, :UI_BLOCK.ACCOUNT_NO)
          AND det.service_identifier = DECODE(:UI_BLOCK.SERVICE_NO, NULL, det.service_identifier, :UI_BLOCK.SERVICE_NO)
          AND ast.asset_type = DECODE(:UI_BLOCK.ASSET_TYPE, NULL, ast.asset_type, :UI_BLOCK.ASSET_TYPE)
          AND TRUNC(det.effectivity_date) BETWEEN NVL(:UI_BLOCK.START_DATE, det.effectivity_date) 
                                              AND NVL(:UI_BLOCK.END_DATE, det.effectivity_date)
        ORDER BY det.created_date                  
    ) LOOP
        
        -- 5. Business logic conditional check: Calculate active contract lifecycle volumes
        IF r_asset.workflow_status <> 'TERMINATED' AND (r_asset.terminated_date IS NULL OR r_asset.terminated_date > TRUNC(SYSDATE)) THEN
            v_active_assets := v_active_assets + 1;
        END IF;
        
        -- 6. Concatenate active database elements systematically into flat stream file rows
        v_text_buffer := r_asset.vendor_code
                    ||v_delimiter||r_asset.vendor_name
                    ||v_delimiter||r_asset.effectivity_date
                    ||v_delimiter||r_asset.asset_type 
                    ||v_delimiter||r_asset.modified_date       
                    ||v_delimiter||r_asset.branch_code
                    ||v_delimiter||r_asset.assignee_name
                    ||v_delimiter||r_asset.account_no
                    ||v_delimiter||r_asset.service_identifier 
                    ||v_delimiter||r_asset.account_status
                    ||v_delimiter||r_asset.portal_sync_status
                    ||v_delimiter||r_asset.contract_rate
                    ||v_delimiter||r_asset.rate_breakdown
                    ||v_delimiter||r_asset.assigned_manager
                    ||v_delimiter||r_asset.created_date
                    ||v_delimiter||r_asset.effectivity_date 
                    ||v_delimiter||r_asset.terminated_date
                    ||v_delimiter||r_asset.admin_remarks;
                   
        -- Push clean line item block into data warehouse repository
        TEXT_IO.PUT_LINE(v_output_file, v_text_buffer);                                                                                                                   
        
        v_total_records := v_total_records + 1; 
    END LOOP;
    
    -- 7. Safe close protocol to protect system resources and file memory allocations
    TEXT_IO.FCLOSE(v_output_file);
    
    -- System validation / Guard clause triggers if report returns blank metrics
    IF v_total_records = 0 THEN
        System_Interface.Display_Alert('No result found on the filters entered. Kindly check inputs and try again.', 'I');
    ELSE
        -- Automatically launches secure dynamic spreadsheet link via browser interface for executive review
        System_Interface.Launch_Browser_URL('https://company.com' || v_filename);
    END IF;
    
    -- Pass metadata parameter back to active system block interface
    :UI_BLOCK.FILE_NAME := v_filename;
END; 
/

Business Value: This enterprise-grade automated reporting module extracts infrastructure dataset layers across 5 relational tables, processes structural row tracking parameters, and compiles flat spreadsheet documents dynamically utilizing file stream handling library utilities (TEXT_IO). It automatically handles relational data joins and features an absolute data guard validation rule to optimize remote cloud report distributions.
```

## Core Technical Expertise
* **Data Engineering:** Query optimization, advanced indexing, database change management (Oracle PL/SQL, MS SQL Server foundations)
* **Administrative Platforms:** Google Workspace, Advanced MS Excel (XLOOKUP, PivotTables, formatting structures).
