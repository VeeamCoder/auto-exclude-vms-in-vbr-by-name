## **Purpose**

This script **automates the management of Global VM Exclusions in Veeam Backup & Replication (VBR)** for Hyper-V VMs, based on their names. It’s designed to:

*   **Exclude VMs** whose names match a specific pattern (e.g., contain “-control-plan” and are longer than 20 characters).
*   **Add new matches** to the exclusion list.
*   **Remove exclusions** for VMs that no longer exist.
*   **Send an email notification** if any changes are made.

***

## **How It Works**

### 1. **Configuration**

*   You set the name pattern (`-control-plan`) and minimum name length (21).
*   SMTP settings are configured for email notifications.
*   Logging options are set.

### 2. **Inventory Collection**

*   The script connects to all Hyper-V sources known to Veeam (hosts, clusters, optionally SCVMM).
*   It checks if each source is reachable (using a TCP probe).
*   Only reachable sources are queried for VM inventory.

### 3. **Candidate Selection**

*   From the discovered VMs, it selects those whose names:
    *   Are at least 21 characters long.
    *   Contain the pattern `-control-plan`.

### 4. **Exclusion Management**

*   **Additions:** If a matching VM is not already excluded, it’s added to the Global VM Exclusions.
*   **Removals:** If a previously excluded VM (matching the pattern and length) no longer exists, it’s removed from the exclusions.
*   **Safety:** Removals are only performed if at least one inventory source was successfully queried (to avoid false positives).

### 5. **Notification**

*   If any exclusions were added or removed, an email is sent to the configured recipients, summarizing the changes and listing any unreachable sources.

### 6. **Logging**

*   Actions and changes are logged to a file for auditing.

***

## **Typical Use Case**

This script is ideal for environments where you have **Azure Arc or Azure Local control-plane VMs** (with long, patterned names) that should **not be backed up** by Veeam, and you want to automate their exclusion from backup jobs.

***

## **Summary Table**

| Step                 | What It Does                                                       |
| -------------------- | ------------------------------------------------------------------ |
| Configuration        | Sets name pattern, length, email, and logging options              |
| Inventory            | Collects VM info from all reachable Hyper-V sources                |
| Candidate Selection  | Finds VMs matching the name pattern and length                     |
| Exclusion Add/Remove | Adds new matches to exclusions, removes old ones no longer present |
| Notification         | Emails a summary if changes were made                              |
| Logging              | Logs actions for auditing                                          |
