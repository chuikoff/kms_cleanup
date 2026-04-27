🧹 KMS Cleanup Script

PowerShell script for removing traces of KMS activators and restoring Windows security settings.

🔧 Features
Audit-first approach (no changes before review)
Interactive remediation
Removes:
Scheduled tasks (KMS-related)
Services
Run/RunOnce entries
Defender exclusions
Restores:
Windows Defender (if disabled)
Resets:
Windows activation (removes KMS)
⚙️ Usage
Dry run (recommended first)
.\kms_cleanup.ps1 -DryRun
Interactive mode
.\kms_cleanup.ps1
Automatic mode
.\kms_cleanup.ps1 -AutoApprove
🛡️ Notes
Script does not modify Defender if it is already running correctly
Handles slmgr errors (like rearm limit reached)
Requires Administrator privileges
⚠️ Limitations
Does not guarantee full malware removal
Advanced persistence (WMI, DLL hijacking) not covered
💡 After running
Reboot system
Activate Windows using a valid license
