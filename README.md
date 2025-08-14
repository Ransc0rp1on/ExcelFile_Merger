# ExcelFile Merger
### **Features:**

1. **Supports both CSV and Excel files** (.csv and .xlsx)
2. **Preserves all columns** from both files
3. **Handles multi-line fields** (like Plugin Output)
4. **Maintains data integrity** for special characters
5. **Automatically aligns columns** from both files
6. **Works without external modules** (uses native COM objects)

Run this command in PowerShell:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
```

Then run your script:
```powershell
 .\Merge-NessusReports.ps1 -FIles ".\GB_UPI- Internal - 9 Ips_4w6a0w.csv" ".\GB_UPI- Internal - 9 Ips_irxvnq.csv" -OutputFile "Merged.csv"
```
