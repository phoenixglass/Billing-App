# Audit Log Review Procedure

Regular review of audit logs is required for HIPAA compliance and security monitoring.

## Log Location & Access

**Log files:** `audit_logs/audit_YYYYMMDD.log`

```bash
# View today's logs
tail -f audit_logs/audit_$(date +%Y%m%d).log

# View logs for a specific date
cat audit_logs/audit_20260615.log

# Search for specific events
grep "ERROR\|FAILED\|LOCKED" audit_logs/audit_*.log
```

## What to Look For

### 1. Authentication Issues

**Red Flags:**
- Multiple failed login attempts from same username
- Login attempts during unusual hours
- Failed attempts from multiple usernames in short period (password guessing)
- Account lockout events (>5 failed attempts)

**Example:**
```
2026-06-10 14:22:15 - WARNING - Failed login attempt for user 'attacker' (attempt 1)
2026-06-10 14:22:18 - WARNING - Failed login attempt for user 'attacker' (attempt 2)
2026-06-10 14:22:21 - WARNING - Failed login attempt for user 'attacker' (attempt 3)
2026-06-10 14:22:24 - WARNING - Failed login attempt for user 'attacker' (attempt 4)
2026-06-10 14:22:27 - WARNING - Failed login attempt for user 'attacker' (attempt 5)
2026-06-10 14:22:30 - WARNING - User 'attacker' locked out for 15 minutes
```

**Action:** Investigate source of attempts. If external: firewall block. If internal: security incident.

---

### 2. Session Timeout Events

**Expected:**
```
2026-06-10 15:30:00 - INFO - Session timeout for user 'jasmine'
```

**Red Flag:** Session timeout without corresponding login after (user session abandoned).

**Action:** Normal if user walked away. Unusual if frequent.

---

### 3. File Operations

**What to expect:**
```
2026-06-10 10:30:15 - INFO - User 'jasmine' uploaded file: billing_06102026.xlsx (245678 bytes)
2026-06-10 10:31:25 - INFO - Successfully processed file: billing_06102026.xlsx, generated 6 output files
2026-06-10 10:32:10 - INFO - User 'jasmine' downloaded: billing_06102026_Rosanna.xlsx
```

**Red Flags:**
- Uploads of unusual file sizes (>50MB rejected, but check for pattern)
- Processing of unusual filenames (contains patient names, sensitive data)
- Downloads not matching expected output files
- Same file processed multiple times in short period (accidental re-upload?)

**Action:** Verify uploads were intentional. Check for data leakage (filenames with PHI).

---

### 4. Errors & Exceptions

**Example:**
```
2026-06-10 11:05:42 - ERROR - Failed to process file: billing.xlsx - Invalid format
2026-06-10 11:06:30 - ERROR - Database connection failed
```

**Red Flags:**
- Repeated errors for same file (may indicate attack/malformed input)
- File processing errors after successful upload (data corruption?)
- Connection errors (system compromise? network issues?)

**Action:** Investigate root cause. Check file integrity. Review network status.

---

### 5. Access Patterns

**Questions to ask:**
- ✓ Do uploads/downloads match user roles?
- ✓ Are there downloads at unusual times (3am, weekends)?
- ✓ Is data being accessed from unexpected locations?
- ✓ Are there rapid-fire operations (automated script?)

**Example (suspicious):**
```
2026-06-10 03:15:22 - INFO - User 'cb' logged in successfully
2026-06-10 03:15:45 - INFO - User 'cb' downloaded: report_patient_001.xlsx
2026-06-10 03:15:52 - INFO - User 'cb' downloaded: report_patient_002.xlsx
2026-06-10 03:15:59 - INFO - User 'cb' downloaded: report_patient_003.xlsx
... (100 more downloads in 5 minutes)
```

**Action:** Possible data exfiltration. Investigate immediately.

---

## Weekly Review Checklist

Create a log review checklist and complete weekly:

```markdown
## Audit Log Review - Week of [DATE]

**Reviewer:** ________________  
**Date Reviewed:** ________________

### Authentication
- [ ] Any failed login attempts? _____ (expected: 0-2)
- [ ] Any account lockouts? _____ (expected: 0)
- [ ] Any logins from unusual times? (list):

### Session Activity
- [ ] Normal session timeouts observed? Yes / No
- [ ] Any extended sessions (>2 hours without timeout)? Yes / No

### File Operations
- [ ] File uploads match expected files? Yes / No
- [ ] Any uploads with suspicious filenames? Yes / No
- [ ] Any files >10MB? (list):
- [ ] Download activity reasonable? Yes / No

### Errors
- [ ] Any ERROR entries in logs? (list):
- [ ] Any patterns in errors? Yes / No

### Overall Assessment
- [ ] No suspicious activity detected ✓
- [ ] Minor issues, documented in notes
- [ ] Significant findings, escalate to security team

**Notes:**
[Add any observations, anomalies, or questions]

**Escalation Required?** Yes / No
**If yes, escalated to:** ________________ **Date:** ________________
```

---

## Monthly Summary Report

Aggregate weekly reviews into a monthly report:

```markdown
# Audit Log Summary - June 2026

## Key Metrics

| Metric | Count | Status |
|--------|-------|--------|
| Total Login Attempts | 247 | ✓ Normal |
| Failed Logins | 3 | ✓ Low |
| Account Lockouts | 0 | ✓ None |
| Files Processed | 42 | ✓ Normal |
| Data Downloads | 156 | ✓ Normal |
| Errors Encountered | 2 | ✓ Minor |

## Activity by User

| User | Logins | Uploads | Downloads | Status |
|------|--------|---------|-----------|--------|
| jasmine | 42 | 15 | 42 | ✓ Normal |
| cb | 38 | 14 | 38 | ✓ Normal |

## Incidents & Findings

- **2026-06-05 11:22** - File processing error (invalid format) - Resolved
- **2026-06-12 14:55** - 2 failed login attempts for 'jasmine' - User error (wrong password), user retried successfully

## Recommendations

- [ ] Increase monitoring frequency if incidents rise
- [ ] Review access procedures with team
- [ ] Consider log storage upgrade (approaching capacity)

**Signed:** ________________  
**Date:** ________________
```

---

## Automated Log Monitoring (Optional)

For larger deployments, set up automated alerts:

```bash
#!/bin/bash
# monitor_logs.sh - Alert on suspicious activity

LOG_DIR="audit_logs"
TODAY=$(date +%Y%m%d)
LOG_FILE="$LOG_DIR/audit_$TODAY.log"

# Check for failed logins (last 5 minutes)
failed_logins=$(grep "Failed login" "$LOG_FILE" | tail -50 | wc -l)
if [ "$failed_logins" -gt 10 ]; then
    echo "ALERT: $failed_logins failed login attempts in last 5 minutes" | mail -s "Security Alert" security@org.com
fi

# Check for ERROR entries
errors=$(grep "ERROR" "$LOG_FILE" | tail -50)
if [ ! -z "$errors" ]; then
    echo "ALERT: Errors found in audit log" | mail -s "Security Alert" security@org.com
fi

# Run daily
# Add to crontab: 0 * * * * /path/to/monitor_logs.sh
```

---

## Retention Policy

**Audit Log Retention:**
- [ ] 7 days: On-disk (fast access)
- [ ] 90 days: Archive storage (encrypted)
- [ ] 7 years: Offline backup (per HIPAA requirement)

```bash
# Archive logs older than 90 days
find audit_logs/ -name "audit_*.log" -mtime +90 -exec gzip {} \;
tar czf audit_archive_$(date +%Y%m).tar.gz audit_logs/audit_*.gz
# Move to archival storage (encrypted, offline backup)
```

---

## Questions for Security Officer

When reviewing logs, if you find anything suspicious:

1. **What was the impact?** (Data accessed? Files downloaded? System changed?)
2. **How did it happen?** (Account compromise? Misconfiguration? User error?)
3. **When did it start?** (First incident? Ongoing pattern?)
4. **Is it ongoing?** (Single event? Continued attempts?)
5. **What's the fix?** (Reset password? Update rules? Monitor?)

---

## Contact Information

**Audit log questions:**
- [ ] Security Officer: ________________
- [ ] HIPAA Compliance Officer: ________________
- [ ] Help Desk / IT Support: ________________

**Escalation procedure:**
1. Document finding with date/time/details
2. Contact security officer immediately if severity is HIGH
3. Follow incident response procedures (see INCIDENT_RESPONSE.md)
