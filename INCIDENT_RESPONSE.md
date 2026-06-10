# Incident Response Procedure

This document outlines procedures for responding to security incidents involving PHI.

## Incident Severity Levels

### 🔴 CRITICAL (Respond Immediately - <1 hour)
- Active unauthorized access detected
- PHI suspected to be exfiltrated or accessed
- System compromise (malware, backdoor)
- Ransomware attack
- Large-scale data breach

**Action:** Isolate system immediately. Contact security & legal.

---

### 🟠 HIGH (Respond Urgently - <4 hours)
- Multiple failed login attempts (password guessing attack)
- Suspicious data download patterns
- Unauthorized user account activity
- System integrity issues (file corruption, unexpected errors)
- Intrusion detection alerts

**Action:** Investigate. Preserve evidence. Notify security team.

---

### 🟡 MEDIUM (Respond Soon - <24 hours)
- Single failed login from unusual location
- Occasional file processing errors
- Performance degradation
- Minor access control issues

**Action:** Document. Monitor. Investigate root cause.

---

### 🟢 LOW (Respond Normally - routine)
- Expected errors (invalid file format, user lockout)
- Normal access patterns
- Scheduled maintenance

**Action:** Document. Follow up in next review cycle.

---

## Incident Response Flowchart

```
Detect Incident
    ↓
Assess Severity
    ├─ CRITICAL → Isolate System (Stop →)
    ├─ HIGH     → Preserve Evidence
    ├─ MEDIUM   → Document
    └─ LOW      → Monitor
    ↓
Notify Contacts
    ├─ Security Officer
    ├─ HIPAA Officer (if PHI involved)
    ├─ Incident Response Team
    └─ Legal (if critical)
    ↓
Investigate Root Cause
    ├─ Review logs
    ├─ Check system status
    ├─ Identify affected systems/data
    └─ Determine scope
    ↓
Contain & Remediate
    ├─ Stop the attack
    ├─ Patch vulnerability
    ├─ Reset credentials
    └─ Restore from backup (if needed)
    ↓
Notify Affected Parties
    ├─ If PHI compromised: Patient notification (HIPAA requires <60 days)
    ├─ If business associate breached: Report to business partner
    └─ If regulatory breach: Report to HHS/OCR
    ↓
Post-Incident Review
    ├─ Document lessons learned
    ├─ Update procedures
    ├─ Implement preventive measures
    └─ Close ticket
```

---

## Incident Detection Examples

### Example 1: Failed Login Attack (HIGH)

**Detection:**
```
Audit log shows:
2026-06-10 14:22:15 - WARNING - Failed login attempt for user 'admin' (attempt 1)
2026-06-10 14:22:18 - WARNING - Failed login attempt for user 'admin' (attempt 2)
2026-06-10 14:22:21 - WARNING - Failed login attempt for user 'admin' (attempt 3)
2026-06-10 14:22:24 - WARNING - Failed login attempt for user 'admin' (attempt 4)
2026-06-10 14:22:27 - WARNING - Failed login attempt for user 'admin' (attempt 5)
2026-06-10 14:22:30 - WARNING - User 'admin' locked out for 15 minutes
```

**Severity:** HIGH

**Response:**
1. [ ] Review source IP of attempts
2. [ ] Check if VPN firewall rules are blocking
3. [ ] Contact user (jasmine/cb) to confirm account status
4. [ ] If external attack: enable IP blocking or WAF
5. [ ] Monitor for pattern (repeated attempts)
6. [ ] Document and close if one-time user error

---

### Example 2: Suspicious Download Pattern (HIGH)

**Detection:**
```
Audit log shows:
2026-06-10 03:15:22 - INFO - User 'cb' logged in successfully
2026-06-10 03:15:45 - INFO - User 'cb' downloaded: report_patient_001.xlsx
2026-06-10 03:15:52 - INFO - User 'cb' downloaded: report_patient_002.xlsx
2026-06-10 03:15:59 - INFO - User 'cb' downloaded: report_patient_003.xlsx
[... many more downloads in rapid succession ...]
```

**Severity:** CRITICAL (possible data exfiltration)

**Response:**
1. [ ] IMMEDIATELY isolate the system from network
2. [ ] Preserve the audit log (copy to secure location)
3. [ ] Contact CB to verify if they made these downloads
4. [ ] If unauthorized:
   - [ ] Force password reset for 'cb'
   - [ ] Check for account compromise (password shared? device infected?)
   - [ ] Determine what data was downloaded
   - [ ] Notify affected patients (HIPAA breach notification)
5. [ ] Audit log to see what data was accessed
6. [ ] Check other systems for similar patterns
7. [ ] Notify security, legal, and HIPAA officer
8. [ ] Follow breach notification procedures

---

### Example 3: System Error (MEDIUM)

**Detection:**
```
Audit log shows:
2026-06-10 10:30:15 - INFO - User 'jasmine' uploaded file: billing_06102026.xlsx
2026-06-10 10:30:22 - ERROR - Failed to process file: billing_06102026.xlsx - Disk full
```

**Severity:** MEDIUM

**Response:**
1. [ ] Check available disk space
2. [ ] Clear temporary files / old logs if needed
3. [ ] Retry file processing
4. [ ] Notify user of issue
5. [ ] Set up monitoring for disk space threshold
6. [ ] Schedule storage expansion if needed

---

## Step-by-Step Response Procedure

### Step 1: Confirm & Assess (5-10 minutes)

**[ ] Action Items:**
- [ ] Confirm the incident is real (not false alarm)
- [ ] Determine severity level (use matrix above)
- [ ] Identify affected systems (Billing App? Database? Network?)
- [ ] Determine affected data scope (which files? which users?)

**[ ] Questions to Answer:**
- What incident occurred? (specific event/error)
- When did it start? (timestamp from logs)
- When was it detected? (how long was it active?)
- Who detected it? (audit review? automated alert?)
- Is it ongoing? (still happening?)

---

### Step 2: Notify Contacts (5 minutes)

**[ ] Notify in Priority Order:**

For **CRITICAL** severity:
1. [ ] System administrator (isolate system if needed)
2. [ ] Security officer - IMMEDIATE
3. [ ] HIPAA compliance officer - IMMEDIATE
4. [ ] Incident response team lead
5. [ ] Legal (if breach suspected)

For **HIGH** severity:
1. [ ] Security officer - within 1 hour
2. [ ] HIPAA officer (if PHI involved)
3. [ ] Incident response team

For **MEDIUM/LOW** severity:
1. [ ] Security officer - within 24 hours
2. [ ] Document in incident log

**[ ] Notification Template:**
```
SECURITY INCIDENT REPORT
Date/Time Detected: [datetime]
Severity Level: [CRITICAL/HIGH/MEDIUM/LOW]
Detected By: [name/method]
Incident Type: [failed login / data breach / system error / etc]

SUMMARY:
[One sentence description]

TIMELINE:
- [timestamp]: [event]
- [timestamp]: [event]

AFFECTED SYSTEMS:
- [system 1]
- [system 2]

AFFECTED DATA:
- [data type/count]

INITIAL ASSESSMENT:
- Ongoing? Yes/No
- Contained? Yes/No
- PHI Exposed? Yes/No/Unknown

ACTION TAKEN:
- [action 1]
- [action 2]

NEXT STEPS:
- [investigation step]
- [remediation step]

Contact: [your name & phone]
```

---

### Step 3: Preserve Evidence (10 minutes)

**[ ] Immediately capture:**
- [ ] Copy audit log to secure backup: `cp audit_logs/audit_*.log /secure/backup/`
- [ ] Screenshot any error messages or alerts
- [ ] Document current system state (disk space, processes, network)
- [ ] Record any relevant metrics (CPU, memory, connections)

**[ ] Maintain chain of custody:**
- [ ] Don't modify original log files
- [ ] Use `cp` not `mv` (preserve originals)
- [ ] Record who accessed logs and when
- [ ] Store backups in secure location (encrypted)

---

### Step 4: Investigate Root Cause (1-4 hours)

**[ ] Investigation Steps:**

1. **Review Logs:**
   ```bash
   # Get detailed timeline
   cat audit_logs/audit_*.log | grep -A 5 -B 5 "ERROR\|FAILED\|suspicious_pattern"
   
   # Check for related events
   grep "username" audit_logs/audit_*.log
   
   # Export for detailed analysis
   cat audit_logs/audit_*.log > /secure/backup/incident_log_$(date +%s).txt
   ```

2. **Check System Status:**
   ```bash
   # Disk space
   df -h
   
   # Running processes
   ps aux | grep streamlit
   
   # Network connections
   netstat -tulpn | grep LISTEN
   
   # System logs
   journalctl -xe  # Linux
   ```

3. **Determine Scope:**
   - Which files were accessed/modified?
   - Which users were involved?
   - What data might be exposed?
   - How many records affected?

4. **Identify Root Cause:**
   - User error (accident)?
   - System misconfiguration?
   - Software bug?
   - Security attack?
   - Hardware failure?

---

### Step 5: Contain & Remediate (varies)

**[ ] For Failed Login Attack:**
- [ ] Block attacking IP at firewall
- [ ] Enable additional monitoring/WAF rules
- [ ] Reset affected user password(s)
- [ ] Require password re-entry on next login

**[ ] For Suspicious Access:**
- [ ] Revoke user session/credentials
- [ ] Force password reset
- [ ] Audit all recent activity by user
- [ ] Check for credential compromise (device infected?)

**[ ] For System Compromise:**
- [ ] ISOLATE system from network
- [ ] Preserve forensics (don't touch anything)
- [ ] Restore from clean backup if available
- [ ] Apply security patches
- [ ] Change all credentials

**[ ] For Data Breach:**
- [ ] Determine exactly what data was exposed
- [ ] Count affected individuals
- [ ] Determine if PII/PHI involved
- [ ] Prepare breach notification (see Section 6)

---

### Step 6: Notify Affected Parties (24-60 hours)

**ONLY if PHI/PII was actually exposed or accessed without authorization.**

**[ ] HIPAA Breach Notification Requirements:**

If **fewer than 500 residents**:
1. [ ] Individual notification to affected parties (within 60 days)
2. [ ] Notification to media (no requirement, unless notified individuals)
3. [ ] Report to HHS (within 60 days)

If **500+ residents**:
1. [ ] Individual notification (within 60 days)
2. [ ] Media notification required
3. [ ] HHS notification (within 60 days)

**[ ] Notification Content:**
- [ ] What happened (incident description)
- [ ] Date(s) of incident
- [ ] Date breach discovered
- [ ] What information was involved (specific data types)
- [ ] What is being done to fix it
- [ ] What individuals should do to protect themselves
- [ ] Organization's contact information

**[ ] Sample Notification:**
```
HIPAA BREACH NOTIFICATION

Dear [Patient Name],

We are writing to inform you of a security incident that may have involved your 
protected health information.

INCIDENT DETAILS:
- What happened: [description]
- When it occurred: [dates]
- What information was affected: [data types]
- How it happened: [root cause]

WHAT WE'RE DOING:
- Immediate action taken: [remediation]
- Ongoing monitoring: [measures]
- Prevention measures: [improvements]

WHAT YOU SHOULD DO:
- Monitor your accounts for unauthorized activity
- Report suspicious activity to [organization]
- Consider credit monitoring (if SSN/financial info involved)
- Resources available at [website]

CONTACT US:
[Phone, email, mailing address]

For more information about breach notification, visit:
www.hhs.gov/hipaa/
```

**[ ] Submit Breach Report to HHS:**
Visit: https://ocrportal.hhs.gov/ocr/breach/breach_report.action

---

### Step 7: Post-Incident Review (1 week later)

**[ ] After incident is contained, schedule debrief:**

**[ ] Document Lessons Learned:**
1. What was the root cause?
2. How could this have been prevented?
3. How quickly was it detected?
4. How quickly was it contained?
5. What gaps exist in controls?

**[ ] Update Procedures:**
- [ ] Update this incident response plan
- [ ] Improve monitoring/alerting
- [ ] Update security training
- [ ] Patch identified vulnerabilities

**[ ] Implement Preventive Measures:**
- [ ] Deploy additional monitoring
- [ ] Update firewall rules
- [ ] Increase authentication strength (MFA?)
- [ ] Enhance access controls

**[ ] Close Incident:**
- [ ] Archive all evidence
- [ ] Document final status
- [ ] Update incident log
- [ ] Schedule follow-up review (30 days)

---

## Incident Log Template

Keep a running log of all incidents:

```markdown
# Security Incident Log

| Date | Severity | Type | Summary | Root Cause | Resolution | Status |
|------|----------|------|---------|------------|------------|--------|
| 2026-06-10 | HIGH | Login Attack | 5 failed login attempts for 'admin' | User error (wrong pwd) | User notified, no action | CLOSED |
| 2026-06-12 | MEDIUM | File Error | Processing error on invalid file | Bad file format | User re-uploaded correct file | CLOSED |
```

---

## Emergency Contacts

**Fill in before deploying to production:**

| Role | Name | Phone | Email | Availability |
|------|------|-------|-------|--------------|
| Security Officer | ________________ | ________________ | ________________ | On-call? |
| HIPAA Officer | ________________ | ________________ | ________________ | Business hrs? |
| System Admin | ________________ | ________________ | ________________ | 24/7? |
| Legal Contact | ________________ | ________________ | ________________ | On-call? |
| Incident Commander | ________________ | ________________ | ________________ | On-call? |

**Escalation Path:**
1. First contact: _____________________
2. If unavailable: _____________________
3. If still unavailable: _____________________

---

## Testing the Incident Response Plan

**Recommended: Conduct tabletop exercise quarterly**

```
Scenario: "Unauthorized download of patient files detected in audit log"
1. Read scenario to team
2. Walk through steps 1-4 of response
3. Time how long each step takes
4. Identify gaps in procedure
5. Update plan based on findings
6. Document lessons learned
```

---

## Reference Documents

- SECURITY.md - Overview of security controls
- SECURITY_DEPLOYMENT.md - Deployment procedures
- AUDIT_LOG_REVIEW.md - Log monitoring procedures
- HIPAA Breach Notification Rule: https://www.hhs.gov/hipaa/for-professionals/breach-notification/
- OWASP Incident Response: https://owasp.org/

---

**Last Updated:** 2026-06-10  
**Next Review:** 2026-09-10 (quarterly)
