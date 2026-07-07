# Security Features

This application includes the following security controls for handling Protected Health Information (PHI):

## Authentication & Authorization

⚠️ **The app has no authentication.** A previous version required per-user
username/password login (PBKDF2-HMAC-SHA256, brute-force lockout, 15-minute
session timeout); that gate was removed at the owner's request, and the app
now opens directly to anyone who reaches the URL, with no session or access
control of any kind. If this app is deployed anywhere reachable by more than
its intended users, access must be restricted at the infrastructure level
(private network/VPN, a reverse proxy with its own auth, Streamlit Cloud
"private" visibility, etc.) — the application itself will not stop anyone
who has the link.

## Audit Logging

### What is Logged
File activity is logged to `audit_logs/audit_YYYYMMDD.log`:
- File uploads (filename, size, timestamp)
- File processing operations
- File downloads
- Errors and security events

Since there is no authentication, log entries are not attributed to an
individual user — anyone with the link can perform these actions.

### Log Format
```
2026-01-26 10:31:22 - INFO - Uploaded file: billing_01252026.xlsx (245678 bytes)
2026-01-26 10:31:25 - INFO - Successfully processed file: billing_01252026.xlsx, generated 6 output files
2026-01-26 10:32:10 - INFO - Downloaded: billing_01252026_Rosanna.xlsx
```

### Log Protection
- Logs are excluded from git via `.gitignore`
- Logs contain timestamps and actions (no username, since there's no login)
- Store logs securely with restricted access

## File Security

### Input Validation
- File size limit: 50MB maximum
- File type validation: Only .xlsx files accepted
- Filename validation: Blocks suspicious characters (`..`, `/`, `\`, etc.)
- Prevents directory traversal attacks

### Secure File Handling
- Temporary files created in system temp directory
- **Secure deletion**: Files overwritten with random data before deletion
- Automatic cleanup in `finally` block ensures files are always removed
- In-memory processing where possible (BytesIO)

## Data Protection

### Encryption in Transit
- **Local Use**: Run locally, no network transmission
- **Network Use**: Deploy behind HTTPS/TLS (nginx, Apache, or Streamlit Cloud)
- VPN recommended for remote access

### Encryption at Rest
- Temporary files deleted immediately after processing
- No persistent storage of PHI data
- Audit logs should be stored on encrypted volumes

## Deployment Recommendations

### For Local/VPN Use
✅ Current security features are implemented:
- Audit logging
- Secure file handling

⚠️ There is no authentication. Access must be restricted at the network
level (local machine only, or VPN) since the application will not do it.

### Additional Steps Needed
- [ ] Restrict network/deployment access since the app has no login of its own
- [ ] Restrict file system permissions on audit logs
- [ ] Regular review of audit logs
- [ ] Document access procedures
- [ ] Train users on security policies

### For Production/Internet Deployment
Additional requirements for HIPAA compliance:
- [ ] Deploy behind HTTPS with valid certificates
- [ ] Add authentication (this app has none) — OAuth, SSO, MFA, or a
  reverse-proxy auth layer
- [ ] Encrypt audit logs at rest
- [ ] Implement role-based access control (RBAC)
- [ ] Regular security audits and penetration testing
- [ ] Business Associate Agreements (BAAs)
- [ ] Incident response plan
- [ ] Data backup and disaster recovery
- [ ] Physical security controls

## Limitations

⚠️ **Important**: These security controls provide a baseline but are NOT sufficient for full HIPAA compliance without additional infrastructure-level controls:

- **No authentication of any kind** — anyone with the URL has full access
- No encryption at rest for temp files (relies on OS/filesystem encryption)
- No role-based access control
- No intrusion detection/prevention
- No data loss prevention (DLP)

## Security Contact

For security issues or questions, contact your organization's security team or HIPAA compliance officer.

## Compliance Status

✅ **Implemented Controls**:
- Audit logging
- Input validation
- Secure file cleanup

⚠️ **Partial/Missing Controls**:
- Authentication (none — removed at the owner's request)
- Session management
- Encryption at rest (depends on OS/filesystem)
- Multi-factor authentication
- Role-based access control
- Intrusion detection
- Full HIPAA compliance certification

**Recommendation**: Consult with a HIPAA compliance specialist before using in a production healthcare environment.
