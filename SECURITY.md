# Security Features

This application includes the following security controls for handling Protected Health Information (PHI):

## Authentication & Authorization

### Password Protection
- Simple password authentication required to access the application
- Default credentials: username='admin', password='billing2026'
- **⚠️ IMPORTANT**: Change default password before production use
- Password is hashed using SHA-256 (not stored in plain text)

### Session Management
- **Automatic timeout**: 15 minutes of inactivity
- Sessions expire and require re-authentication
- Last activity tracked to enforce timeout

## Audit Logging

### What is Logged
All PHI access is logged to `audit_logs/audit_YYYYMMDD.log`:
- User login attempts (success and failure)
- File uploads (filename, size, timestamp)
- File processing operations
- File downloads
- Session timeouts
- Errors and security events

### Log Format
```
2026-01-26 10:30:15 - INFO - User 'admin' logged in successfully
2026-01-26 10:31:22 - INFO - User 'admin' uploaded file: billing_01252026.xlsx (245678 bytes)
2026-01-26 10:31:25 - INFO - Successfully processed file: billing_01252026.xlsx, generated 6 output files
2026-01-26 10:32:10 - INFO - User 'admin' downloaded: billing_01252026_Rosanna.xlsx
```

### Log Protection
- Logs are excluded from git via `.gitignore`
- Logs contain timestamps, usernames, and actions
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
- Password protection
- Session management
- Audit logging
- Secure file handling

### Additional Steps Needed
- [ ] Change default password
- [ ] Restrict file system permissions on audit logs
- [ ] Regular review of audit logs
- [ ] Document access procedures
- [ ] Train users on security policies

### For Production/Internet Deployment
Additional requirements for HIPAA compliance:
- [ ] Deploy behind HTTPS with valid certificates
- [ ] Implement stronger authentication (OAuth, SSO, MFA)
- [ ] Encrypt audit logs at rest
- [ ] Implement role-based access control (RBAC)
- [ ] Regular security audits and penetration testing
- [ ] Business Associate Agreements (BAAs)
- [ ] Incident response plan
- [ ] Data backup and disaster recovery
- [ ] Physical security controls

## Limitations

⚠️ **Important**: These security controls provide a baseline but are NOT sufficient for full HIPAA compliance without additional infrastructure-level controls:

- No encryption at rest for temp files (relies on OS/filesystem encryption)
- Simple password authentication (should use MFA in production)
- No role-based access control
- No intrusion detection/prevention
- No data loss prevention (DLP)

## Changing the Default Password

To change the default password:

1. Generate a new password hash:
```python
import hashlib
new_password = "your_secure_password_here"
hash_value = hashlib.sha256(new_password.encode()).hexdigest()
print(hash_value)
```

2. Replace the `correct_hash` value in `app.py` line ~42:
```python
correct_hash = "your_new_hash_here"
```

3. Update the login page caption to remove default credentials display

## Security Contact

For security issues or questions, contact your organization's security team or HIPAA compliance officer.

## Compliance Status

✅ **Implemented Controls**:
- Authentication
- Session management  
- Audit logging
- Input validation
- Secure file cleanup

⚠️ **Partial/Missing Controls**:
- Encryption at rest (depends on OS/filesystem)
- Multi-factor authentication
- Role-based access control
- Intrusion detection
- Full HIPAA compliance certification

**Recommendation**: Consult with a HIPAA compliance specialist before using in a production healthcare environment.
