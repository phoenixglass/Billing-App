# Security Deployment Guide for PHI

This guide covers deploying the Billing App in a HIPAA-compliant manner for Protected Health Information (PHI) processing.

## Pre-Deployment Checklist

### 1. Credentials Setup ✓

**Generate password hashes:**
```bash
python3 env_setup.py
```

Follow the prompts to create one user account per person. This generates hashes for your `.env` file.

**Create `.env` file:**
```bash
# Copy output from env_setup.py
BILLING_APP_CREDENTIALS='{"username":"pbkdf2_sha256$..."}'
```

**Important:**
- ✓ `.env` is in `.gitignore` — never commit it
- ✓ Use unique, strong passwords (16+ characters, mixed case/numbers/symbols)
- ✓ One account per user (audit trail attribution)
- ✓ Store `.env` securely (encrypted, restricted access)

---

### 2. Audit Log Security ✓

**File permissions:**
The app automatically sets audit logs to `600` (owner read/write only).

Verify after startup:
```bash
ls -la audit_logs/
# Should show: -rw------- 1 user user audit_YYYYMMDD.log
```

**Encryption at rest:**
Audit logs MUST be encrypted. Options:

**Option A: Encrypted Volume (Recommended)**
```bash
# Linux with LUKS
sudo cryptsetup luksFormat /dev/sdX
sudo cryptsetup luksOpen /dev/sdX audit_volume
sudo mkfs.ext4 /dev/mapper/audit_volume
sudo mkdir /mnt/audit_logs
sudo mount /dev/mapper/audit_volume /mnt/audit_logs
sudo chown $(whoami) /mnt/audit_logs
# Update LOG_DIR in app.py to /mnt/audit_logs
```

**Option B: OS-Level Encryption**
- macOS: FileVault (System Settings → Security & Privacy)
- Windows: BitLocker (Settings → System → About)
- Linux: Full-disk encryption during OS installation

**Option C: Cloud Storage with Encryption**
- AWS S3 with SSE-S3 or SSE-KMS
- Azure Blob Storage with encryption
- Google Cloud Storage with CMEK

---

### 3. HTTPS/TLS (Required for Network Access)

**Local/VPN Deployment:**

If accessing only over VPN or localhost, HTTPS can be optional but is still recommended.

**Internet-Facing Deployment:**

HTTPS is MANDATORY. Use one of these approaches:

**Option A: nginx Reverse Proxy + Let's Encrypt**

```bash
# Install nginx and certbot
sudo apt-get install nginx certbot python3-certbot-nginx

# Create nginx config: /etc/nginx/sites-available/billing-app
upstream streamlit_app {
    server 127.0.0.1:8501;
}

server {
    listen 80;
    server_name your-domain.com;
    return 301 https://$server_name$request_uri;  # Redirect HTTP to HTTPS
}

server {
    listen 443 ssl http2;
    server_name your-domain.com;

    ssl_certificate /etc/letsencrypt/live/your-domain.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/your-domain.com/privkey.pem;
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers HIGH:!aNULL:!MD5;
    ssl_prefer_server_ciphers on;
    add_header Strict-Transport-Security "max-age=31536000; includeSubDomains" always;

    location / {
        proxy_pass http://streamlit_app;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}

# Enable and test
sudo ln -s /etc/nginx/sites-available/billing-app /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl enable nginx
sudo systemctl start nginx

# Get SSL certificate
sudo certbot certonly --nginx -d your-domain.com

# Auto-renew
sudo systemctl enable certbot.timer
```

**Option B: Streamlit Cloud**

Deploy directly to Streamlit Cloud (built-in HTTPS):
```bash
git push origin your-branch
# Create account at streamlit.io
# Connect GitHub repo and deploy
```

**Important:** Verify Streamlit has a Business Associate Agreement (BAA) if handling PHI.

**Option C: Cloud Provider (AWS/Azure/GCP)**

- AWS EC2 + Application Load Balancer (ALB)
- Azure App Service (built-in SSL)
- Google Cloud Run (built-in SSL)

All provide managed HTTPS. Verify their BAA coverage.

---

### 4. Network Security

**VPN/Private Network (Recommended for PHI):**
```bash
# Allow access only over VPN
# Use network firewall rules:
# - SSH: allowed from VPN only
# - HTTP/HTTPS: allowed from VPN only
# - All other: denied
```

**IP Whitelisting:**
```nginx
# In nginx config, restrict to known IPs
allow 192.168.1.0/24;      # Your office
allow 10.0.0.5;             # VPN gateway
deny all;
```

**Firewall Rules:**
```bash
# UFW (Ubuntu Firewall)
sudo ufw default deny incoming
sudo ufw default allow outgoing
sudo ufw allow from 192.168.1.0/24 to any port 22      # SSH from office only
sudo ufw allow from 10.0.0.0/8 to any port 443         # HTTPS from VPN
sudo ufw enable
```

---

### 5. Deployment Script

Create a safe deployment script:

```bash
#!/bin/bash
# deploy.sh - Deploy billing app with security checks

set -e  # Exit on error

# 1. Verify .env exists
if [ ! -f .env ]; then
    echo "ERROR: .env file not found"
    echo "Run: python3 env_setup.py"
    exit 1
fi

# 2. Check .env is not in git tracking
if git ls-files --error-unmatch .env 2>/dev/null; then
    echo "ERROR: .env is tracked by git! Remove it:"
    echo "git rm --cached .env"
    exit 1
fi

# 3. Verify permissions
if [ -f audit_logs/audit_*.log 2>/dev/null ]; then
    for log in audit_logs/audit_*.log; do
        perms=$(stat -c %a "$log")
        if [ "$perms" != "600" ]; then
            echo "WARNING: Log file $log has permissions $perms (should be 600)"
            chmod 600 "$log"
        fi
    done
fi

# 4. Load environment
source .env

# 5. Start app
echo "Starting Billing App..."
streamlit run app.py --server.port=8501 --logger.level=info
```

---

### 6. Ongoing Security Operations

**Daily:**
- [ ] Monitor for failed login attempts in audit logs
- [ ] Check for unusual file activity

**Weekly:**
- [ ] Review audit logs (see AUDIT_LOG_REVIEW.md)
- [ ] Check disk space (audit logs grow over time)

**Monthly:**
- [ ] Run `pip-audit` to check for dependency vulnerabilities
- [ ] Update dependencies if security patches available
- [ ] Review access logs for unauthorized access attempts

**Quarterly:**
- [ ] Full security audit
- [ ] Review and update access procedures
- [ ] Test incident response procedures

---

## Deployment Decision Tree

```
Are you accessing over VPN or local network only?
├─ YES → Local deployment (see below)
└─ NO → Internet-facing (HTTPS required, proceed to option selection)

LOCAL DEPLOYMENT:
├─ 1. Run: python3 env_setup.py
├─ 2. Create .env with output
├─ 3. Run: source .env && streamlit run app.py
├─ 4. Access: http://localhost:8501 (over VPN or direct)
└─ Audit logs stored locally on encrypted disk

INTERNET-FACING DEPLOYMENT:
├─ Choose TLS provider:
│  ├─ nginx + Let's Encrypt (most control)
│  ├─ Streamlit Cloud (easiest, verify BAA)
│  └─ Cloud provider (AWS/Azure/GCP)
├─ Setup credentials (env_setup.py)
├─ Configure encrypted audit log storage
├─ Setup monitoring/alerting
└─ Document incident response procedures
```

---

## Incident Response

See INCIDENT_RESPONSE.md for breach notification and logging procedures.

**Security incident contact:**
- [ ] Organization security officer: _____________
- [ ] HIPAA compliance officer: _____________
- [ ] Legal contact: _____________
- [ ] Incident response team lead: _____________

---

## Compliance Checklist

Before going live, verify:

- [ ] Credentials generated with unique passwords (env_setup.py)
- [ ] .env not committed to git
- [ ] Audit logs set to 600 permissions
- [ ] Audit logs on encrypted volume or encrypted filesystem
- [ ] HTTPS configured (if Internet-facing)
- [ ] Firewall rules configured
- [ ] Access procedures documented
- [ ] Incident response plan in place
- [ ] Staff trained on security procedures
- [ ] Business Associate Agreements reviewed
- [ ] Regular audit log review schedule established
- [ ] Monitoring and alerting configured

---

## Security Contacts & References

- **HIPAA Compliance:** https://www.hhs.gov/hipaa/
- **OWASP Top 10:** https://owasp.org/Top10/
- **CIS Controls:** https://www.cisecurity.org/cis-controls/
- **NIST Cybersecurity Framework:** https://www.nist.gov/cyberframework/

For questions, contact your organization's **security team** or **HIPAA compliance officer**.
