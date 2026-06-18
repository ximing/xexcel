# Security

Please do **not** open a public GitHub issue for a vulnerability.

Report privately via [GitHub Security Advisories](https://github.com/ximing/xexcel/security/advisories/new) on this repository.

xexcel runs in the browser and in Node. Typical concerns: malicious xlsx/CSV input, formula evaluation on untrusted workbooks, and stored XSS if a host app renders cell values as HTML (this project does not). Include a repro and impact if you can.
