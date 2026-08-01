# Security policy

DocFence processes potentially hostile ZIP and XML input and treats report
redaction as a security property. Please report a suspected parser-bypass,
resource-exhaustion, unintended network/command execution, or report-content
leak through the repository's private security-reporting channel once it is
enabled. Do not include confidential source documents or sensitive package
content in a public issue.

Until a supported version is published, the current development line receives
security fixes. The parser limits and known boundaries are documented in
[docs/threat-model.md](docs/threat-model.md).
