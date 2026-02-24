# SECURITY.md

# Security Policy

## Supported Versions

Only the latest release is supported with security updates.

## Reporting a Vulnerability

Please do not open public issues for security vulnerabilities.

Instead email:

security@your-domain.example

Include a description of the issue, reproduction steps, affected version, and potential impact.

## Scope

`pptxrender` does not implement a PPTX parser, delegates rendering to Microsoft PowerPoint, and performs local file operations only.

Security issues are therefore most likely related to file handling, automation permissions, or command execution boundaries.

## Disclosure

Responsible disclosure is appreciated. Credit will be given unless anonymity is requested.