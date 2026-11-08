# Gap Analysis Summary - Example Fintech Ltd (FICTIONAL)

> Learning exercise with invented data. Numbers match the workbook's `Gap Summary` sheet.

## Control status (37-control Annex A sample)

| Status | Controls |
|---|---|
| Implemented | 10 |
| Partially implemented | 16 |
| Not implemented | 10 |
| Not applicable | 1 |

- Fully implemented: **27.8%** of 36 applicable controls
- At least partially implemented: **72.2%**

## By theme

| Theme | Implemented | Partial | Not implemented | N/A | % implemented |
|---|---|---|---|---|---|
| Organizational | 4 | 5 | 5 | 0 | 29% |
| People | 2 | 2 | 0 | 0 | 50% |
| Physical | 1 | 1 | 0 | 1 | 50% |
| Technological | 3 | 8 | 5 | 0 | 19% |

## Risks by rating

| Rating | Inherent | Residual (target) |
|---|---|---|
| Critical | 7 | 0 |
| High | 2 | 0 |
| Medium | 4 | 11 |
| Low | 0 | 2 |

## Highest inherent risks

| ID | Threat | Score | Rating | Treatment |
|---|---|---|---|---|
| R-01 | Phishing leads to compromised staff credentials | 16 | Critical | Mitigate |
| R-03 | Known vulnerability in an internet-facing service is exploited | 16 | Critical | Mitigate |
| R-09 | Security incident goes undetected for days | 16 | Critical | Mitigate |
| R-02 | Ransomware encrypts production servers | 15 | Critical | Mitigate |
| R-04 | Insider copies customer data | 15 | Critical | Mitigate |
| R-06 | Misconfigured cloud storage exposes data publicly | 15 | Critical | Mitigate |

## Priority gaps (controls not implemented)

- **A.5.19 Information security in supplier relationships**: Create a supplier security questionnaire and review the top 5 vendors
- **A.5.24 Information security incident management planning and preparation**: Write incident response procedure and escalation matrix
- **A.5.26 Response to information security incidents**: Create playbooks for phishing, ransomware and data leak
- **A.5.30 ICT readiness for business continuity**: Define recovery objectives and test a restore
- **A.5.36 Compliance with policies, rules and standards for information security**: Schedule an annual internal review
- **A.8.8 Management of technical vulnerabilities**: Monthly scans and a patch SLA by severity
- **A.8.9 Configuration management**: Adopt secure configuration baselines
- **A.8.16 Monitoring activities**: Introduce a SIEM or managed monitoring
- **A.8.22 Segregation of networks**: Segment production, staging and office networks
- **A.8.32 Change management**: Introduce change tickets and approvals

## Suggested sequence

1. Detection and response: A.8.15, A.8.16, A.5.24, A.5.26 (risk R-09 and support for all incident scenarios)
2. Vulnerability and change management: A.8.8, A.8.9, A.8.32 (risks R-02, R-03, R-06)
3. Resilience: A.5.30 and a tested restore for A.8.13 (risk R-08)
4. Suppliers: A.5.19 and A.5.20 (risk R-07)
5. Network segmentation: A.8.22 (risk R-02)
