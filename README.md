# Snowflake CoCo — Excel & Google Sheets Enablement

A practical Snowflake implementation demonstrating governed data access directly from Excel and Google Sheets.

## Architecture

Traditional:
Snowflake → SQL Export → CSV → Spreadsheet

With Snowflake CoCo:
Spreadsheet → OAuth PKCE → Snowflake User Role
→ Governance Policies → Query Results → Citations

## Implementation

- Database: COCO_SHEETS_DEMO
- Schemas: SALES, GOVERNANCE
- Analyst Role: COCO_SHEETS_ANALYST
- Steward Role: COCO_SHEETS_STEWARD
- OAuth authentication with PKCE
- Dynamic data masking
- Region-based row access policies
- Semantic views for business metrics
- Analyst-role validation
- Usage monitoring

## Repository Structure

- sql/ — Snowflake SQL scripts
- excel/ — Excel configuration
- google-sheets/ — Google Sheets configuration
- docs/ — Architecture and deployment guides
- images/ — Diagrams and screenshots

## Testing Note

Snowflake SQL implementation and governance validation should be performed in your own account.

Excel and Google Sheets add-in workflows are documented procedures, not end-to-end tested paths in this demonstration.

## Author

Satish Kumar  
Snowflake Chronicles

Learn. Build. Document. Share.

LinkedIn: https://www.linkedin.com/in/satishkumar-snowflake/
GitHub: https://github.com/satishkumarai
Medium: https://snowflakechronicles.medium.com/
