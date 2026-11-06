-- OAuth security integration shared by CoCo for Excel and CoCo for Google Sheets.
-- Use IF NOT EXISTS, not CREATE OR REPLACE: replacing the integration rotates the client ID
-- and invalidates refresh tokens for every user.
-- Requires ACCOUNTADMIN or the global CREATE INTEGRATION privilege.

CREATE SECURITY INTEGRATION IF NOT EXISTS COCO_EXCEL_OAUTH
  TYPE = OAUTH
  OAUTH_CLIENT = CUSTOM
  OAUTH_CLIENT_TYPE = 'PUBLIC'
  OAUTH_REDIRECT_URI = 'https://ai.snowflake.com/static/excel/oauth/callback.html'
  OAUTH_ENFORCE_PKCE = TRUE
  OAUTH_ISSUE_REFRESH_TOKENS = TRUE
  OAUTH_REFRESH_TOKEN_VALIDITY = 86400
  ENABLED = TRUE
  COMMENT = 'Snowflake CoCo for Excel and Google Sheets demo';

-- Retrieve the client ID for the add-in pane. Do not commit the output.
-- SELECT SYSTEM$SHOW_OAUTH_CLIENT_SECRETS('COCO_EXCEL_OAUTH');
