ONLINE EXAMINATION SYSTEM — STARTER

CONTENTS
- index.html: Landing page with Admin Login and Student Login buttons
- css/styles.css: Responsive mobile-friendly styling
- js/config.js: Supabase project URL and anon/publishable key placeholders
- js/app.js: Admin login, Excel validation/preview, and question import calls
- sql/01_admin_import_setup.sql: Admin user/session tables and secure login/import RPCs

FIRST SETUP
1. Open Supabase SQL Editor.
2. Run sql/01_admin_import_setup.sql.
3. The temporary admin credentials are:
   Username: MAHESH
   Password: mmm
4. Change this password before public launch. To change it, edit the seed password in the SQL and run the INSERT...ON CONFLICT block again, or create a dedicated password-reset function.
5. Open js/config.js and replace:
   SUPABASE_URL with your project URL
   SUPABASE_ANON_KEY with your anon/publishable key
   Never use the service_role key in browser code.
6. Serve the folder from a local web server (for example VS Code Live Server) or host it on HTTPS.
7. Click Admin Login, sign in, enter an existing Exam ID, select the workbook, click Validate & Preview, then Confirm & Import Questions.

EXCEL REQUIREMENTS
Headers:
QUESTIONID, SECTIONID, EXAMID, QNAME, OPT1, OPT2, OPT3, OPT4, ANSWER, HINT
- ANSWER: A, B, C, or D
- SECTIONID: MA, SCI, SS, MATH
- Number of questions must equal exam_details.max_marks
- The exam must currently have zero imported questions. Re-import/replace workflow is intentionally not enabled yet.
- The import SQL validates and inserts all questions in one database transaction.

NOTES
- Student Login is a placeholder and will be implemented in the next step.
- The temporary admin password is deliberately simple only for initial development. Do not publicly launch until it is changed to a strong password and admin protections are reviewed.
- Supabase URL/key are placeholders until you configure your project.
- This starter expects the tables and question retrieval/attempt tables created in the previous steps.
