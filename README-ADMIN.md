# Add the screenshot importer to Fire Gremlins

Replace your existing `app.py` with this one and add `screenshot_admin.py` to the **same repository**.
Keep your existing `progress.xlsx`, `requirements.txt`, and `.streamlit/` folder.

Add these two lines to your existing `requirements.txt`:

```
openai>=1.80,<3
openpyxl>=3.1,<4
```

In **Streamlit Community Cloud → App settings → Secrets**, retain your existing
`WAR_COUNCIL_PASSWORD` and add:

```toml
ADMIN_PASSWORD = "a-different-strong-password"
OPENAI_API_KEY = "your-api-key"
```

Do not commit actual secrets to GitHub. OpenAI API usage is billed separately from ChatGPT.

After redeployment, sign in with your normal War Council password, then choose
**Admin / Upload Screenshots** in the sidebar and enter the admin password.
Choose an alliance and day; upload the appropriate alliance Excel workbook
(NONAME/MALICE must be supplied) and its screenshots. Extract, correct any
recognition errors, tick the review checkbox, and download the updated workbook.

**Important:** This is a review-first, download-based first version. It does not
automatically write to GitHub or a database. To publish the update, rename the
downloaded workbook `progress.xlsx` and replace that file in the matching
alliance's GitHub repository. Existing dashboards remain unchanged.
