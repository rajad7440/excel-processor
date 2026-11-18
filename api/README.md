# RM Assist Excel processor — Teams / Copilot Studio integration

This update keeps the existing Streamlit app (`app.py`) in place and adds a separate API service under `api_service/`. The Streamlit UI is not modified. The API uses the same transformation rules copied from the existing app.

## API endpoints

- `GET /health` — health check.
- `POST /api/excel/process-base64` — recommended for Power Automate HTTP actions. JSON request:
  ```json
  {
    "file_name": "source.xlsx",
    "file_content_base64": "<base64-encoded file bytes>"
  }
  ```
  Returns the processed `.xlsx` as a binary HTTP response.
- `POST /api/excel/process` — multipart upload, useful for direct tests.

Set `PROCESSOR_API_KEY` in the hosting environment. When set, every processing request must include header `X-API-Key` with the same value. Do not commit the key to source control. Set `MAX_FILE_BYTES` if you need a different upload limit; default is 16 MiB.

## Run locally

From this folder:

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements-api.txt
# Set an API key before starting:
# Windows PowerShell: $env:PROCESSOR_API_KEY="replace-with-a-long-random-secret"
# macOS/Linux: export PROCESSOR_API_KEY="replace-with-a-long-random-secret"
uvicorn api:app --host 0.0.0.0 --port 8000
```

Health check: `http://localhost:8000/health`

## Test the API

```bash
curl -X POST "http://localhost:8000/api/excel/process" \
  -H "X-API-Key: YOUR_KEY" \
  -F "file=@sample.xlsx" \
  --output RM_Assist_Processed_Output.xlsx
```

For the base64 endpoint, send JSON with `file_name` and `file_content_base64`, and save the binary response as an `.xlsx` file.

## Power Automate / Copilot Studio flow outline

1. Create an agent flow with **When an agent calls the flow**.
2. Add a **File** input to the trigger and connect the flow to the topic's `UploadedFile` value.
3. Add an HTTP action:
   - Method: `POST`
   - URI: `https://YOUR-API-HOST/api/excel/process-base64`
   - Header: `Content-Type: application/json`
   - Header: `X-API-Key: YOUR_KEY` (store the key securely; don't hardcode it in exported flow definitions where avoidable)
   - Body must contain the original filename and Base64 file bytes from the trigger's File input.

   Power Automate's exact dynamic-content token for a File input can vary by trigger designer. Inspect the trigger output in a test run and map its file name and content bytes. If the content is exposed as a `$content` property, use that value as the Base64 field; if exposed as binary, convert it to Base64 with `base64(...)`. Do not send the whole trigger object as the Base64 string.
4. The HTTP action receives binary XLSX output. Add **OneDrive for Business — Create file** or **SharePoint — Create file**:
   - File name: `RM_Assist_Processed_Output.xlsx` (or a unique name)
   - File content: HTTP action Body.
5. Create a sharing link restricted to your organization or intended users. Avoid anonymous "Anyone" links unless explicitly approved.
6. Add **Respond to the agent** outputs such as `DownloadUrl`, `SourceRows`, `OutputRows`, and `FileName`; return the secure link to the topic.
7. In the topic, display the returned link in the Teams conversation.

Agent flows used synchronously as tools have a 100-second response constraint in Copilot Studio. For large workbooks or longer runs, use an asynchronous pattern rather than holding the agent response open.

## Notes / security

- The original `app.py` is preserved unchanged as requested.
- The original Streamlit source contains hardcoded login credentials. They were not copied into the API service, but they remain in the original app file. Rotate that password and move credentials into Streamlit secrets/environment variables before sharing the original app source or repository.
- Deploy this API separately from Streamlit (for example, Azure App Service or Azure Container Apps). This ZIP does not deploy it automatically.
- Use HTTPS in production and restrict access to the API.
