# Proposal E-Signature & Workflow Integration Architecture
## Automated BoldSign, Monday.com, and SharePoint Archival Pipeline

**Organization:** Hospitality Technologies | Engineering & Operations  
**Target Environment:** Render Cloud Platform  
**Document Type:** System Architecture & Integration Specification  

---

## Executive Summary & Core Objective

This specification details the end-to-end architectural workflow for transforming generated client presentation decks into electronically executed contracts. By establishing a dynamic gateway between the proposal document, BoldSign's e-signature engine, Monday.com's project tracking board, and Microsoft SharePoint via Graph API, this solution eliminates manual sales touchpoints, prevents link expiration, and guarantees automated contract archival.

---

## 1. End-to-End Workflow Pipeline

| Stage | System / Layer | Operational Action |
| :--- | :--- | :--- |
| **Phase 1: Deck Generation** | Render Web Service | Upload Base `.pptx` & Quote `.xlsx`. Web service parses quote figures and compiles presentation in memory (`io.BytesIO`). |
| **Phase 2: BoldSign Draft Creation** | BoldSign REST API | Web service automatically registers draft document in BoldSign with mapped signature fields, returning a persistent `documentId`. |
| **Phase 3: Interactive CTA in PDF** | Proposal PDF / PPTX | Final Acceptance slide embeds an active `"Ready to Move Forward? Click Here to Sign"` button pointing to the Render redirector gateway. |
| **Phase 4: Customer E-Signing** | BoldSign Hosted Session | Customer clicks link in PDF. Render gateway requests a fresh, non-expiring signing session URL and redirects client directly into the signing interface. |
| **Phase 5: Automated CRM & Archival** | BoldSign Webhook &rarr; Monday & SharePoint | Customer completes signature. BoldSign broadcasts webhook: Monday.com board status shifts to `"Signed"`, and signed PDF is streamed into the SharePoint folder. |

---

## 2. Dynamic Gateway & Link Expiration Prevention

### The In-Slide Call-to-Action (CTA) Element
On the final acceptance and disclosure slide, an eye-catching button shape is injected:
* **Label:** `"Ready to Move Forward? Click Here to Electronically Sign"`
* **Target:** Permanent Render gateway endpoint:
  ```
  https://quote-presentation-generator.onrender.com/sign/<PROPERTY_CODE>
  ```

### The Link Expiration Dilemma
BoldSign embedded signing URLs include an expiration timestamp (Time-To-Live). If a direct, static BoldSign link were hardcoded inside the customer's PDF proposal, the link could expire before the customer opens the file days or weeks later.

### Architectural Solution (Just-In-Time Token Resolution)
When the customer clicks the link in their proposal:
1. **Inbound Request:** The browser requests the Render gateway route (`/sign/<PROPERTY_CODE>`).
2. **On-Demand Token Request:** Render executes a secure server-to-server call to BoldSign (`GET /v1/document/getEmbeddedSignLink`) using the stored `documentId`.
3. **Fresh Token:** BoldSign generates a fresh, active signing token on demand.
4. **Instant 302 Redirect:** Render returns an immediate HTTP 302 Redirect, seamlessly opening the signature interface in the customer's browser.
5. **Real-Time Telemetry:** The gateway immediately logs an event or notifies Monday.com: *"Customer has opened proposal contract for signing."*

---

## 3. BoldSign API & Signature Anchor Strategy

BoldSign provides REST APIs that natively accept PPTX and PDF files. When the presentation generator compiles the deck in memory, it calls the BoldSign dispatch service:

* **Endpoint:** `POST https://api.boldsign.com/v1/document/send`
* **Key Payload Parameters:**
  * **`Title`:** `"Hospitality Technologies Agreement - [INN_CODE]"`
  * **`Files`:** Streamed directly from RAM (`io.BytesIO`)
  * **`DisableEmails`:** `true` (Suppresses BoldSign default emails so the sales team controls presentation delivery)
  * **`ExpiryDays`:** `14` (Enforces a 14-day expiration to protect against hardware pricing fluctuations)
  * **`Signers`:** Customer Name, Email, and role
  * **`CustomField`:** Embeds the Property Code or Monday Item ID for callback tracking
  * **`FormFields`:** Pre-positions Signature, Printed Name, Title, and Date fields onto the final Acceptance Slide coordinates.

### Automated Price Protection & Cleanup Policy
To prevent old, invalidated quotes from being signed and to keep the BoldSign dashboard clutter-free:
1. **14-Day Expiration:** The API payload strictly enforces a 14-day signing window via `ExpiryDays: 14`. After 14 days, the document status automatically shifts to **Expired**.
2. **Automated Deletion:** In the BoldSign Admin Dashboard (Settings > Document Settings), the **Automatic Document Deletion** policy must be enabled to automatically and permanently delete all documents with the status **Expired** after a specified retention period.

---

## 4. Real-Time Webhook Engine & Monday.com Synchronization

Upon full completion of the electronic signature ceremony, BoldSign immediately pushes an HTTP POST event to the Render service:

* **Webhook Endpoint:** `POST https://quote-presentation-generator.onrender.com/api/boldsign-webhook`
* **Event Trigger:** `DocumentCompleted` / `Signed`
* **Monday.com Board Target:** Board `9609739665` (WiFi Projects)

### Automated Actions on Monday.com
1. **Item Resolution:** Look up the corresponding project item by Property Code / INN Code.
2. **Mutate Status Column:** Update status to `"Signed"` or `"Contract Executed"`.
3. **Post Audit Note:** Automatically add an item update comment specifying the signatory name, verified email, IP address, and timestamp.

---

## 5. SharePoint Archival via Microsoft Graph API

Rather than leaving the executed PDF in BoldSign or requiring manual downloads, the system archives it into SharePoint automatically:

1. **Extract Destination Folder:** Read the existing SharePoint project folder link from Monday.com's SharePoint column.
2. **Encode Graph Path:** Convert the link into an encoded sharing reference using `u!` + Base64 encoding.
3. **Download Executed PDF:** Fetch the signed document and audit certificate directly from BoldSign (`GET /v1/document/download`).
4. **Stream to SharePoint:** Execute a PUT request to Microsoft Graph:
   ```http
   PUT https://graph.microsoft.com/v1.0/shares/{encoded_sp_url}/driveItem/root:/{Property_Code}_Signed_Proposal.pdf:/content
   ```
5. **Attach Back to Monday:** Update Monday.com's file/link column with the verified SharePoint URL of the signed contract.

---

## 6. Technical Prerequisites & Configuration Matrix

| Component | Key / Parameter | Configuration Location | Functional Role |
| :--- | :--- | :--- | :--- |
| **BoldSign API Key** | `BOLDSIGN_API_KEY` | Render Environment Variable | Authenticates document creation and embedded sign token retrieval. |
| **BoldSign Webhook** | `/api/boldsign-webhook` | BoldSign Portal Settings | Configured in BoldSign dashboard to alert Render when client signs. |
| **Monday.com Token** | `MONDAY_API_TOKEN` | Render Environment Variable | Authorizes GraphQL queries and mutations on Board `9609739665`. |
| **Microsoft Entra ID** | `MS_CLIENT_ID`<br>`MS_CLIENT_SECRET`<br>`MS_TENANT_ID` | Render Environment Variable | Authorizes OAuth2 app-only daemon access to upload executed contracts to SharePoint. |
| **Render Public Host** | `https://...onrender.com` | Render Dashboard | Hosts the web interface, dynamic gateway, and inbound webhook listener. |
