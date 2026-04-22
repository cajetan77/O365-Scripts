# Steps – App Registration with Certificate-Based Authentication

## Objective

Create an App Registration in Microsoft Entra ID and configure certificate-based authentication for access to both **Microsoft Graph API** and **SharePoint API**.

---

## Implementation Steps

### 1. Create App Registration

1. Log in to **Azure Portal**
2. Navigate to **Microsoft Entra ID**
3. Select **App registrations**
4. Click **New registration**
5. Enter the application name
6. Select the required supported account type
7. Click **Register**

---

### 2. Capture Application Details

Record the following:

* **Application (Client) ID**
* **Directory (Tenant) ID**

---

### 3. Upload Certificate

1. Open the App Registration
2. Navigate to **Certificates & secrets**
3. Select **Certificates**
4. Click **Upload certificate**
5. Upload the **.cer** file
6. Confirm the certificate is added successfully

---

### 4. Add Microsoft Graph API Permission

1. Go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions**
5. Add:

   * **Sites.FullControl.All**
6. Click **Add permissions**

---

### 5. Add SharePoint API Permission

1. In **API permissions**, click **Add a permission**
2. Select **SharePoint**
3. Choose **Application permissions**
4. Add:

   * **Sites.FullControl.All**
5. Click **Add permissions**

---

### 6. Grant Admin Consent

1. In **API permissions**, click **Grant admin consent**
2. Confirm consent for both:

   * **Microsoft Graph – Sites.FullControl.All**
   * **SharePoint – Sites.FullControl.All**

---

### 7. Configure Application Authentication

1. Store the certificate private key securely
2. Configure the application with:

   * Tenant ID
   * Client ID
   * Certificate thumbprint or certificate reference

---

### 8. Validate Access

1. Request an access token using the certificate
2. Test **Microsoft Graph API** access to SharePoint sites
3. Test **SharePoint API** access
4. Confirm the application can authenticate and access SharePoint successfully

---

## Validation

* App Registration created successfully
* Certificate uploaded successfully
* **Microsoft Graph – Sites.FullControl.All** added
* **SharePoint – Sites.FullControl.All** added
* Admin consent granted
* Authentication using certificate is successful
* Access to both Graph and SharePoint APIs is successful

---

## Rollback Plan

* Remove Microsoft Graph API permission
* Remove SharePoint API permission
* Remove uploaded certificate
* Delete App Registration if no longer required

---
