# Submission Checklist

### 1. Determine your launch time line.

DOUG: I have good availability from now until March 2027. I will be unavailable from 18-20 November and 22-28 December.

### Step 2: Review all the Microsoft 365 app policies

See metools_proposition.md

### Step 3: Check that your manifest is compliant

The validation check returned the following information.

**Validation Information:**
Manifest ID Valid Prefix: The product ID in the manifest has a valid prefix
  - Details: 2fd3ebf4-8c70-447c-9ebe-40fa41f89128

Manifest ID Correct Structure: The structure of the product ID is correct.
  - Details: 2fd3ebf4-8c70-447c-9ebe-40fa41f89128

Package Type Identified: Package of your add-in was parsed successfully.

Correct Package: Your package matches the submission type.

Valid Manifest Schema: Your manifest does adhere to the current set of XML schema definitions for Add-in manifests.

Manifest Version Correct Structure: The manifest version number has the correct structure for the platform that it supports.

Manifest Version Correct Value: The manifest version number is greater or equal to 1.0.

Desktop Source Location Present: A desktop or default source location URL is found.

Secure Desktop Source Location: The manifest desktop source location URLs use HTTPS.

The manifest source location URLs are valid.: The manifest source location URLs are valid.

Supported Office Identified: Supported Office products were successfully determined.

Support URL Present: The manifest support URL is present.
  - Details: https://metools.aacb.asn.au/help

Valid Support URL structure: The manifest support URL has valid structure.

Valid OnlineMeetingCommandSurface ExtensionPoint.: OnlineMeetingCommandSurface ExtensionPoint extracted from manifest is found to be valid.

High Resolution Icon Present: A high resolution icon element was expected and is present.
  - Details: https://metools.aacb.asn.au/assets/icon-64.png

Supported High Resolution Icon URL File Extension: The manifest high resolution icon URL has a valid image file extension.
  - Details: png

Secure High Resolution Icon URL: The manifest high resolution icon URL uses HTTPS.
  - Details: https://metools.aacb.asn.au/assets/icon-64.png

Icon Present: A icon element was expected and is present.
  - Details: https://metools.aacb.asn.au/assets/icon-32.png

Supported Icon URL File Extension: The manifest icon URL has a valid image file extension.
  - Details: png

The manifest icon URL uses HTTPS.: Secure Icon URL
  - Details: https://metools.aacb.asn.au/assets/icon-32.png

All GetStarted strings are present in Resources: All GetStarted strings are present in Resources

Acceptance Test Completed: Acceptance test service has finished checking provided add-in.


Based on the requirements specified in your manifest, your add-in can run on the following platforms; your add-in will be tested on these platforms when you submit it to the Office Store:
  - Excel 2019 or later on Mac
  - Excel on Windows (Microsoft 365)
  - Excel on the web
  - Excel on Mac (Microsoft 365)
Important: This analysis is based on the requirements specified in your manifest and does not account for any runtime JavaScript calls within your add-in. For information about which API sets and features are supported on each platform, see Office Add-in host and platform availability. (https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).

*This does not include mobile apps. You can opt-in to support mobile apps when you submit your add-in.

**The manifest is valid.**

### Step 4: Decide what platforms you want to support

- Microsoft 365 Excel (I do not possess a Mac, and have not been able to test)
- Excel on the web

### Step 5: Provide testing information, including for any additional purchases of required external services or SaaS offers

- Microsoft 365 Excel: No additional services required. Test data is available in the sample method verification template that can be downloaded from the add-in website. https://metools.chesher.id.au/assets/mvw.xlsx

### Step 6: Include critical information links with https:// URLs

The addresses for a support URL, a privacy policy URL, and an end user license agreement (EULA) URL are required. One of the top reasons an app submission fails our validation process is when these links aren't included in submission. Any provided links must be valid. The following image shows the box that asks for the support information.

- Application Support URL: https://metools.aacb.asn.au/help/
- AACB Support URL: https://aacb.asn.au/AACB/AACB/Contact-Us.aspx
- Privacy Policy URL: https://metools.aacb.asn.au/help/privacy_policy.html (Draft only. Needs review prior to submission)
- End User License Agreement (EULA) URL: https://metools.aacb.asn.au/help/#_license

### Step 7: Prepare your store listing with your team

You need to provide the correct metadata indicating how you want your app to be listed in the store, including:

**Your app name:**

AACB Method Evaluation Tools

**A short description of your app and its value.**

Excel add-in to assist with validation or verification of clinical laboratory methods.

**A long description of your app with HTML formatting.**

This Excel add-in is targetted at clinical laboratory professionals, including scientists and pathologists. It delivers specialized statistical procedures used for verifying and validating clinical pathology methods. It provides essential calculations missing from native Excel functions, including Weighted Deming regression, Passing-Bablok regression, and variance component analysis. Additionally, the add-in features dedicated utilities to help users derive or verify clinical laboratory test reference intervals. The add-in has been developed by the Australsian Association for Clinical Biochemistry and Laboratory Medicine (AACB). A <a href="https://metools.chesher.id.au/assets/mvw.xlsx">method verification template</a> that makes optimal use of the add-in is available for download from the Add-in website.

**Marketplace Icon**

metools-icon-256.png (Sent separately)