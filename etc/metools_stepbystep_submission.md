# Microsoft Marketplace step-by-step submission

### Step 1: Select the type of app you're submitting

- Sign into Partner Center.
- Select Marketplace offers tile.
- Select the Microsoft 365 and Copilot tab. If not present then open a developer account (see https://learn.microsoft.com/en-us/partner-center/marketplace-offers/open-a-developer-account).
- Select "+ New offer" -> "Office add-in"

### Step 2: Name your app

- Enter the name of the app: "AACB Method Evaluation Tools"
- Check availability to verify the the name you chose is available. (Very unlikely that it won't be)
- Associate the enw offer with a publisher. In this case, accept the publisher account you are signed in to.
- Select "Create".

### Step 3: Tell us about your product setup

1. When filling out the Product setup page.
    - Will your app be listed on the Apple Store? No
    - Does your app use Microsoft Entra ID or single sign-on (SSO)? No
    - Does your app require additional purchases? No
    - Do you want to connect with your lead management CRM system? No (the app does not collect any customer information)
2. Save draft before continuing to the next page: Packages

### Step 4: Upload your manifest for package testing

- A copy of the manifest.xml file is attached to this email but the most up to date version is available on GitHub at https://github.com/dche658/method-eval-tools/blob/main/manifest.xml. At the top right of the window you will see and icon to download the raw file. Click on it to download the manifest.xml file.

### Step 5: Define the metadata that categorizes your app in the store

1. On the Properties page, select at least one and up to three categories.
2. You can optionally select up to two industries. If none of the options are relevant, then do not select any.
3. Under Legal and support info. If the AACB is happy to release the software under the MIT Licence, then the link is https://metools.aacb.asn.au/help/#_license. If not, then further discussions will need to be had before the submission is made and the documents updated accordingly.
4. Privacy policy link. My draft privacy policy link is https://metools.aacb.asn.au/help/privacy_policy.html. If changes are required or if you have an AACB privacy policy, then link there. The yellow warning reminds you of the following things your privacy policy must include to pass certification.
    * Information on your policies regarding user's personal information.
    * Refer to the app or your service overall and NOT your website.
    * Your service description must include the name of app submitted.
    * A valid URL link that doesn't generate a 404 error.
5. Support document link: https://metools.aacb.asn.au/help/
6. Select "Save"

### Step 6: Define your languages in Marketplace listings

- Currently, only English is supported.

### Step 7: Not applicable

### Step 8: Decide on your availability date

- I'll leave this up to you.
- Select "Save Draft"

### Step 9: Make sure you add your critical testing instructions

No test accounts, licence keys, or testing credentials are required.

To test the =METOOLS.BOXCOX(x,lambda) custom function, you can use the following data

For lambda = 0.5
x = 10 -> 4.325
x = 20 -> 6.944
x = 100 -> 18.000
x = 2000 -> 87.443

For lambda = 0, METOOLS.BOXCOX(x, 0) = LN(x)
x = 10 -> 2.303
x = 20 -> 2.996
x = 100 -> 4.605
x = 2000 -> 7.601










