# Method Evaluation Tools Excel Add-in

The purpose of this add-in is to provide some statistical procedures that are
necessary for verifying for validating clinical laboratory methods but are not easy to implement
using builtin Excel spreadsheet functions. It is not meant to be a comprehensive
statistical analysis tool. The current interation allows the user to perform
linear regression techniques including Passing-Bablok, Deming, and Weighted Deming.

Procedures are also provide to allow users to analyse variance components as
described in CLSI EP15 and EP05.

This tools has been designed for use by laboratory professionals in association
with the accompanying
[method verification workbook](https://metools.chesher.id.au/assets/mvw.xlsx).

## Installation

The Add-In has yet to be added to Microsoft AppSource. However, if you would
like to try it out, the Add-in can be side-loaded.

### Side-loading the add-in

Follow these steps to sideload the add-in:

1. Download the add-in repository from https://github.com/dche658/method-eval-tools
   a. You can do this by clicking the green "Code" button and selecting
   "Download ZIP".
   b. Extract the contents of the ZIP file to a directory on your computer.
2. Download nvm-setup.exe for windows from https://github.com/coreybutler/nvm-windows/releases
3. Install nvm-setup.exe
4. Open a command prompt and run the following commands (update the version number
   as needed). Note, lts just stands for long-term-support.

```
nvm install lts
nvm use 24.6.0
```

5. Navigate to the directory where you extracted the downloaded repository
   and run:

```
npm install
npm run start
```

'npm install' will download and installall the required nodejs modules
and can take a while.

Startup can take a few seconds to complete. Once it is finished, excel should
open automatically with the add-in loaded.

## Copyright

Copyright (c) 2025 Doug Chesher. All rights reserved.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

