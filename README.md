# Method Evaluation Tools Excel Add-in

This Excel Add-in provides a suite of tools for evaluating and comparing the 
performance of different laboratory methods. It includes functionalities for 
calculating various performance metrics, visualizing results, and conducting 
statistical tests to assess the significance of differences between methods.

This tools has been designed for use by laboratory professionals in association
with the accompanying 
[method verification workbook](https://metools.chesher.id.au/assets/mvw.xlsx).
The current version provides various methods for linear
regression analysis, including Deming regression, Weighted Deming regression,
and Passing-Bablok regression along with their associated confidence intervals,
which are not easy to implement using the builtin functions in Excel.

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

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**