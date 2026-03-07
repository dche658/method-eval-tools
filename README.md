# AACB Method Evaluation Tools Excel Add-in

Purpose: The MEtools add-in provides specialized statistical procedures for the
verification and validation of clinical laboratory methods. It focuses on
essential calculations such as linear regression and variance component analysis
that are not natively available in standard Excel functions.

This tools has been designed for laboratory professionals to be used with the
accompanying 
[method verification workbook](https://metools.chesher.id.au/assets/mvw.xlsx).

Documents: \
[http://metools.chesher.id.au/help/]()

## Getting Started

The Add-In has yet to be added to Microsoft AppSource. However, if you would
like to try it out, the Add-in can be side-loaded.

### Side-load the add-in

Node.js must be installed on your computer. Rather than installing a particular
version of Node.js directly, installing via a node version manager is recommended.

* Download nvm-setup.exe for windows from https://github.com/coreybutler/nvm-windows/releases
* Install nvm-setup.exe
* Open a command prompt and run the following commands (update the version number
as needed). Note, lts is an abbreviation for long-term-support.

```
nvm install lts
nvm use 24.6.0
```

* If you are familiar with Git, clone the repository

```
git clone https://github.com/dche658/method-eval-tools.git
```

* Alternatively, download the add-in repository from
https://github.com/dche658/method-eval-tools
    a. You can do this by clicking the green "Code" button and selecting
    "Download ZIP". 
    b. Extract the contents of the ZIP file to a directory on
    your computer.
* Navigate to the directory where you extracted the downloaded repository 
and run:

```
npm install
```

. Start the add-in with:
```
npm run start
```

A separate terminal window will open, after which Excel will be launched.
Startup can take a few minutes to complete. 

When you have finished using Excel and the add-in, do not just close the 
terminal window. You must execute the stop command so the cache is cleared.
Otherwise, problems can arise when you try to run the add-in again.

```
npm run stop
```

## Contributing

Welcome, and thank you for your interest in contributing to MEtools!

There are many ways for you to contribute.

* Create [issues](https://github.com/dche658/method-eval-tools/issues)
* Create [documentation](https://github.com/dche658/method-eval-tools/tree/main/documentation)
* [Contributing code](https://docs.github.com/en/get-started/using-github/github-flow)

You can contribute as following:

## Copyright

Copyright (c) 2025 Doug Chesher. All rights reserved.

## License

License
Distributed under the MIT License. See [LICENSE](LICENSE) for more information

