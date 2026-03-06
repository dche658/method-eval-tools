import { mkdir, cp, copyFile } from "node:fs";

try {
  mkdir("./dist/help/images", { recursive: true }, (err) => {
    if (err) console.log(err);
  });
  mkdir("./dist/help/styles", { recursive: true }, (err) => {
    if (err) console.log(err);
  });
  console.log("created help directory");
  copyFile("./documentation/add-in/index.html", "./dist/help/index.html", (err) => {
    if (err) console.log(err);
  });
  console.log("copied index help file")
  cp("./documentation/add-in/images","./dist/help/images", {recursive: true} , (err) => {
    if (err) console.log(err);
  });
  console.log("copied images")
  cp("./documentation/add-in/styles","./dist/help/styles", {recursive: true} , (err) => {
    if (err) console.log(err);
  });
  console.log("copied stylesheets")
} catch {
  console.error("The file could not be copied");
}
