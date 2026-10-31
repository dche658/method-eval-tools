import officeAddinsConfig from "office-addin-lint/config/eslint.config.mjs";

export default [
  ...officeAddinsConfig,
  {
    languageOptions: {
      globals: {
        console: "readonly",
      },
    },
  },
];
