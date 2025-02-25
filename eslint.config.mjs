import eslint from "@eslint/js";
import tseslint from "typescript-eslint";

export default tseslint.config(
  eslint.configs.all,
  tseslint.configs.all,
  {
    rules: {
      "@typescript-eslint/no-magic-numbers": "off",
      "no-continue": "off",
      "one-var": "off",
    },
  },
  { files: ["src/**/*.ts"] },
  { ignores: ["eslint.config.mjs"] },
  {
    languageOptions: {
      parserOptions: {
        ecmaFeatures: { impliedStrict: true },
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
        sourceType: "commonjs",
      },
    },
  }
);
