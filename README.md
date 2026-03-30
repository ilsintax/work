# work Repository Structure

This folder is a GitHub-ready example for managing multiple static tools inside one repository.

## Suggested repository layout

- `auto-classifier/` : automatic classification and summary tool
- future tools can be added as sibling folders

## GitHub Pages usage

If your repository name is `work`, publish the repository root with GitHub Pages.

- Repository root page: `/work/`
- Auto-classifier page: `/work/auto-classifier/`

## Notes

- The rule set JSON file name can be changed freely.
- The app reads the JSON content, not the file name.
- `auto-classifier/index.html` currently keeps some CDN dependencies for Tailwind, Tabulator, and Google Fonts.
- `xlsx.full.min.js` is included locally in the app folder.
