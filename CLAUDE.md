# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

An Obsidian community plugin that creates meeting notes from Outlook `.msg` files (Outlook Classic) and `.ics` files (saved from new Outlook) dropped onto its ribbon icon. It parses them with `@kenjiuno/msgreader` and `ical.js` and renders them with Mustache templates. It needs no Outlook add-in or Microsoft Graph access. The README is also the user documentation, and the settings tab links to it. Update the README whenever you add or change a template helper or setting.

## Commands

- `npm install`: install dependencies.
- `npm run dev`: esbuild watch mode. Writes `main.js` with inline sourcemaps.
- `npm run build`: type-checks with `tsc -noEmit -skipLibCheck`, then builds a minified production `main.js`.
- `npm run lint`: lints the `.ts` files with ESLint (flat config in `eslint.config.mjs`).

There is no test suite. [CONTRIBUTING.md](CONTRIBUTING.md) covers manual testing in Obsidian, the release process and this fork's branch model. In short, `dev` is the default branch, `master` mirrors upstream, and topic branches come off `dev`.

## Architecture

The plugin is `main.ts`, plus `ics.ts` for `.ics` parsing. esbuild bundles them into `main.js` as CommonJS targeting es2018. `obsidian`, `electron` and the CodeMirror packages are externals. `main.js` is a build artifact and is not committed.

Flow:
1. `onload` adds a ribbon icon and attaches `dragenter`/`dragleave`/`dragover`/`drop` listeners to it directly. Clicking the icon does nothing. The `is-being-dragged-over` drag-hover style is in `styles.css`.
2. `handleDropEvent` accepts exactly one file and reads it into an `ArrayBuffer` with `FileReader`. It passes `createMeetingNote` a function that turns the file into template data: `msgToTemplateData` (MsgReader) or, for `.ics` files, `icsToTemplateData` in `ics.ts`, which maps iCalendar fields onto the same names MsgReader uses (`subject`, `apptStartWhole`, `recipients`, …) so templates work for both.
3. `createMeetingNote` calls that function inside its `try`, so parse errors (including `msgToTemplateData` rejecting a `.msg` whose `messageClass` is not `IPM.Appointment`) show as a `Notice`. It then adds helpers to the data, renders the filename pattern, creates any missing folders, creates the note (or opens it if it already exists), and reveals the note in the file explorer through the private `app.internalPlugins` API.

Template rendering details:
- **Helpers** (`addHelperFunctions`): Mustache lambdas named `helper_<name>`, such as `helper_dateFormat` (moment format, with the argument split on `|`) and `helper_firstWord`. They are copied onto every object inside array fields, such as `recipients`, so they work inside Mustache sections. To add a helper, add it to the `helperFunctions` object and to the `func` union type. Helper *fields* such as `helper_currentDT` are set directly on the data in `createMeetingNote`.
- **Context-specific escaping** (`renderTemplate`): the template is split into a leading YAML frontmatter block (`---`…`---`) and a Markdown body, and each part is rendered with its own Mustache `escape` function. The YAML escaping turns multi-line values into `|` block scalars, and double-quotes values that contain `:#[]{},`, start with a YAML indicator character, or have leading or trailing whitespace. The Markdown escaping backslash-escapes Markdown and Obsidian syntax (`%%`, `~~`, `==`). Filenames use a third escaping function: `/ # ^ [ ]` and control characters inside field values become the replacement character, and after rendering, `*"\<>:|?#^[]` and control characters are replaced too. Note paths go through `normalizePath`.
- All rendering wraps the data in `mustache-validator`'s `proxyData`, so a template that references an unknown field throws an error (shown as a `Notice`) instead of rendering an empty string.

Settings (`OutlookMeetingNotesSettings`) are merged over `DEFAULT_SETTINGS` in `loadSettings`. If the filename pattern is set to an empty string, it reverts to the default.
