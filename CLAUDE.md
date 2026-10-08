# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

An Obsidian community plugin that creates meeting notes from Outlook `.msg` files dragged and dropped onto its ribbon icon. It parses the files with `@kenjiuno/msgreader` and renders them with Mustache templates. It needs no Outlook add-in or Microsoft Graph access. The README is also the user documentation, and the settings tab links to it. Update the README whenever you add or change a template helper or setting.

## Commands

- `npm install`: install dependencies.
- `npm run dev`: esbuild watch mode. Writes `main.js` with inline sourcemaps.
- `npm run build`: type-checks with `tsc -noEmit -skipLibCheck`, then builds a minified production `main.js`.
- `npx eslint main.ts`: lint using `.eslintrc`. ESLint is not in `devDependencies`, so install it if needed.
- `npm version patch|minor|major`: bumps the version in `package.json` and runs `version-bump.mjs`, which syncs `manifest.json` and adds an entry to `versions.json` keyed to `minAppVersion`. Update `minAppVersion` in `manifest.json` by hand first if it changed.

There is no test suite. To test the plugin manually, copy `main.js`, `manifest.json` and `styles.css` into `<vault>/.obsidian/plugins/outlook-meeting-notes/`, then drop a `.msg` appointment onto the ribbon icon.

Releases: pushing a git tag triggers `.github/workflows/release.yml`. The workflow builds the plugin and creates a **draft** GitHub release with `main.js`, `manifest.json` and `styles.css`. Tags are the bare version number with no `v` prefix, as Obsidian requires.

## Architecture

All the code is in `main.ts`, which esbuild bundles into `main.js` as CommonJS targeting es2018. `obsidian`, `electron` and the CodeMirror packages are externals. `main.js` is a build artifact and is not committed.

Flow:
1. `onload` adds a ribbon icon and attaches `dragenter`/`dragleave`/`dragover`/`drop` listeners to it directly. Clicking the icon does nothing. The `is-being-dragged-over` drag-hover style is in `styles.css`.
2. `handleDropEvent` accepts exactly one file, reads it into an `ArrayBuffer` with `FileReader`, and builds a `MsgReader` from it.
3. `createMeetingNote` rejects any file whose `messageClass` is not `IPM.Appointment`. It then adds helpers to the parsed field data, renders the filename pattern, creates any missing folders, creates the note (or opens it if it already exists), and reveals the note in the file explorer through the private `app.internalPlugins` API.

Template rendering details:
- **Helpers** (`addHelperFunctions`): Mustache lambdas named `helper_<name>`, such as `helper_dateFormat` (moment format, with the argument split on `|`) and `helper_firstWord`. They are copied onto every object inside array fields, such as `recipients`, so they work inside Mustache sections. To add a helper, add it to the `helperFunctions` object and to the `func` union type. Helper *fields* such as `helper_currentDT` are set directly on the data in `createMeetingNote`.
- **Context-specific escaping** (`renderTemplate`): the template is split into a leading YAML frontmatter block (`---`…`---`) and a Markdown body, and each part is rendered with its own Mustache `escape` function. The YAML escaping turns multi-line values into `|` block scalars and quotes values containing `:#[]{},`. The Markdown escaping backslash-escapes Markdown and Obsidian syntax (`%%`, `~~`, `==`). Filenames use a third escaping function: `/` inside field values becomes the replacement character, and after rendering, the characters `*"\<>:|?` are replaced too.
- All rendering wraps the data in `mustache-validator`'s `proxyData`, so a template that references an unknown field throws an error (shown as a `Notice`) instead of rendering an empty string.

Settings (`OutlookMeetingNotesSettings`) are merged over `DEFAULT_SETTINGS` in `loadSettings`. If the filename pattern is set to an empty string, it reverts to the default.
