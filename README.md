# App Script Starter Project

This is a sample apps script starter project that enables use of modern
JS tooling for both Apps Script and client-side development.

The starter project uses

* Typescript
* Rollup for bundling. This allows the use of ES modules as
  well as import NPM packages for both apps script and client-side scripts
* Lit for client-side Web Components
* Example of displaying HTML UIs in bound scripts and editor add-ons
* OAuth example code 

## Setup

For development:

* [Create a new Google Sheet and bound script.](https://developers.google.com/apps-script/guides/projects#create_a_project_from_google_docs_sheets_or_slides).
* Edit `.clasp.dev.json` and replace the `scriptId` property
  with the ID of your script.

For production:

* [Create a standalone script.](https://developers.google.com/apps-script/guides/projects#create_a_project_from_google_drive)
* Edit `.clasp.json` and replace the `scriptId` property
  with the ID of your script.

## Building

1. Install dependencies:

```sh
npm i
```

1. Buid the app:

```sh
npm run build
```

1. Deploy

For development:

```sh
npm run deploy:dev
```

For production:
```sh
npm run deploy
```

### Building in watch mode

Use `npm run deploy:watch` to run in watch mode. The add-on will
be continually built and deployed to the development script whenever
local files are changed.

## Projecct structure

* `./pages/` contains client-side HTML, javascript, and CSS files.
* `./server/` contains apps-script code executed server side
* `./shared/` contains shared files used across both environments

## Rollup.js config

While Apps Script supports most modern Javascript syntax, the environment itself imposes several constraints that that need
to be accounted for when developing:

* Modules are not supported. All script files exist in the same namespace.
* While HTML content can be served, other resources such as raw javascript and CSS files can not.
* Built-ins are different from standard browser & Node.js environments.

That said, tools like [Rollup.js](https://rollupjs.org/guide/en/) help bridge those gaps, allowing add-ons to be developed using modern syntax and module support.

The `rollup.config.js` contains rules for processing source files
to make them compatible with the Apps Script environment.

For Apps Script code:

* The `@rollup/plugin-typescript` plugin enables Typescript support.
* The `@rollup/plugin-node-resolve` plugin allows importing NPM 
  packages. Any package that does not rely on node or browser
  globals should work correctly. Imported packages are inlined
  in the transpiled script.

For client-side HTML:

* Both Typescript & Node resolution modules are enabled
* The `@web/rollup-plugin-html` plugin is used to process
  any local CSS and JS files included in HTML files. This is
  combined with an extension to inline those resources so they're
  served as a single file.
* Each HTML file is defined as a separate entry point to prevent
  chunking of shared code.

The `dist/` directory contains the transpiled code after the build is run.



