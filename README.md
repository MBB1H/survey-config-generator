# Survey Config Generator

A Google Apps Script tool that generates JSON configurations for surveys from Excel files.

The generator parses the survey structure, identifies question types, builds navigation logic, calculates progress weights, applies Russian typography rules, validates the resulting survey flow, and produces a ready-to-use JSON configuration.

## Features

The generator automatically:

- detects the survey structure from an Excel file;
- supports `radio`, `checkboxes`, `select`, `numberInput`, and `textArea` question types;
- builds transitions between questions;
- handles supported conditional branching;
- calculates `blockWeight` and `progressWeight`;
- generates `nextIdPrecalculate` and `progressWeightPrecalculate` when required;
- adds service properties such as `isCollected`, `collectedKey`, and `defaultNextId`;
- detects custom answer fields and generates `otherOption` and `otherMaxLength`;
- generates the final survey result card;
- builds article lists using the `<list>` format;
- supports hidden Excel sheets;
- detects survey sheets named `Вопросы`, `Вопросы2`, `Вопросы3`, etc.

## Typography

Survey text is automatically processed before the JSON configuration is generated.

The typography module handles:

- non-breaking spaces;
- numeric ranges;
- particles such as `ли`, `же`, and `бы`;
- postfixes such as `-то`, `-либо`, and `-нибудь`;
- hyphenated constructions;
- numbers followed by measurement units, currencies, and percentages.

The typography rules are primarily designed for Russian-language surveys.

## Validation

The generator analyzes the survey flow and reports potential problems before the configuration is used.

It can detect:

- unresolved transitions;
- transitions to missing questions;
- circular navigation;
- unreachable questions;
- backward transitions;
- complex routes with multiple incoming transitions.

Warnings do not necessarily mean that the generated configuration is invalid. Some complex survey structures require manual review.

## Usage

1. Prepare the survey in an Excel file.
2. Open the generator.
3. Enter the survey slug and configure the result card if needed.
4. Upload the Excel file.
5. Generate the survey configuration.
6. Review validation warnings.
7. Copy the generated JSON.

## Current Limitations

The following features are not currently supported:

- `checkboxGroups`;
- multiple survey results;
- multiple survey blocks (`blocksCount > 1`);
- `placeholder`.

These cases require manual configuration.

## Project Structure

```text
Code.js          Google Apps Script entry points and server-side functions
SurveyUtils.js   Survey parsing, flow calculation, validation and JSON generation
Sidebar.html     Generator user interface
appsscript.json  Google Apps Script manifest
```

## Tech Stack

- JavaScript
- Google Apps Script
- HTML / CSS
- SheetJS

## Testing

The generator has been tested against more than 20 real-world Excel survey configurations, including surveys with linear navigation, conditional branching, different question types, custom answer fields, and complex progress calculations.

## Background

The project was created to automate a repetitive manual workflow: converting structured Excel survey specifications into JSON configurations.

The main goal is to reduce repetitive configuration work while keeping complex or ambiguous cases visible for manual review.