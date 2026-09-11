# TN-Dinars2ArabicWords

Convert Tunisian dinar amounts from numerals to classical legal Arabic in Microsoft Word.

**Author / Maintainer:** Habib FARHAT  
**Version:** 3.4  
**Platform:** Microsoft Word / VBA

[Français](README.fr.md) · [العربية](README.ar.md)

## Overview

TN-Dinars2ArabicWords is a VBA module for Microsoft Word that converts Tunisian monetary amounts written in numerals into their amount-in-words form in classical Arabic, while preserving the original numeric amount in parentheses.

Example:

```text
2523.551 د
```

becomes:

```text
ألفان وخمسمائة وثلاثة وعشرون دينارا وخمسمائة وواحد وخمسون مليما (2523.551 د)
```

The project is intended especially for Tunisian legal and administrative documents: contracts, agreements, pleadings, correspondence, minutes and similar documents.

## Features

- Tunisian dinars and millimes.
- Three decimal digits for millimes.
- Western, Arabic-Indic and Persian digits.
- Common Tunisian currency markers such as `د`, `د.ت`, `DT`, `TND` and `دينار`.
- Singular, dual and plural monetary forms.
- Classical/legal Arabic wording.
- Conversion of a selected amount, a selection, or an entire document.
- Protection against repeated conversion.
- Unicode generation through `ChrW`.
- Regression tests for representative amounts.
- `TN33_` namespace prefix to reduce collisions in `Normal.dotm`.

## Examples

| Numeric amount | Generated text |
|---|---|
| `0.000 د` | `صفر دينار (0.000 د)` |
| `1.000 د` | `دينار واحد (1.000 د)` |
| `2.000 د` | `ديناران (2.000 د)` |
| `6.009 د` | `ستة دنانير وتسعة مليمات (6.009 د)` |
| `33.104 د` | `ثلاثة وثلاثون دينارا ومائة وأربعة مليمات (33.104 د)` |
| `2523.551 د` | `ألفان وخمسمائة وثلاثة وعشرون دينارا وخمسمائة وواحد وخمسون مليما (2523.551 د)` |
| `10004.585 د` | `عشرة آلاف وأربعة دنانير وخمسمائة وخمسة وثمانون مليما (10004.585 د)` |
| `20002.003 د` | `عشرون ألفا وديناران وثلاثة مليمات (20002.003 د)` |
| `50845.801 د` | `خمسون ألفا وثمانمائة وخمسة وأربعون دينارا وثمانمائة ومليم واحد (50845.801 د)` |
| `70575.300 د` | `سبعون ألفا وخمسمائة وخمسة وسبعون دينارا وثلاثمائة مليم (70575.300 د)` |
| `160426.591 د` | `مائة وستون ألفا وأربعمائة وستة وعشرون دينارا وخمسمائة وواحد وتسعون مليما (160426.591 د)` |

See [`examples/`](examples/) for more cases.

## Grammatical rules

### Dinars

- `1` → `دينار واحد`
- `2` → `ديناران`
- `3–10` → number + `دنانير`
- `11+` → number + `دينارا` where required by the implemented construction.

### Millimes

- `1` → `مليم واحد`
- `2` → `مليمان`
- `3–10` → number + `مليمات`
- `11+` → number + `مليما` where required.

The implementation distinguishes construct forms such as `ألفا دينار` from compound forms such as `ألفان وديناران`.

## Installation

1. Open Microsoft Word.
2. Press `Alt+F11`.
3. Select `Normal` / `Normal.dotm`.
4. Import `TN_Montants_V3_3.bas` using **File → Import File...**.
5. Run **Debug → Compile Normal**.
6. Save `Normal.dotm`.

Only enable VBA macros from sources you trust. Review the `.bas` source before installing it in a production environment.

## Usage

### Selected amount

Run:

```vb
TN33_ConvertirMontantSelectionne
```

Select a value such as `2523.551 د`.

### All amounts in a selection

Run:

```vb
TN33_ConvertirMontantsSelection
```

### All amounts in the document

Run:

```vb
TN33_ConvertirMontantsDocument
```

### Regression tests

Run:

```vb
TN33_TesterV33
```

Press `Ctrl+G` in the VBA editor to display the Immediate window.

## Architecture

The module separates:

1. Word integration and range handling.
2. Amount detection and normalization.
3. Arabic number generation.
4. Monetary grammar.
5. Unicode generation.
6. Regression testing.

The `TN33_` prefix is deliberate: generic names such as `Unite`, `Dizaine` or `NombreSousCent` can collide with existing procedures or variables in `Normal.dotm`.

## Limitations

The module is designed for Tunisian dinars and millimes. It is not a general Arabic number-to-words engine and does not attempt to cover every possible Arabic grammatical variant.

Legally significant documents should be reviewed before signing, filing or publication.

## Development and AI assistance

This project was developed iteratively by **Habib FARHAT** with assistance from **OpenAI ChatGPT (GPT-5.6 Luna)**.

AI assistance contributed to architecture, VBA debugging, Unicode handling, algorithm design, grammatical-rule implementation, testing strategy and documentation. Functional requirements, reference outputs, linguistic preferences and validation criteria were defined and reviewed during development.

This repository therefore represents a collaborative human–AI development process.

## Contributing

Contributions are welcome, especially corrections to Arabic forms, Tunisian legal examples, regression tests, amount detection, Word compatibility and documentation.

Changes to linguistic behavior should include regression tests.

## License

MIT License. See [`LICENSE`](LICENSE).

## Author

**Habib FARHAT**

Developed with assistance from OpenAI ChatGPT (GPT-5.6 Luna).

Microsoft Word and VBA are trademarks of Microsoft Corporation. This project is not affiliated with or endorsed by Microsoft or OpenAI.
