# Contributing to TN-Montants

## Code

- Use `Option Explicit`.
- Prefix internal identifiers with the version/module prefix.
- Avoid generic identifiers such as `Unite`, `Dizaine`, `Montant`, or `Test`.
- Avoid excessive VBA line continuations.
- Keep Unicode construction compatible with the VBE.
- Do not add unnecessary external references.
- Test compilation in the target Office application.

## Linguistic changes

Any change affecting Arabic output must include:
1. affected input(s);
2. previous output;
3. proposed output;
4. linguistic reason;
5. regression-test impact.

Do not replace frozen expected results merely to make tests pass.

## Pull requests

Explain:
- what changed;
- why;
- affected Word/Excel behavior;
- regression results;
- Office version used for validation.
