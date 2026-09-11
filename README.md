# TN-Montants V3.4

Conversion de montants en dinars tunisiens en arabe juridique classique.

## Word

Ruban personnalisé :
- **المبلغ المحدد**
- **مبالغ التحديد**
- **جميع مبالغ الوثيقة**

## Excel

Use:

`=TN_Montant(B2)`

The numeric value is converted to the Arabic legal wording while retaining the original numeric amount in parentheses.

## Regression safety

V3.4 includes a frozen Golden Master corpus containing the established reference outputs. A release should not proceed if `TN33_RunRegressionTests` reports a failure.

## Structure

- `src/` — VBA Word, Ribbon and Excel modules
- `ribbon/` — Ribbon XML
- `tests/` — frozen reference corpus and regression runner
- `examples/` — Excel worksheet automation example
- `documentation/` — installation and roadmap

Developed iteratively by Habib FARHAT with assistance from OpenAI ChatGPT.
