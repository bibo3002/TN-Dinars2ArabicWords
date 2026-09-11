# TN-Montants — Tests V3.4

## Golden Master

`expected-results.csv` is the frozen reference corpus.

The regression runner compares the current V3.3/V3.4 conversion result byte-for-byte with the stored expected result. It never generates or updates the expected result.

### Run

1. Import `tests/TN33_RegressionTests.bas` into the Word VBA project.
2. Keep `TN_Montants_V33.bas` installed.
3. Run `TN33_RunRegressionTests`.
4. Select `expected-results.csv`.
5. Require `RESULT : PASS` before releasing a change.

## Rule

If a linguistic output changes intentionally, do not silently update the CSV. First document the change, review the affected cases, and explicitly approve a new reference baseline/version.
