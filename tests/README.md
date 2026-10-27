# Tests for ea-impact-matrix

Run from the skill folder:

```
python3 tests/run_tests.py            # run all tests
python3 tests/run_tests.py --report   # also rewrite tests/TEST_REPORT.md
```

Needs Python 3, `openpyxl` and LibreOffice (`soffice`) on the path. The runner writes only to a temporary folder, except the report.

| Path | What it is |
|---|---|
| `TEST_PLAN.md` | Objective, requirements R01 to R21, scenario, answer key, test cases, limits |
| `TEST_REPORT.md` | Latest results, generated |
| `run_tests.py` | The runner and the independent oracle |
| `scenario/` | Fictional client documents and the config produced from them |
| `fixtures/legacy_config.json` | A config in the old format, to prove backward compatibility |

All people and companies in the scenario are invented.
