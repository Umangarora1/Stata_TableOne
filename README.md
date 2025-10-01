# Stata_TableOne
"export_table" is a Stata 14+ utility that exports a tidy two-group summary table to Excel with correct denominators, tests, and sensible handling of categorical vs continuous variables.

Output file: Table Export Group.xlsx (created/updated via putexcel)

Groups: defined by by(); must have exactly two levels (numeric or string)

Columns:
B/C = Group 0: n (row%) | N
D/E = Group 1: n (row%) | N
F/G = p-value | test
H/I = Total: n (row%) or mean±SD or median (IQR) | N

Installation

Put export_table.ado somewhere on your adopath (see adopath in Stata).
Typical user folders:

Windows: c:\ado\personal\

macOS: ~/Library/Application Support/Stata/ado/personal/

Linux: ~/ado/personal/

Verify Stata can see the files:

adopath
findfile export_table.ado
which export_table

Syntax
export_table varlist , by(varname) [ medianonly ]


varlist may include i.var to force a variable to be treated as categorical.

by(varname) is required and must have exactly two levels (numeric or string).

Option

medianonly — for continuous variables, skip normality checks and always report median (IQR) with Wilcoxon rank-sum p-value.

How variables are classified

Forced categorical: Any variable written as i.var in varlist.

Automatic categorical:

Numeric variables with exactly 2 distinct values (binomial).

If values are 0/1, output collapses to a single “=1” row.

String variables (any number of levels) are treated as categorical (cannot be analyzed as continuous).

Continuous: Numeric variables with >2 levels and not written as i.var.

Percentages for categorical rows use variable-specific, non-missing denominators within each group and overall.

Statistics performed

Categorical

2×2 → Fisher’s exact (p), otherwise Pearson chi-square.

P-values formatted as <0.001, 0.abc, or 0.ab.

Continuous (default behavior)

Shapiro–Wilk in each group; if both normal → t-test (reports mean±SD).

Otherwise → Wilcoxon rank-sum (reports median [p25–p75]).

Continuous with medianonly

Always Wilcoxon rank-sum, always report median (IQR).

Examples

Mixed table; categorical via i., continuous auto (t-test vs ranksum):

export_table i.sex age i.uceis_score bmi, by(group)


Always report medians (no normality testing):

export_table i.sex age i.uceis_score bmi, by(group) medianonly


Auto binomial detection (no i. needed):

export_table infection_minor infection_major outcome_seriousinfection, by(steroidfailure)

Header labels & row labels

If by() (or its encoded version) has value labels, those are used as group headers.

Variable labels are used as row names; for multi-level categoricals, the level label/value is appended.

Notes & limitations

Requires Stata 14+ (putexcel).

Variables with all values missing are skipped.

Strings with many levels are summarized as categorical counts; consider recoding if needed.

The Excel workbook is opened with modify (not overwritten).

Troubleshooting

“by() must have exactly two levels”: Recode or select a two-level grouping variable.

Help not showing: Ensure export_table.sthlp is on your adopath (findfile export_table.sthlp), then help export_table.

Wrong counts/percentages: Confirm categories are truly 0/1 or specify i.var to force categorical handling.

Author

Prepared for internal analyses of the ASUC trial workflow. Use and adapt as needed.
