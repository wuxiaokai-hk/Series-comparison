# scenario_explain — explain why two scenarios of an equation model differ

A standard-library-only Python tool (no pip install). Give it the results of two scenarios (A and B) and the model's equations; it tells you what differs, why each variable moved, and which shocks matter most.

Works with any model written as `Y = formula` in EViews syntax:
- with or without geographies (geography tabs appear only if names carry geography codes);
- with the **same or different equations** in A and B (`--eq` or `--eq-a`/`--eq-b`);
- variables that no equation computes are labelled **potential exogenous** and treated as set by the scenario.

A companion tool, `scenario_compare`, adds features for large national-plus-regional macro models (for example vendor scenario models with 50 state blocks).

## Quick start

```bash
# interactive report, one self-contained .html
python scenario_explain.py report --a base.csv --b stress.csv \
       --eq-a equations_base.txt --eq-b equations_stress.txt -o explain.html

# same equations in both scenarios
python scenario_explain.py report --a base.csv --b stress.csv --eq model.txt

# report in your browser, plus exact "impact X -> Y" and "split gap across shocks" on demand
python scenario_explain.py serve  --a base.csv --b stress.csv --eq-a eq_base.txt --eq-b eq_stress.txt

# terminal
python scenario_explain.py top      --a base.csv --b stress.csv -n 40
python scenario_explain.py explain  CET1 --a base.csv --b stress.csv --eq model.txt
python scenario_explain.py impact   UNEMP CET1 --a base.csv --b stress.csv --eq model.txt
python scenario_explain.py attribute CET1 --a base.csv --b stress.csv --eq model.txt
python scenario_explain.py importance --a base.csv --b stress.csv --eq model.txt --csv importance.csv
python scenario_explain.py check    --a base.csv --b stress.csv --eq model.txt -o not_evaluable.csv
```

## Inputs

**Scenario data**: `.csv`, `.tsv` or `.xlsx` (not old `.xls`). The layout is detected automatically:
- series in rows (`Mnemonic, Description, 2000Q1, 2000Q2, …`)
- series in columns with dates down the first column, including Data Buffet-style header rows (`Description:`, `Mnemonic:`, `Source:` …)
- long format (`mnemonic, date, value`)

Date styles accepted: `2026Q1`, `Q1 2026`, `2026:1`, `2026M03`, `2026-03-31`, `3/31/2026`, `Mar-2026`, `2026`, and Excel date serials. Use `--a f1.csv --a f2.csv` to combine several files into one scenario. Use `--sheet` to pick an xlsx sheet and `--meta` to add a CSV of descriptions.

**Equations** (`--eq`, or `--eq-a` and `--eq-b` when they differ): text in EViews syntax: a file, several files, or a folder (`.txt .prg .eq .mod`). The parser reads:
- `Y = …`, `@IDENTITY Y = …`, `series Y = …`, equations that wrap across lines
- `LOG, EXP, ABS, SQR, D(x[,n[,s]]), DLOG, @PCH, @PC, @PCHY, @PCY, @PCA, @MOVAV, @MOVSUM, @TREND`, lags `X(-1)`
- any left-hand side, e.g. `DLOG(Y)` or `LOG(Y/P)`
- spec lines like `eq_y.ls dlog(y) c dlog(x)` (dependencies only)

Equations need **substituted coefficients** (numbers, not `C(1)`) to be decomposed. In EViews, an equation's *Representations* view shows a "Substituted Coefficients" line, and a model's source/text view lists its equations. If your model text only shows links (`:EQ_GDP`), export the substituted form of the equations. Lines with `C(i)` still count for tracing, and `info` tells you how many there are.


**Add-factor series:** by default none are assumed; each equation's add-factor is its residual (left-hand side minus formula). If your model keeps add-factors as separate series with a common suffix, name it: `--af-suffix _A` (comma-separate several).

## What's in the report
- **Summary** (first tab): the transmission map can be zoomed (scroll or +/−) and moved (drag). A written summary of how B differs from A, plus:
  - the scenario's shocks, ranked by how much of the key variables they explain;
  - the **key variables**: the changed variables most others depend on, one per concept (or your own with `--key CET1,PD,…`);
  - the share of each key variable's gap coming from each shock, with a quarter slider;
  - a **transmission map**: shocks on the left, key variables to the right, in the order the shocks reach them, coloured by their gap. Line thickness = share of the receiving variable's gap; the slider changes the quarter.
- **Impact X → Y**: a quarter slider, a chart of the effect and of its share of Y's gap over time, and the variables in between at the chosen quarter. With `serve` this works for any X and Y; in the static file it works for scenario shocks → key variables.

- **Key changes**: the complete list of series that differ. You can filter it, page through all of it, and download it as a CSV. Each series is tagged *changed directly* (an assumption or add-factor changed) or *via formula only* (same equation, same add-factor; it moved only because its inputs moved).
- **Formula vs add-factor**: the formula/add-factor split for every equation. A time slider picks any quarter, or choose the whole window.
- **By concept / By geography**: count only *directly changed* series as changed. Formula-only movers are shown in a separate column. Click a row to open its full list.
- **Layers**: how many layers of variables each variable affects (longest chain of equations), inputs into an area, and a layer map. See below.
- **Variable detail**: a stacked chart of each input's contribution to the gap in every quarter, a chart of the formula vs add-factor share over time, and a quarter slider for the exact numbers.
- **Equations A vs B** (when they differ): every changed, added or dropped equation, with the changed parts highlighted.

## Reading the numbers

- **Window**: from the first period where any series differs to the end of the data (override with `--start/--end`). Contributions are shown at the window end by default (`--at 2027Q4`, or the period selector in the report).
- **Contributions** are in the variable's own units and add up to the gap, apart from a small *interaction* term from non-linearity.
- **Add-factor** = the effect of switching the equation's residual (left-hand side minus formula) from A to B. When `--af-suffix` is set and a matching series exists, it is named and its change is shown. Residual changes that are only rounding noise are ignored.
- **Impact / attribute** hold the add-factors of every other equation at their A values and re-solve the equations period by period. Same-period loops are iterated to convergence. Variables whose equation can't be evaluated are held fixed and listed.
- **Channel share** in `path` = product of the link shares. It is a rough guide to which channel matters; `impact` gives the exact number.


## Which shocks matter most (Importance)

`python scenario_explain.py importance --a … --b … --eq … --csv importance.csv` (also the *Importance* tab of the report).

**Shocks** are the changed exogenous variables (no equation, or listed in `--exog`) and the changed add-factors. Each shock is switched from A to B on its own and everything it feeds is re-solved. For each variable y that moved, its share of y's gap is computed over the whole window:

    share = Σ_t effect_t · gap_t / Σ_t gap_t²

For each variable, the shares of all shocks add up to about 100% (exactly 100% when the model is linear). **Importance** is the average share across variables:
- `--weight concept` (default): each concept counts once, so 50 state copies of one variable don't dominate
- `--weight all`: every variable counts equally
- `--weight key`: only the key variables from the Summary tab (or `--key …`)

The output also gives how many variables a shock explains ≥10% / ≥50% of, and which variables it explains most. The CSV comes with `importance_matrix.csv`, listing every variable × shock share ≥1%.

**Two methods:**
- `--method fast` (default): traces each shock through the model using each equation's own split of its gap into inputs and add-factor (the same split as *Where the gap comes from*). `S(y,s) = Σ_k w_k(y)·S(k,s)`, with the model's loops solved by iteration. One pass over the model whatever the number of shocks (seconds even on 20,000 equations). It is exact for a linear model; on the demo, importance scores are within 1–2 points of the exact method and the ranking is the same.
- `--method exact`: re-solves the model once per shock. Use it to confirm the top few: `--method exact --max-shocks 20`.

**One geography only:** `importance --geo <code>` (or the *Variables to explain* selector in the report) averages each shock's share only over that geography's variables. It answers "what drives this region's forecast?". National or other-country shocks still rank if they move that geography. `--match REGEX` limits to variables by name or description instead.

The report uses the fast method; `--importance-method exact` switches it, and `--no-importance` skips the tab.

**Importance as a report:** add `-o` to `importance` to also write the full HTML report, opened on the Importance tab with your `--geo` and `--weight` already selected (same numbers as the terminal; nothing is computed twice):

```bash
python scenario_explain.py importance --a base.csv --b s4.csv --eq model.txt --geo ICAN -o s4.html
```


## How many layers a variable affects (Layers tab)

**Layers** is the longest chain of equations from a variable downstream. 0 means no equation uses it. 3 means it feeds a variable, which feeds a variable, which feeds a variable. Lags count as links. A loop (GDP → income → consumption → GDP) counts as **one** layer, because inside a loop the longest chain never ends. This describes the model's structure, not the size of an effect.

The tab has:
- a histogram of how many variables reach each number of layers;
- a table with each variable's layers and its full longest chain;
- with an area (`--target` or `--area`), a second number per variable: the longest chain that **ends in the area**, e.g. `--area ICAN`;
- **Inputs into the area**: the variables outside it that its equations use directly, and which area variables each one feeds. Everything else outside reaches the area only through these;
- **What the area feeds outside**: the variables outside it whose equations use an area variable directly, which area variables each uses, and its gap. One step only, not traced further. With `--target`, these are listed even though they're dropped from the rest of the report.
- **Maps of the model**: three maps, chosen one of four ways:
  - *Groups joined by key variables* (default). The model is split into groups of variables that are densely linked inside and loosely linked to the rest (Louvain community detection on the equation links). A group's **key variables** are the fewest variables that every link between it and the rest of the model passes through: a minimum vertex cover of its boundary links (König's theorem). A group joined through one or two key variables is almost a separate model hanging off them. Map 1 shows each group as one box, plus only the key variables between groups. Groups that hang off exactly the same keys (e.g. 40 states joined through the same 4 national variables) are bundled. Click a group to open it in Map 2; a large group opens as sub-groups, which open in Map 3. Terminal: `python scenario_explain.py groups … --max-keys 2`.
  - *Main chains, model structure* (main path analysis; Hummon & Doreian 1989). For a link u → v, the search path count SPC = N⁻(u) × N⁺(v): the number of chains from any input (no equation) to u, times the number from v to any end variable. It counts the input-to-end chains that run through the link, so hubs that many chains pass through score high, and a variable at the top of one long chain does not. Chain 1 starts at the link with the highest SPC and extends backward and forward along the highest-SPC links. Each further chain starts from the best unused link whose starting variable hasn't started a chain yet. Same for every scenario.
  - *Main chains, this scenario's gap*. Link weight = s(u→v) × N(v): s is the share of v's gap that comes from u, and N(v) = 1 + Σ s(v→w) N(w) is the share-weighted number of changed variables whose gap comes through v. It has no units, so rates and levels compare fairly. Chains are where the B − A difference actually travels.
  - *Longest chain (layers)*: Map 1 = 5+ layers, Map 2 = 3–4 fed by Map 1, Map 3 = the rest. Click a box to narrow the next map.

  Map 1 shows chain 1, Map 2 chains 2–4, Map 3 chains 5–12 (all adjustable). Each chain is listed under its map, and dashed boxes were already in a map above. Loops count as one step. Double-click a box for its details; scroll to zoom, drag to move.

With `--target`, layers are counted on the full model before variables are dropped, so the overall numbers stay true. With `--target-exog`, only the target's own equations count as links: variables outside it are exogenous, so links among them (e.g. oil → US CPI → US rates) are ignored.

```bash
python scenario_explain.py layers --a base.csv --b s4.csv --eq model.txt --area ICAN            # histogram, top chains, inputs into ICAN
python scenario_explain.py layers --a base.csv --b s4.csv --eq model.txt --area ICAN --by-area  # rank by chains ending in ICAN
```

## Equations that cannot be used

`python scenario_explain.py check --a … --b … --eq … -o not_evaluable.csv` writes one row per equation the tool cannot use, with the reason:
- coefficients not substituted (`C(1)`)
- specification only (`ls y c x`)
- could not parse
- unsupported function (the function is named)
- input series missing from the data (the series is named)
- equation gives missing values (e.g. log of a negative number)
- missing values in some quarters (a warning; the equation is still used)

`report` writes the same file automatically, next to the report, as `<report>_not_evaluable.csv`.

Functions the tool understands: all EViews functions that can appear in a series or model equation (about 170, taken from the EViews Command Reference), including:
- **math:** `LOG, EXP, ABS, SQRT, LOG10, @LOGX, @INV, @FACT, @EXPM1, @LOG1P, @SIGN, @LOGIT, @GAMMA, @MOD, @POW, @ROUND/@FLOOR/@CEILING(x[,n])`
- **trigonometry:** `@SIN, @COS, …`
- **distributions:** `@CNORM, @DNORM, @QNORM, @CLOGISTIC, @ERF, …`
- **logic and NA handling:** `@RECODE, @IIF/@IFF, @NAN, @ISNA, @EQNA, @NEQNA, @BOUNDS, @BETWEEN, @INLIST`, comparisons and `AND/OR`
- **time series:** `D, DLOG, @LAG, @PC/@PCH/@PCA/@PCHA/@PCY/@PCHY, @BRIDGE, @YTD, @DEMEAN, @DETREND`
- **dates:** `@TREND, @TRENDC, @TRENDBR, @OBSNUM, @DURING, @ISPERIOD, @BEFORE, @AFTER, @SEAS, @YEAR, @QUARTER, @MONTH, @ELEM`
- **all moving statistics:** `@MOV*` (NAs propagate) and `@M*` (NAs skipped): sum, average, centred average, stdev, var, cov, cor, max, min, median, …
- **all cumulative statistics:** `@CUM*` and `@CUMB*`, with an optional sample
- **sample statistics:** `@MEAN, @SUM, @STDEV, @MAX, @OBS, @QUANTILE, @FIRST, @LAST, …(x[,"sample"])`

**Sharing a failing equation without its names:** `python scenario_explain.py mask --eq model.txt --failed` prints the equation lines that cannot be used with every series name replaced by `x` and every number by `9`. EViews functions (`@MOVAV` …) and keywords (`AR`, `UNCOND`, `DLOG` …) stay visible, and hidden or unusual characters are spelled out, e.g. `<U+200B zero width space>`. Without `--failed` it shows all lines with `[...]` or AR/MA terms. `--match REGEX` picks other lines.

String, matrix and program functions are not included, because they can't appear in a model equation.

## Different equations in A and B

The two equation sets are compared variable by variable. Spacing, upper/lower case and `[ar(1)=…]` options are ignored:

| Status | Meaning | How its gap is split |
|---|---|---|
| changed | the formula differs | inputs + **equation changed**. Switching to B's equation, with its add-factor, is one effect. Every effect is averaged over A's and B's equation, so an input that only one formula uses still gets its share. |
| only in A | B sets the variable directly | inputs (A's formula) + **equation dropped in B** |
| only in B | A sets the variable directly | inputs (B's formula) + **equation added in B** |

These switches behave like shocks everywhere: Where B starts, `attribute`, Importance and the Summary map.

**Potential exogenous:** any variable that no equation computes is labelled *potential exogenous* and treated as set directly by the scenario. Equations that cannot be evaluated are treated the same way, and listed in `<report>_not_evaluable.csv`.

## Variables the modeller set exogenously

If the modeller overrode a variable that has an equation (exogenized it), list it with `--exog`:

```bash
python scenario_explain.py report --a base.csv --b stress.csv --eq model.txt --exog exog_list.txt
python scenario_explain.py impact UNEMP PD ... --exog 'HPI,RATE'
```

The list can be a file (names separated by spaces, commas or new lines, `.txt/.csv/.xlsx`) or a comma-separated list.

Listed variables are treated as **assumptions**:
- their whole gap is a scenario choice, so no formula breakdown is shown;
- they are never re-solved, so shocks to their inputs stop there;
- `attribute` counts them as root shocks;
- the report shows their equation as *switched off*.

Without `--exog`, an exogenized variable shows up as *mostly add-factor*. The totals are still right, but the formula split and the `impact` / `influence` results would assume the equation is active.

## Options worth knowing

| option | default | what |
|---|---|---|
| `--target SPEC` | | keep only these variables and everything that feeds them (directly or through any chain of equations, lags included). A geography code (`ICAN`), a wildcard (`"*_ICAN"`), names or a comma list; repeatable. Results for the kept variables are unchanged, just faster |
| `--target-exog` | off | with `--target`: treat every variable outside the target as exogenous. Only the target's own equations are solved; the outside variables they use (e.g. US GDP for `ICAN`) are taken from each scenario's data and become the shocks. Without it, gaps are traced back to the model's root shocks |
| `--area SPEC` | `--target` | area for the Layers tab: chains that end in these variables (geography code, wildcard or names) |
| `--tol` | 1e-6 | relative tolerance before a series counts as different |
| `--geo-sep` | auto | character before the geography code: `.`, `_`, or auto-detect. A name ending counts as a geography only if 3+ concepts share it |
| `--exog LIST` | | variables set exogenously by the modeller (file or comma list) |
| `--af-suffix` | none | suffix of add-factor series, if your model keeps them as series (e.g. `_A`); comma-separate several |
| `--adj-threshold` | 0.25 | add-factor share that tags a variable "mostly add-factor" |
| `--embed` | auto | series with full charts in the HTML: a number, `all`, or auto (all if ≤5,000 changed, else top 1,500). Every other variable still gets its per-quarter shares |
| `--focus VAR` | | always include VAR in full detail. A name, a comma list, or a wildcard like `"*_ICAN"` (all changed Canada series); repeatable |
| `--key LIST` | auto | key variables for the Summary tab (file or comma list) |
| `--summary-n` | 8 | how many key variables to pick automatically (0 turns the Summary off) |
| `-o` (importance) | | also write the full report, opened on the Importance tab |
| `--influence-n` | 20 | scenario shocks simulated for the Influence tab  |

## Demo

`python make_demo.py demo` builds a small bank stress model (no geographies): `base.csv`, `stress.csv`, `equations_base.txt`, `equations_stress.txt`. In it, PD and RWA are re-specified in the stress scenario, OPEX has an equation only in base and DEPOSITS only in stress, and PD has an AR(1) term and an add-factor in 2026.

```bash
python make_demo.py demo
python scenario_explain.py report --a demo/base.csv --b demo/stress.csv \
       --eq-a demo/equations_base.txt --eq-b demo/equations_stress.txt -o demo_report.html
```

`demo_report.html` is the result.
