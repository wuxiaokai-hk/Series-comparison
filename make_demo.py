#!/usr/bin/env python3
"""Demo: a bank stress model with NO geographies, where scenario B
also uses DIFFERENT equations from scenario A.

  python make_demo.py demo
  python scenario_explain.py report --a demo/base.csv --b demo/stress.csv \
         --eq-a demo/equations_base.txt --eq-b demo/equations_stress.txt -o demo_report.html

Exogenous variables (no equation in either file): UNEMP, HPI, RATE, GDPG.
Differences between the equation files:
  PD        changed: more sensitive to unemployment in the stress equations
  RWA       changed: risk weights rise with PD in the stress equations
  OPEX      equation only in base; in the stress scenario management sets costs directly
  DEPOSITS  equation only in stress; in base it is a given path
Scenario B also puts a judgmental add-factor on PD in 2026.
"""
import csv
import math
import os
import random
import sys

OUT = sys.argv[1] if len(sys.argv) > 1 else 'demo'
os.makedirs(OUT, exist_ok=True)
random.seed(5)
per = ['%dQ%d' % (y, q) for y in range(2019, 2031) for q in range(1, 5)]
T = len(per)
F = per.index('2026Q1')

EQ_BASE = """' Bank stress model: base-case equations
PD = 0.004 + 0.0015*UNEMP - 0.0004*@PCY(HPI)
LGD = 0.25 - 0.004*@PCY(HPI)
DLOG(LOANS) = 0.004 + 0.002*GDPG
ECL = PD*LGD*LOANS
PROV = ECL - ECL(-1)
NII = LOANS*(0.018 + 0.25*RATE/100)
OPEX = 0.009*LOANS
NETINC = NII - PROV - OPEX
CAP = CAP(-1) + 0.7*NETINC
RWA = 0.55*LOANS
CET1 = 100*CAP/RWA
"""
EQ_STRESS = """' Bank stress model: stress equations (modeller re-specified PD and RWA)
PD = 0.004 + 0.0025*UNEMP - 0.0004*@PCY(HPI) + [ar(1)=0.4, estsmpl="2008q1 2019q4"]
LGD = 0.25 - 0.004*@PCY(HPI)
DLOG(LOANS) = 0.004 + 0.002*GDPG
ECL = PD*LGD*LOANS
PROV = ECL - ECL(-1)
NII = LOANS*(0.018 + 0.25*RATE/100)
NETINC = NII - PROV - OPEX
CAP = CAP(-1) + 0.7*NETINC
RWA = 0.55*LOANS*(1 + 4*PD)
CET1 = 100*CAP/RWA
DEPOSITS = 0.92*LOANS
"""
DESC = {'UNEMP': 'Unemployment rate, %', 'HPI': 'House price index', 'RATE': 'Policy rate, %',
        'GDPG': 'Real GDP growth, % y/y', 'PD': 'Probability of default (12-month)', 'LGD': 'Loss given default',
        'LOANS': 'Loan book, $M', 'ECL': 'Expected credit loss allowance, $M', 'PROV': 'Provisions, $M',
        'NII': 'Net interest income, $M', 'OPEX': 'Operating expenses, $M', 'NETINC': 'Net income, $M',
        'CAP': 'CET1 capital, $M', 'RWA': 'Risk-weighted assets, $M', 'CET1': 'CET1 ratio, %',
        'DEPOSITS': 'Deposits, $M'}


def run(stress):
    unemp = [5.5 - 0.02 * t + (2.5 * math.sin(math.pi * min(1, (t - F + 1) / 8)) if stress and t >= F else 0)
             for t in range(T)]
    hpi = [100 * 1.008 ** t * ((1 - 0.12 * min(1, (t - F + 1) / 6)) if stress and t >= F else 1) for t in range(T)]
    rate = [2.5 + 0.5 * math.sin(t / 7) - (1.5 if stress and t >= F else 0) for t in range(T)]
    gdpg = [2.0 + 0.3 * math.sin(t / 5) - (3.0 * math.exp(-(t - F) / 4) if stress and t >= F else 0) for t in range(T)]
    v = {k: [0.0] * T for k in DESC}
    v.update(UNEMP=unemp, HPI=hpi, RATE=rate, GDPG=gdpg)
    v['LOANS'][0], v['CAP'][0], v['ECL'][0] = 5000.0, 600.0, 15.0
    u_pd = 0.0
    for t in range(T):
        pchy = (hpi[t] / hpi[t - 4] - 1) * 100 if t >= 4 else 3.2
        e = random.gauss(0, 0.0002)
        if stress:
            u_pd = 0.4 * u_pd + e + (0.002 if per[t].startswith('2026') else 0)   # AR(1) error + add-factor in 2026
            v['PD'][t] = 0.004 + 0.0025 * unemp[t] - 0.0004 * pchy + u_pd
        else:
            v['PD'][t] = 0.004 + 0.0015 * unemp[t] - 0.0004 * pchy + e
        v['LGD'][t] = 0.25 - 0.004 * pchy
        if t:
            v['LOANS'][t] = v['LOANS'][t - 1] * math.exp(0.004 + 0.002 * gdpg[t])
        v['ECL'][t] = v['PD'][t] * v['LGD'][t] * v['LOANS'][t]
        v['PROV'][t] = v['ECL'][t] - (v['ECL'][t - 1] if t else v['ECL'][t])
        v['NII'][t] = v['LOANS'][t] * (0.018 + 0.25 * rate[t] / 100)
        v['OPEX'][t] = 0.009 * v['LOANS'][t] * ((0.94 if t >= F else 1) if stress else 1)   # cost cuts, set directly
        v['NETINC'][t] = v['NII'][t] - v['PROV'][t] - v['OPEX'][t]
        if t:
            v['CAP'][t] = v['CAP'][t - 1] + 0.7 * v['NETINC'][t]
        v['RWA'][t] = 0.55 * v['LOANS'][t] * ((1 + 4 * v['PD'][t]) if stress else 1)
        v['CET1'][t] = 100 * v['CAP'][t] / v['RWA'][t]
        v['DEPOSITS'][t] = 0.92 * v['LOANS'][t] if stress else 4600 * 1.009 ** t
    return v


for name, stress in [('base.csv', False), ('stress.csv', True)]:
    v = run(stress)
    with open(os.path.join(OUT, name), 'w', newline='') as f:
        w = csv.writer(f)
        w.writerow(['Variable', 'Description'] + per)
        for k in DESC:
            w.writerow([k, DESC[k]] + ['%.8g' % x for x in v[k]])
with open(os.path.join(OUT, 'equations_base.txt'), 'w') as f:
    f.write(EQ_BASE)
with open(os.path.join(OUT, 'equations_stress.txt'), 'w') as f:
    f.write(EQ_STRESS)
print('generic demo written to', OUT)
