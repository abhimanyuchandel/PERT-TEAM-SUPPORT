from __future__ import annotations

import json
import re
import subprocess
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path('/Users/abhichandel/Documents/Research/PE decision tool')
HTML_PATH = ROOT / 'index.html'
OUTPUT_PATH = ROOT / 'PE_tool_permutation_review.xlsx'


@dataclass
class Scenario:
    scenario_id: str
    group: str
    name: str
    summary: str
    overrides: dict[str, Any] | None = None
    initial_pending: bool = False


def scenario(scenario_id: str, group: str, name: str, summary: str, initial_pending: bool = False, **overrides: Any) -> Scenario:
    return Scenario(
        scenario_id=scenario_id,
        group=group,
        name=name,
        summary=summary,
        overrides=overrides or {},
        initial_pending=initial_pending,
    )


SCENARIOS: list[Scenario] = [
    scenario('INIT-00', 'Init', 'Blank first-load state', 'No clinical criteria entered; validates the pending landing state.', initial_pending=True),
    scenario('A1-01', 'A', 'Incidental subsegmental PE', 'Asymptomatic, incidental, subsegmental confirmed PE.', symptomatic='no', incidental=True, confirmedPe='confirmed', clotLocation='subsegmental', weightKg=80, crcl=90),
    scenario('A2-02', 'A', 'Incidental proximal PE', 'Asymptomatic, incidental, segmental/proximal confirmed PE.', symptomatic='no', incidental=True, confirmedPe='confirmed', clotLocation='segmental', weightKg=82, crcl=88),
    scenario('B1-03', 'B', 'Low-risk symptomatic subsegmental PE', 'Symptomatic subsegmental PE with low-risk score.', symptomatic='yes', clotLocation='subsegmental', spesi=0, patientAge=44, weightKg=78, crcl=95),
    scenario('B2-04', 'B', 'Low-risk symptomatic proximal PE', 'Symptomatic segmental/proximal PE with low-risk score.', symptomatic='yes', clotLocation='segmental', spesi=0, patientAge=52, weightKg=84, crcl=92),
    scenario('BC-05', 'Pending', 'Stable symptomatic PE without severity score', 'Symptomatic stable PE with no PESI/sPESI/Hestia entered.', symptomatic='yes', clotLocation='segmental', weightKg=80, crcl=90),
    scenario('B2-06', 'B', 'Low-risk PE with severe CKD', 'Low-risk proximal PE with CrCl <30 mL/min.', symptomatic='yes', clotLocation='segmental', spesi=0, patientAge=63, weightKg=84, crcl=22, severeCKD=True),
    scenario('B2-07', 'B', 'Low-risk pregnancy', 'Low-risk proximal PE in pregnancy.', symptomatic='yes', clotLocation='segmental', spesi=0, patientAge=31, weightKg=78, crcl=110, pregnancy=True),
    scenario('B2-08', 'B', 'Low-risk pregnancy with severe renal dysfunction', 'Low-risk proximal PE in pregnancy with CrCl <30 mL/min.', symptomatic='yes', clotLocation='segmental', spesi=0, patientAge=33, weightKg=79, crcl=24, severeCKD=True, pregnancy=True),
    scenario('B2-09', 'B', 'Low-risk pregnancy with severe renal dysfunction and weight >150 kg', 'Low-risk proximal PE in pregnancy with severe renal dysfunction and extreme weight.', symptomatic='yes', clotLocation='segmental', spesi=0, patientAge=35, weightKg=162, crcl=22, severeCKD=True, pregnancy=True),
    scenario('B2-10', 'B', 'Low-risk breastfeeding', 'Low-risk proximal PE while breastfeeding.', symptomatic='yes', clotLocation='segmental', spesi=0, patientAge=29, weightKg=76, crcl=105, breastfeeding=True),
    scenario('B2-11', 'B', 'Low-risk breastfeeding with severe renal dysfunction', 'Low-risk proximal PE while breastfeeding with CrCl <30 mL/min.', symptomatic='yes', clotLocation='segmental', spesi=0, patientAge=32, weightKg=80, crcl=20, severeCKD=True, breastfeeding=True),
    scenario('B2-12', 'B', 'Low-risk breastfeeding with weight >150 kg', 'Low-risk proximal PE while breastfeeding with extreme weight.', symptomatic='yes', clotLocation='segmental', spesi=0, patientAge=34, weightKg=158, crcl=95, breastfeeding=True),
    scenario('B2-13', 'B', 'Low-risk breastfeeding with severe renal dysfunction and weight >150 kg', 'Low-risk proximal PE while breastfeeding with severe renal dysfunction and extreme weight.', symptomatic='yes', clotLocation='segmental', spesi=0, patientAge=36, weightKg=168, crcl=19, severeCKD=True, breastfeeding=True),
    scenario('B2-14', 'B', 'Low-risk PE with APS', 'Low-risk proximal PE with thrombotic APS.', symptomatic='yes', clotLocation='segmental', spesi=0, patientAge=46, weightKg=72, crcl=88, aps=True),
    scenario('C1-15', 'C', 'Category C1 elevated score only', 'Elevated clinical severity score without RV dysfunction or biomarkers.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=70, weightKg=83, crcl=76),
    scenario('C1R-16', 'C', 'Category C1R with supplemental oxygen', 'Elevated clinical severity score and low-flow oxygen need.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=68, weightKg=81, crcl=74, oxygenSupport='o2-low', oxygenSat=92, rr=24),
    scenario('C2-17', 'C', 'Category C2 with RV dysfunction', 'Elevated clinical severity score and RV dysfunction.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=73, weightKg=86, crcl=78, rvDysfunction='yes'),
    scenario('C2-18', 'C', 'Category C2 with biomarker elevation', 'Elevated clinical severity score and positive troponin.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=74, weightKg=85, crcl=70, troponin='yes'),
    scenario('C2-19', 'C', 'Suspected C2 with delayed imaging', 'Suspected PE, C2-risk physiology, low bleeding risk, imaging delayed.', symptomatic='yes', confirmedPe='suspected', imagingDelayed=True, clotLocation='segmental', spesi=1, patientAge=69, weightKg=82, crcl=87, rvDysfunction='yes'),
    scenario('C2-20', 'C', 'C2 with active bleeding', 'Category C2 physiology with active bleeding selected.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=66, weightKg=80, crcl=79, rvDysfunction='yes', bleedAbsoluteActive=True),
    scenario('C3-21', 'C', 'Category C3 intermediate-high risk', 'Elevated score with RV dysfunction and biomarker elevation.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=72, weightKg=88, crcl=75, rvDysfunction='yes', troponin='yes'),
    scenario('C3-22', 'C', 'C3 with score-derived tachycardia alert', 'Normotensive C3 profile with manual score HR >110 bpm.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=61, weightKg=90, crcl=84, rvDysfunction='yes', troponin='yes', scoreHr=126, map=74),
    scenario('C3-23', 'C', 'C3 with clot-in-transit', 'Intermediate-high-risk PE with right-heart clot-in-transit.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=64, weightKg=89, crcl=82, rvDysfunction='yes', troponin='yes', clotTransit=True),
    scenario('C3-24', 'C', 'C3 recurrent PE on therapy', 'Intermediate-high-risk PE with recurrent PE despite therapeutic anticoagulation.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=62, weightKg=87, crcl=80, rvDysfunction='yes', troponin='yes', recurrentOnTherapy=True),
    scenario('C3-25', 'C', 'C3 with severe renal dysfunction and weight >150 kg', 'Intermediate-high-risk PE with severe renal dysfunction and extreme weight.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=67, weightKg=156, crcl=24, severeCKD=True, rvDysfunction='yes', troponin='yes'),
    scenario('C3-26', 'C', 'C3 with active cancer and recurrent PE', 'Intermediate-high-risk PE, active cancer, recurrent PE on therapeutic anticoagulation.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=68, weightKg=83, crcl=73, rvDysfunction='yes', troponin='yes', activeCancer=True, recurrentOnTherapy=True),
    scenario('D1-27', 'D', 'Category D1 transient hypotension', 'Transient hypotension without isolated hypoperfusion markers.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=71, weightKg=86, crcl=72, transientHypotension=True),
    scenario('D1-28', 'D', 'D1 with relative thrombolysis contraindication', 'Transient hypotension plus recent major surgery.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=73, weightKg=84, crcl=70, transientHypotension=True, bleedRelativeRecentSurgery=True),
    scenario('D1-29', 'D', 'D1 with absolute thrombolysis contraindication', 'Transient hypotension plus prior intracranial hemorrhage.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=75, weightKg=80, crcl=68, transientHypotension=True, bleedAbsoluteICh=True),
    scenario('D2-30', 'D', 'Category D2 normotensive shock by lactate', 'Normotensive shock from lactate >2 mmol/L.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=70, weightKg=85, crcl=78, lactate=3.4),
    scenario('D2-31', 'D', 'D2 subsegmental disproportionate presentation', 'Normotensive shock physiology with subsegmental-only clot burden.', symptomatic='yes', clotLocation='subsegmental', spesi=1, patientAge=69, weightKg=81, crcl=76, lactate=3.1),
    scenario('D2-32', 'D', 'D2 unable to lie flat', 'Normotensive shock physiology with inability to lie flat.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=72, weightKg=88, crcl=77, lactate=3.2, unableLieFlat=True),
    scenario('D2-33', 'D', 'D2 with severe renal dysfunction', 'Normotensive shock physiology with severe renal dysfunction.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=74, weightKg=82, crcl=21, severeCKD=True, lactate=3.6),
    scenario('E1-34', 'E', 'Category E1 shock / single vasopressor', 'Persistent hypotension/shock with one vasopressor.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=71, weightKg=86, crcl=72, persistentHypotension=True, vasopressors='1', map=58),
    scenario('E1-35', 'E', 'E1 with relative thrombolysis contraindication', 'Persistent hypotension with recent major surgery.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=73, weightKg=84, crcl=70, persistentHypotension=True, vasopressors='1', map=56, bleedRelativeRecentSurgery=True),
    scenario('E1-36', 'E', 'E1 with absolute thrombolysis contraindication', 'Persistent hypotension with prior intracranial hemorrhage.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=76, weightKg=80, crcl=65, persistentHypotension=True, vasopressors='1', map=55, bleedAbsoluteICh=True),
    scenario('E1-37', 'E', 'E1 with clot-in-transit', 'Category E1 shock physiology with clot-in-transit.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=67, weightKg=88, crcl=74, persistentHypotension=True, vasopressors='1', map=57, clotTransit=True),
    scenario('E2-38', 'E', 'Category E2 refractory shock', 'Refractory shock with two vasopressors.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=70, weightKg=90, crcl=73, persistentHypotension=True, vasopressors='2plus', map=52),
    scenario('E2-39', 'E', 'E2 with relative thrombolysis contraindication', 'Refractory shock with recent major surgery.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=74, weightKg=87, crcl=68, persistentHypotension=True, vasopressors='2plus', map=51, bleedRelativeRecentSurgery=True),
    scenario('E2-40', 'E', 'E2 with absolute thrombolysis contraindication', 'Refractory shock with prior intracranial hemorrhage.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=77, weightKg=82, crcl=62, persistentHypotension=True, vasopressors='2plus', map=50, bleedAbsoluteICh=True),
    scenario('E2-41', 'E', 'E2 suspected PE with delayed imaging', 'Suspected PE, E2-shock physiology, imaging delayed.', symptomatic='yes', confirmedPe='suspected', imagingDelayed=True, clotLocation='segmental', spesi=1, patientAge=69, weightKg=89, crcl=75, persistentHypotension=True, vasopressors='2plus', map=50),
    scenario('E2-42', 'E', 'E2 with severe renal dysfunction', 'Refractory shock with severe renal dysfunction.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=72, weightKg=86, crcl=18, severeCKD=True, persistentHypotension=True, vasopressors='2plus', map=49),
    scenario('C2-43', 'C', 'C2 HI-PEITHO eligible by HR and RR', 'Confirmed C2 PE with at least two HI-PEITHO clinical severity features (HR >100 and RR >20).', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=66, weightKg=82, crcl=84, rvDysfunction='yes', scoreHr=108, rr=24),
    scenario('C3-44', 'C', 'C3 HI-PEITHO eligible by HR and SBP', 'Confirmed C3 PE with HI-PEITHO features HR >100 and score-derived SBP <110 mm Hg.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=67, weightKg=85, crcl=80, rvDysfunction='yes', troponin='yes', scoreHr=104, scoreMode='calc-spesi', calcSpesiSbp=105),
    scenario('D1-45', 'D', 'D1 HI-PEITHO eligible by transient hypotension and tachypnea', 'Confirmed D1 PE with transient hypotension plus RR >20.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=70, weightKg=84, crcl=78, transientHypotension=True, rr=26),
    scenario('D2-46', 'D', 'D2 HI-PEITHO eligible by HR and HFNC', 'Confirmed D2 PE with lactate-defined shock physiology and HI-PEITHO features HR >100 plus HFNC.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=71, weightKg=87, crcl=76, lactate=3.1, scoreHr=109, oxygenSupport='hfnc'),
    scenario('C3-47', 'C', 'C3 HI-PEITHO features but thrombolysis risk selected', 'Confirmed C3 PE with HI-PEITHO clinical features but a relative bleeding-risk factor, so the HI-PEITHO recommendation should not appear.', symptomatic='yes', clotLocation='segmental', spesi=1, patientAge=68, weightKg=86, crcl=79, rvDysfunction='yes', troponin='yes', scoreHr=107, rr=24, bleedRelativeRecentSurgery=True),
]


INPUT_COLUMNS = [
    ('scenario_id', 'Scenario ID'),
    ('group', 'Group'),
    ('name', 'Scenario Name'),
    ('summary', 'Scenario Summary'),
    ('symptomatic', 'Symptomatic'),
    ('confirmedPe', 'Diagnosis Status'),
    ('incidental', 'Incidental'),
    ('clotLocation', 'Clot Location'),
    ('provokingFactor', 'Provoking Factor'),
    ('scoreMode', 'Score Mode'),
    ('pesi', 'PESI'),
    ('spesi', 'sPESI'),
    ('hestia', 'Hestia'),
    ('bova', 'Bova'),
    ('calcSpesiSbp', 'Score SBP'),
    ('patientAge', 'Age'),
    ('scoreHr', 'Score HR'),
    ('weightKg', 'Weight (kg)'),
    ('crcl', 'CrCl (mL/min)'),
    ('map', 'MAP'),
    ('lactate', 'Lactate'),
    ('vasopressors', 'Vasopressors'),
    ('persistentHypotension', 'Persistent Hypotension/Shock'),
    ('transientHypotension', 'Transient Hypotension'),
    ('cardiacArrest', 'Cardiac Arrest'),
    ('aki', 'AKI'),
    ('oliguria', 'Oliguria'),
    ('mentalStatus', 'Mental Status Change'),
    ('lowCardiacIndex', 'Low Cardiac Index'),
    ('shockScore', 'Shock Score Increased'),
    ('unableLieFlat', 'Unable to Lie Flat'),
    ('rvDysfunction', 'RV Dysfunction'),
    ('troponin', 'Troponin'),
    ('bnp', 'BNP'),
    ('oxygenSat', 'O2 Sat'),
    ('rr', 'Respiratory Rate'),
    ('oxygenSupport', 'Oxygen Support'),
    ('absoluteBleedingRisk', 'Absolute Bleeding Risk'),
    ('relativeBleedingRisk', 'Relative Bleeding Risk'),
    ('bleedAbsoluteActive', 'Active Bleeding'),
    ('bleedAbsoluteICh', 'Prior ICH'),
    ('bleedRelativeRecentSurgery', 'Recent Major Surgery'),
    ('bleedRelativeCoagulopathy', 'Coagulopathy'),
    ('bleedRelativeHypertension', 'Severe HTN'),
    ('pregnancy', 'Pregnancy'),
    ('breastfeeding', 'Breastfeeding'),
    ('activeCancer', 'Active Cancer'),
    ('aps', 'APS'),
    ('severeCKD', 'Severe CKD/Dialysis'),
    ('imagingDelayed', 'Imaging Delayed'),
    ('clotTransit', 'Clot in Transit'),
    ('recurrentOnTherapy', 'Recurrent on Therapy'),
]


REVIEW_COLUMNS = [
    'Scenario ID',
    'Group',
    'Scenario Name',
    'Scenario Summary',
    'Observed Category',
    'Observed Profile',
    'Overall Recommendation',
    'PERT Action Bundle',
    'Medication Recommendations',
    'Monitoring Recommendations',
    'Escalation / Backup',
    'Cautions / Modifiers',
    'Rationale From Inputs',
]


BASE_DATA = {
    'discussionTime': '2026-04-08T09:00',
    'symptomatic': 'yes',
    'confirmedPe': 'confirmed',
    'clotLocation': 'segmental',
    'provokingFactor': 'unknown',
    'weightKg': 80,
    'crcl': 90,
    'incidental': False,
    'imagingDelayed': False,
    'clotTransit': False,
    'recurrentOnTherapy': False,
    'scoreMode': 'manual',
    'pesi': None,
    'spesi': None,
    'hestia': None,
    'bova': None,
    'scoreHr': None,
    'calcSpesiSbp': None,
    'calcPesiSbp': None,
    'calcBovaSbp': None,
    'chosenScoreName': '',
    'chosenScoreValue': '',
    'scoreSummary': '',
    'scoreWarnings': [],
    'scoreReady': False,
    'patientAge': None,
    'map': None,
    'lactate': None,
    'vasopressors': '0',
    'persistentHypotension': False,
    'transientHypotension': False,
    'cardiacArrest': False,
    'renalHypoperfusion': False,
    'aki': False,
    'oliguria': False,
    'mentalStatus': False,
    'lowCardiacIndex': False,
    'shockScore': False,
    'unableLieFlat': False,
    'rvDysfunction': 'unknown',
    'troponin': 'unknown',
    'bnp': 'unknown',
    'oxygenSat': None,
    'rr': None,
    'oxygenSupport': 'room-air',
    'contraAnticoag': False,
    'bleedAbsoluteActive': False,
    'bleedAbsoluteICh': False,
    'bleedAbsoluteDissection': False,
    'bleedAbsoluteNeuroSurgery': False,
    'bleedAbsoluteIntracranial': False,
    'bleedRelativeCoagulopathy': False,
    'bleedRelativeHypertension': False,
    'bleedRelativeRecentSurgery': False,
    'bleedRelativeSeriousTrauma': False,
    'bleedRelativePregnancy': False,
    'bleedRelativeRecentGiBleed': False,
    'pregnancy': False,
    'breastfeeding': False,
    'activeCancer': False,
    'aps': False,
    'severeCKD': False,
    'caseNarrative': '',
    'inputValidationWarnings': [],
}


def extract_script(html: str) -> str:
    match = re.search(r'<script>([\s\S]*?)</script>', html)
    if not match:
        raise RuntimeError('Could not find <script> block in index.html')
    return match.group(1)


def build_jxa_runner(script_source: str) -> str:
    export_bridge = """
globalThis.__toolExports = {
  ABSOLUTE_BLEEDING_RISK_FIELDS,
  RELATIVE_BLEEDING_RISK_FIELDS,
  annotateStrength,
  categoryDescriptor,
  classify,
  buildPertConsensus,
  buildRecommendations,
  buildMedicationMonitoringPlan,
  buildPlanText,
  initialPendingClassification
};
"""
    script_literal = json.dumps(script_source + "\n" + export_bridge)
    return f'''
ObjC.import('Foundation');

function readFile(path) {{
  return ObjC.unwrap($.NSString.stringWithContentsOfFileEncodingError(path, $.NSUTF8StringEncoding, null));
}}

function writeFile(path, text) {{
  $(text).writeToFileAtomicallyEncodingError(path, true, $.NSUTF8StringEncoding, null);
}}

function makeClassList() {{
  return {{ add() {{}}, remove() {{}}, contains() {{ return false; }} }};
}}

function makeElement(id) {{
  return {{
    id: id || '',
    value: '',
    defaultValue: '',
    checked: false,
    defaultChecked: false,
    textContent: '',
    innerHTML: '',
    className: '',
    hidden: false,
    disabled: false,
    title: '',
    tagName: 'DIV',
    options: [],
    style: {{}},
    dataset: {{}},
    classList: makeClassList(),
    addEventListener() {{}},
    appendChild() {{}},
    removeAttribute(name) {{ delete this[name]; }},
    setAttribute(name, value) {{ this[name] = value; }},
    querySelectorAll() {{ return []; }},
    querySelector() {{ return null; }},
    closest() {{ return null; }},
    reset() {{}},
    select() {{}},
    focus() {{}},
    blur() {{}}
  }};
}}

var __elements = {{}};
var document = {{
  getElementById: function(id) {{
    if (!__elements[id]) __elements[id] = makeElement(id);
    return __elements[id];
  }},
  querySelectorAll: function() {{ return []; }},
  querySelector: function() {{ return null; }},
  createElement: function(tag) {{ const el = makeElement(''); el.tagName = String(tag || 'div').toUpperCase(); return el; }},
  execCommand: function() {{ return true; }}
}};
var window = {{
  SpeechRecognition: null,
  webkitSpeechRecognition: null,
  matchMedia: function() {{ return {{ matches: false }}; }}
}};
var navigator = {{
  userAgent: '',
  languages: ['en-US'],
  language: 'en-US',
  clipboard: {{ writeText: function() {{ return Promise.resolve(); }} }}
}};
function setTimeout() {{ return 0; }}
function clearTimeout() {{}}

globalThis.document = document;
globalThis.window = window;
globalThis.navigator = navigator;
globalThis.setTimeout = setTimeout;
globalThis.clearTimeout = clearTimeout;

eval({script_literal});
var tool = globalThis.__toolExports;
var ABSOLUTE_BLEEDING_RISK_FIELDS = tool.ABSOLUTE_BLEEDING_RISK_FIELDS;
var RELATIVE_BLEEDING_RISK_FIELDS = tool.RELATIVE_BLEEDING_RISK_FIELDS;
var annotateStrength = tool.annotateStrength;
var categoryDescriptor = tool.categoryDescriptor;
var classify = tool.classify;
var buildPertConsensus = tool.buildPertConsensus;
var buildRecommendations = tool.buildRecommendations;
var buildMedicationMonitoringPlan = tool.buildMedicationMonitoringPlan;
var buildPlanText = tool.buildPlanText;
var initialPendingClassification = tool.initialPendingClassification;

function deepClone(obj) {{
  return JSON.parse(JSON.stringify(obj));
}}

function defaultChosenScoreValue(data) {{
  const parts = [];
  if (data.pesi !== null) parts.push('PESI ' + data.pesi);
  if (data.spesi !== null) parts.push('sPESI ' + data.spesi);
  if (data.hestia !== null) parts.push('Hestia ' + data.hestia);
  if (data.bova !== null) parts.push('Bova ' + data.bova);
  if (data.scoreHr !== null) parts.push('score HR ' + data.scoreHr);
  return parts.length ? parts.join(', ') : 'No formal severity score entered';
}}

function normalizeScenarioData(base, overrides) {{
  const data = deepClone(base);
  Object.keys(overrides || {{}}).forEach((key) => {{
    data[key] = overrides[key];
  }});

  if (data.pregnancy || data.bleedRelativePregnancy) {{
    data.pregnancy = true;
    data.bleedRelativePregnancy = true;
    data.breastfeeding = false;
  }}
  if (data.breastfeeding) {{
    data.pregnancy = false;
    data.bleedRelativePregnancy = false;
  }}
  if (data.crcl !== null && data.crcl < 30) {{
    data.severeCKD = true;
  }}
  if (data.vasopressors !== '0') {{
    data.persistentHypotension = true;
  }}

  data.aki = !!(data.aki || data.renalHypoperfusion);
  data.oliguria = !!(data.oliguria || data.renalHypoperfusion);
  data.inputValidationWarnings = data.inputValidationWarnings || [];
  data.scoreWarnings = data.scoreWarnings || [];
  data.caseNarrative = data.caseNarrative || '';
  data.chosenScoreName = data.chosenScoreName || 'Manual review scenario';
  data.chosenScoreValue = data.chosenScoreValue || defaultChosenScoreValue(data);
  data.scoreSummary = data.scoreSummary || data.chosenScoreValue;
  data.scoreReady = data.pesi !== null || data.spesi !== null || data.hestia !== null || data.bova !== null;

  data.absoluteBleedingRisk = ABSOLUTE_BLEEDING_RISK_FIELDS.some((item) => !!data[item.key]);
  data.relativeBleedingRisk = RELATIVE_BLEEDING_RISK_FIELDS.some((item) => !!data[item.key]);
  data.highBleedingRisk = data.absoluteBleedingRisk || data.relativeBleedingRisk;
  data.contraAnticoag = !!(data.contraAnticoag || data.bleedAbsoluteActive);
  data.contraThrombolysis = !!data.absoluteBleedingRisk;

  return data;
}}

function joinLines(items) {{
  return (items || []).length ? items.join('\\n') : '';
}}

function run(argv) {{
  const scenarioPath = argv[0];
  const outputPath = argv[1];
  const scenarioPayload = JSON.parse(readFile(scenarioPath));
  const base = scenarioPayload.base;
  const scenarios = scenarioPayload.scenarios;

  const results = scenarios.map((scenario) => {{
    if (scenario.initial_pending) {{
      const data = normalizeScenarioData(base, {{}});
      const cls = initialPendingClassification();
      return {{
        scenario_id: scenario.scenario_id,
        group: scenario.group,
        name: scenario.name,
        summary: scenario.summary,
        category: cls.category,
        profile: categoryDescriptor(cls.base),
        overall: ['Overall recommendation pending until clinical criteria are entered.'],
        actions: [],
        medication: [],
        monitoring: [],
        backup: [],
        alerts: ['Awaiting clinical inputs.'],
        rationale: [],
        plan: 'Enter clinical data to generate the PERT documentation draft.',
        normalized_data: data
      }};
    }}

    const data = normalizeScenarioData(base, scenario.overrides || {{}});
    const cls = classify(data);
    const pert = buildPertConsensus(data, cls);
    const bundle = buildRecommendations(data, cls);
    const medicationPlan = buildMedicationMonitoringPlan(data, cls);
    const summaryWithStrength = annotateStrength(pert.summary, 'Operational');
    const actionsWithStrength = annotateStrength(pert.actions, 'Operational');
    const backupWithStrength = annotateStrength(pert.backup, 'Operational');
    const recommendationsWithStrength = annotateStrength(bundle.recommendations, 'Operational');
    const medicationWithStrength = annotateStrength(medicationPlan.medication, 'Operational');
    const monitoringWithStrength = annotateStrength(medicationPlan.monitoring, 'Operational');
    const notesWithStrength = annotateStrength(medicationPlan.notes, 'Operational');
    const plan = buildPlanText(
      data,
      cls,
      {{ summary: summaryWithStrength, actions: actionsWithStrength, backup: backupWithStrength }},
      recommendationsWithStrength,
      bundle.rationale,
      {{ ...medicationPlan, medication: medicationWithStrength, monitoring: monitoringWithStrength, notes: notesWithStrength }}
    );

    return {{
      scenario_id: scenario.scenario_id,
      group: scenario.group,
      name: scenario.name,
      summary: scenario.summary,
      category: cls.category,
      profile: categoryDescriptor(cls.base),
      overall: summaryWithStrength,
      actions: actionsWithStrength,
      medication: medicationWithStrength,
      monitoring: monitoringWithStrength,
      backup: backupWithStrength,
      alerts: bundle.alerts,
      rationale: bundle.rationale,
      plan: plan,
      normalized_data: data
    }};
  }});

  writeFile(outputPath, JSON.stringify(results, null, 2));
  return 'ok';
}}
'''


def column_letter(index: int) -> str:
    result = ''
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def sheet_xml(rows: list[list[Any]], widths: list[int] | None = None, freeze_header: bool = True) -> str:
    last_col = column_letter(max(len(row) for row in rows))
    last_row = len(rows)
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        f'<dimension ref="A1:{last_col}{last_row}"/>'
    ]
    if freeze_header:
        parts.append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    if widths:
        parts.append('<cols>')
        for idx, width in enumerate(widths, start=1):
            parts.append(f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>')
        parts.append('</cols>')
    parts.append('<sheetData>')
    for row_idx, row in enumerate(rows, start=1):
        parts.append(f'<row r="{row_idx}">')
        for col_idx, value in enumerate(row, start=1):
            if value is None:
                continue
            text = '' if value is None else str(value)
            style = '1' if row_idx == 1 else '2'
            ref = f'{column_letter(col_idx)}{row_idx}'
            parts.append(
                f'<c r="{ref}" t="inlineStr" s="{style}"><is><t xml:space="preserve">{escape(text)}</t></is></c>'
            )
        parts.append('</row>')
    parts.append('</sheetData>')
    parts.append(f'<autoFilter ref="A1:{last_col}{last_row}"/>')
    parts.append('<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    parts.append('</worksheet>')
    return ''.join(parts)


def workbook_xml(sheet_names: list[str]) -> str:
    sheets = ''.join(
        f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
        for i, name in enumerate(sheet_names, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="25000" windowHeight="14000"/></bookViews>'
        f'<sheets>{sheets}</sheets>'
        '</workbook>'
    )


def workbook_rels_xml(sheet_count: int) -> str:
    rels = ''.join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, sheet_count + 1)
    )
    rels += f'<Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{rels}'
        '</Relationships>'
    )


def root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )


def content_types_xml(sheet_count: int) -> str:
    overrides = ''.join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, sheet_count + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'{overrides}'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )


def styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font>'
        '</fonts>'
        '<fills count="2">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFE9F3EE"/><bgColor indexed="64"/></patternFill></fill>'
        '</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="top" wrapText="1"/></xf>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )


def build_key_input_summary(data: dict[str, Any]) -> str:
    parts = [
        f"diagnosis={data['confirmedPe']}",
        f"symptomatic={data['symptomatic']}",
        f"clot={data['clotLocation']}",
    ]
    score_bits = []
    for label, key in [('PESI', 'pesi'), ('sPESI', 'spesi'), ('Hestia', 'hestia'), ('Bova', 'bova')]:
        if data.get(key) is not None:
            score_bits.append(f"{label}={data[key]}")
    if score_bits:
        parts.append('scores=' + ', '.join(score_bits))
    hemo_bits = []
    for label, key in [('MAP', 'map'), ('lactate', 'lactate')]:
        if data.get(key) is not None:
            hemo_bits.append(f"{label}={data[key]}")
    if data.get('persistentHypotension'):
        hemo_bits.append('persistent shock')
    if data.get('transientHypotension'):
        hemo_bits.append('transient hypotension')
    if data.get('vasopressors') and data['vasopressors'] != '0':
        hemo_bits.append(f"vasopressors={data['vasopressors']}")
    if hemo_bits:
        parts.append('hemodynamics=' + ', '.join(hemo_bits))
    resp_bits = []
    if data.get('oxygenSupport') and data['oxygenSupport'] != 'room-air':
        resp_bits.append(data['oxygenSupport'])
    if data.get('oxygenSat') is not None:
        resp_bits.append(f"O2 sat={data['oxygenSat']}")
    if data.get('rr') is not None:
        resp_bits.append(f"RR={data['rr']}")
    if resp_bits:
        parts.append('resp=' + ', '.join(resp_bits))
    rv_bits = []
    for label, key in [('RV', 'rvDysfunction'), ('troponin', 'troponin'), ('BNP', 'bnp')]:
        if data.get(key) not in (None, 'unknown'):
            rv_bits.append(f"{label}={data[key]}")
    if rv_bits:
        parts.append('RV/biomarkers=' + ', '.join(rv_bits))
    bleed_bits = []
    if data.get('absoluteBleedingRisk'):
        bleed_bits.append('absolute')
    if data.get('relativeBleedingRisk'):
        bleed_bits.append('relative')
    if bleed_bits:
        parts.append('bleeding=' + ', '.join(bleed_bits))
    special_bits = []
    for label, key in [('pregnancy', 'pregnancy'), ('breastfeeding', 'breastfeeding'), ('cancer', 'activeCancer'), ('APS', 'aps'), ('severe CKD', 'severeCKD'), ('clot-in-transit', 'clotTransit'), ('recurrent PE', 'recurrentOnTherapy')]:
        if data.get(key):
            special_bits.append(label)
    if special_bits:
        parts.append('special=' + ', '.join(special_bits))
    return '; '.join(parts)


def yes_no(value: Any) -> str:
    if isinstance(value, bool):
        return 'yes' if value else 'no'
    if value is None:
        return ''
    return str(value)


def join_items(items: list[str]) -> str:
    return '\n'.join(f'• {item}' for item in items) if items else ''


def build_review_rows(results: list[dict[str, Any]]) -> list[list[Any]]:
    rows = [REVIEW_COLUMNS]
    for result in results:
        rows.append([
            result['scenario_id'],
            result['group'],
            result['name'],
            result['summary'],
            result['category'],
            result['profile'],
            join_items(result['overall']),
            join_items(result['actions']),
            join_items(result['medication']),
            join_items(result['monitoring']),
            join_items(result['backup']),
            join_items(result['alerts']),
            join_items(result['rationale']),
        ])
    return rows


def build_input_rows(results: list[dict[str, Any]]) -> list[list[Any]]:
    rows = [[label for _, label in INPUT_COLUMNS] + ['Key Input Summary']]
    for result in results:
        data = result['normalized_data']
        row: list[Any] = []
        for key, _label in INPUT_COLUMNS:
            if key in {'scenario_id', 'group', 'name', 'summary'}:
                row.append(result[key])
            else:
                row.append(yes_no(data.get(key)))
        row.append(build_key_input_summary(data))
        rows.append(row)
    return rows


def build_metadata_rows(results: list[dict[str, Any]]) -> list[list[Any]]:
    return [
        ['Field', 'Value'],
        ['Generated', datetime.now().isoformat(timespec='seconds')],
        ['Source HTML', str(HTML_PATH)],
        ['Workbook', str(OUTPUT_PATH)],
        ['Scenario count', str(len(results))],
        ['Method', 'Curated representative scenario matrix evaluated against the current index.html decision logic.'],
        ['Note', 'This workbook is intended for guideline cross-checking and QA review; it is not exhaustive of all possible calculator states.'],
    ]


def write_workbook(path: Path, sheets: list[tuple[str, list[list[Any]], list[int]]]) -> None:
    with ZipFile(path, 'w', ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types_xml(len(sheets)))
        zf.writestr('_rels/.rels', root_rels_xml())
        zf.writestr('xl/workbook.xml', workbook_xml([name for name, _rows, _widths in sheets]))
        zf.writestr('xl/_rels/workbook.xml.rels', workbook_rels_xml(len(sheets)))
        zf.writestr('xl/styles.xml', styles_xml())
        for idx, (_name, rows, widths) in enumerate(sheets, start=1):
            zf.writestr(f'xl/worksheets/sheet{idx}.xml', sheet_xml(rows, widths))


def main() -> None:
    html = HTML_PATH.read_text(encoding='utf-8')
    script_source = extract_script(html)
    runner_source = build_jxa_runner(script_source)

    scenario_payload = {
        'base': BASE_DATA,
        'scenarios': [
            {
                'scenario_id': item.scenario_id,
                'group': item.group,
                'name': item.name,
                'summary': item.summary,
                'overrides': item.overrides,
                'initial_pending': item.initial_pending,
            }
            for item in SCENARIOS
        ],
    }

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        scenario_path = tmp / 'scenarios.json'
        output_json_path = tmp / 'results.json'
        runner_path = tmp / 'runner.js'
        scenario_path.write_text(json.dumps(scenario_payload), encoding='utf-8')
        runner_path.write_text(runner_source, encoding='utf-8')

        subprocess.run(
            ['osascript', '-l', 'JavaScript', str(runner_path), str(scenario_path), str(output_json_path)],
            check=True,
            cwd=ROOT,
        )

        results = json.loads(output_json_path.read_text(encoding='utf-8'))

    review_rows = build_review_rows(results)
    input_rows = build_input_rows(results)
    metadata_rows = build_metadata_rows(results)

    sheets = [
        ('Scenario Review', review_rows, [14, 10, 28, 36, 16, 28, 70, 70, 70, 70, 60, 40, 60]),
        ('Scenario Inputs', input_rows, [14, 10, 28, 36] + [16] * (len(input_rows[0]) - 5) + [44]),
        ('Metadata', metadata_rows, [22, 100]),
    ]

    write_workbook(OUTPUT_PATH, sheets)
    print(f'Wrote {OUTPUT_PATH}')


if __name__ == '__main__':
    main()
