"""Reads the rendered report PDF and writes pages.json (TOC key -> printed page label).

Usage: python3 toc_pages.py MediKiosk_Project_Report.pdf pages.json
"""
import json
import subprocess
import sys

ROMAN = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x', 'xi', 'xii']

SECTIONS = ['1.1 INTRODUCTION OF THE PROJECT', '1.2 NEED OF THE PROJECT', '1.3 OBJECTIVE OF THE PROJECT',
            '1.4 THEME OF THE PROJECT', '1.5 ORGANIZATION', '1.6 SUMMARY',
            '2.1 INTRODUCTION', '2.2 EXISTING SYSTEMS AND THEIR LIMITATIONS', '2.3 SCOPE OF THE PROJECT', '2.4 SUMMARY',
            '3.1 FUNDAMENTALS OF AI-BASED CLINICAL INTAKE', '3.2 PROPOSED SYSTEM ARCHITECTURE', '3.3 SYSTEM WORKFLOW',
            '3.4 IMPLEMENTATION DETAILS', '3.5 SUMMARY',
            '4.1 FUNCTIONAL ANALYSIS', '4.2 EXPERIMENTAL ANALYSIS', '4.3 SUMMARY',
            '5.1 CONCLUSION', '5.2 FUTURE SCOPE', '5.3 APPLICATIONS', '5.4 SUMMARY']
CHAPTERS = {'ch1': 'CHAPTER 1 – INTRODUCTION', 'ch2': 'CHAPTER 2 – LITERATURE REVIEW',
            'ch3': 'CHAPTER 3 – SYSTEM DESIGN', 'ch4': 'CHAPTER 4 – RESULT AND DISCUSSION',
            'ch5': 'CHAPTER 5 – CONCLUSION'}


def main(pdf, out):
    text = subprocess.run(['pdftotext', '-layout', pdf, '-'], capture_output=True, text=True, check=True).stdout
    pages = [' '.join(p.split()) for p in text.split('\f')]

    def find(needle, start=0):
        for i in range(start, len(pages)):
            if needle in pages[i]:
                return i
        raise SystemExit(f'not found: {needle}')

    abstract = find('ABSTRACT')
    abbrev = find('LIST OF ABBREVIATIONS', find('TABLE OF CONTENTS') + 1)
    ch1 = find(CHAPTERS['ch1'], abbrev + 1)
    result = {'abstract': ROMAN[4], 'abbrev': ROMAN[4 + abbrev - abstract]}
    for key, needle in CHAPTERS.items():
        result[key] = str(find(needle, ch1) - ch1 + 1)
    for s in SECTIONS:
        result[s.split()[0]] = str(find(s, ch1) - ch1 + 1)
    result['refs'] = str(find('REFERENCES', find(CHAPTERS['ch5'], ch1)) - ch1 + 1)
    json.dump(result, open(out, 'w'), indent=1)
    print(json.dumps(result))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
