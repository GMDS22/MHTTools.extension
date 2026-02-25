import csv
import re
from collections import defaultdict, Counter

CSV_PATH = 'family_name_suggestions.csv'

rows = []
# Read with error-tolerant decoding because exported CSV may contain non-UTF8 bytes (degree symbol, etc.)
with open(CSV_PATH, newline='', encoding='utf-8', errors='replace') as f:
    reader = csv.DictReader(f)
    for r in reader:
        rows.append(r)

total = len(rows)

# helper
def words(s):
    return re.findall(r"[A-Za-z]{2,}", s)

# stats
und_count = 0
short_suggestion_count = 0
numeric_only_suggestion_count = 0
duplicates = Counter()
sug_to_families = defaultdict(list)
problems = []

for r in rows:
    fam = r['Family']
    cat = r['Category']
    sug = (r['SuggestedName'] or '').strip()
    cur = fam
    duplicates[sug] += 1
    sug_to_families[sug].append(fam)
    if 'UND' in sug.upper():
        und_count += 1
    # numeric-only patterns like MHT-ME-AT-01 or ending with -01
    if re.match(r'^MHT-[-A-Z0-9]+-\d{2}$', sug) or re.search(r'-\d{2}$', sug):
        numeric_only_suggestion_count += 1
    # short suggestion: fewer than 3 alpha words
    if len(words(sug)) < 2:
        short_suggestion_count += 1

    # heuristic: current name is more descriptive when it contains >2 words and suggestion contains UND or only short
    cur_word_count = len(words(cur))
    sug_word_count = len(words(sug))
    if cur_word_count >= 3 and (('UND' in sug.upper()) or sug_word_count < 2 or sug_word_count < cur_word_count):
        problems.append({'family': fam, 'category': cat, 'current': cur, 'suggested': sug, 'cur_words': cur_word_count, 'sug_words': sug_word_count})

# top duplicates
dup_list = [(s, c) for s, c in duplicates.items() if c > 1]
dup_list.sort(key=lambda x: x[1], reverse=True)

# prepare report
print('Total rows: {}'.format(total))
print('Suggestions containing UND: {}'.format(und_count))
print('Numeric-only or auto-id suggestions (e.g., -01): {}'.format(numeric_only_suggestion_count))
print('Short suggestions (fewer than 2 alpha words): {}'.format(short_suggestion_count))
print('Distinct suggested names: {}'.format(len(duplicates)))
print('\nTop duplicated suggested names:')
for s, c in dup_list[:20]:
    print('  {}  -> {} occurrences'.format(s, c))
    print('    Examples: {}'.format(', '.join(sug_to_families[s][:5])))

print('\nSample problematic rows where current name seems more descriptive:')
for p in problems[:40]:
    print(' - Family: {}'.format(p['family']))
    print('   Category: {}'.format(p['category']))
    print('   Current: {}'.format(p['current']))
    print('   Suggested: {}'.format(p['suggested']))
    print('   Current word count: {}, Suggested word count: {}'.format(p['cur_words'], p['sug_words']))

# Quick recommendations based on findings
print('\nRecommendations:')
print(' - Add more system token scanning in type parameter VALUES (we implemented some already).')
print(' - Expand size parameter synonyms and parsing (DN, Ø, Dia, "inches").')
print(' - For Duct Fittings, prefer templates that include Shape-Fitting-Modifier-SYS-SIZE to reduce UND fallbacks.')
print(' - When suggestion would be less descriptive than current name, prefer keeping current name or use family short token in template to preserve uniqueness.')

# Save a minimal CSV of problem examples for review
OUT = 'problem_examples.csv'
with open(OUT, 'w', newline='', encoding='utf-8') as wf:
    w = csv.DictWriter(wf, fieldnames=['Family','Category','Current','Suggested'])
    w.writeheader()
    for p in problems:
        w.writerow({'Family': p['family'], 'Category': p['category'], 'Current': p['current'], 'Suggested': p['suggested']})
print('\nWrote {} problem examples for review.'.format(OUT))
