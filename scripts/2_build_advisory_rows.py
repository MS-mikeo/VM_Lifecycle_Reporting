"""Build advisory-specific VM rows. Run from the repository root."""
import argparse
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / 'data'
SIZE_RE = re.compile(r'^Standard_([A-Za-z]+?)(\d+)(?:-\d+)?([a-z]*)(?:_v(\d+))?(?:_Promo)?$')


def family(size):
    match = SIZE_RE.match(size or '')
    if not match:
        return None
    prefix, _, suffix, version = match.groups()
    base = prefix.upper()
    premium = len(base) > 1 and base.endswith('S') and base not in ('HC', 'NP')
    if premium:
        base = base[:-1]
    suffix = suffix.lower()
    premium = premium or suffix == 's'
    if version is None:
        if base == 'B': return 'Bv1'
        if base in ('D', 'F', 'G', 'L') and suffix in ('', 's'):
            return base + ('s' if premium else '')
        if base in ('NP', 'HC'): return base
        return None
    return f"{base}{'s' if premium else ''}v{version}"


def tags_text(tags):
    if not isinstance(tags, dict): return ''
    return '; '.join(f'{k}={v}' for k, v in sorted(tags.items(), key=lambda item: item[0].lower()))


def classify_prfr(fam):
    if fam in {'Dv3', 'Dsv3', 'Ev3', 'Esv3'}: return 'Retirement'
    if fam in {'Bv1', 'D', 'Ds', 'F', 'Fs', 'G', 'Gs', 'Ls', 'NP', 'HC'}: return 'Price increase v1'
    if fam in {'Av2', 'Amv2', 'Dv2', 'Dsv2', 'Fsv2', 'Lsv2'}: return 'Price increase v2'
    return None


parser = argparse.ArgumentParser()
parser.add_argument('--advisory', required=True, choices=('PRFR-_4Z', 'JGW1-KG0'))
args = parser.parse_args()
profiles = json.loads((ROOT / 'scripts' / 'advisories.json').read_text(encoding='utf-8'))
profile = profiles[args.advisory]
vms = json.loads((DATA / 'vms.json').read_text(encoding='utf-8'))
submap = json.loads((DATA / 'submap.json').read_text(encoding='utf-8'))
regions = profile.get('regions', {})
excluded = set(profile.get('excludedFamilies', []))
retire_rows, price_rows, price_subs = [], [], set()

for vm in vms:
    fam = family(vm.get('vmSize'))
    impact = classify_prfr(fam) if profile['mode'] == 'prfr' else None
    rate = profile.get('priceIncreasePercent', 0) / 100 if impact and impact.startswith('Price') else None
    if profile['mode'] == 'jgw':
        location = (vm.get('location') or '').lower()
        if location in regions and fam and fam not in excluded:
            impact, rate = 'Regional price increase', regions[location]
    if not impact: continue
    sid = vm.get('subscriptionId')
    row = {
        'Subscription ID': sid, 'Subscription Name': submap.get(sid, ''),
        'Resource Group': vm.get('resourceGroup'), 'VM Name': vm.get('name'),
        'Region': vm.get('location', ''), 'VM Size': vm.get('vmSize'), 'Family': fam,
        'Impact Type': impact, 'Increase Rate': rate,
        'Power State': (vm.get('powerState') or '').replace('PowerState/', ''),
        'All Tags': tags_text(vm.get('tags')),
    }
    if impact == 'Retirement': retire_rows.append(row)
    else:
        price_rows.append(row); price_subs.add(sid)

safe = args.advisory.replace('-', '_')
out = {'advisory': args.advisory, 'profile': profile, 'retirement_rows': retire_rows,
       'price_rows': price_rows, 'price_subs': sorted(price_subs)}
(DATA / f'rows_{safe}.json').write_text(json.dumps(out, indent=1), encoding='utf-8')
print(f'{args.advisory}: retirement rows={len(retire_rows)} | price rows={len(price_rows)} | subscriptions={len(price_subs)}')
