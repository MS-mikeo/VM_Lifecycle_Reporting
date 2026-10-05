"""
Step 2 — Classify the raw VM inventory into retirement and price-impacted rows.

Reads:   data/vms.json, data/submap.json, config.json
Writes:  data/rows.json  (retire_rows, price_rows, price_subs, retire_subs)

Run from the repository root:  python scripts/2_build_rows.py
"""

import json
import os

import classify as cls

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DATA = os.path.join(ROOT, 'data')


def load_json(path):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def tag_lookup(tags, keys):
    if not isinstance(tags, dict):
        return ''
    low = {k.lower(): v for k, v in tags.items()}
    for k in keys:
        v = low.get(k.lower())
        if v not in (None, ''):
            return v
    return ''


def all_tags_str(tags):
    if not isinstance(tags, dict) or not tags:
        return ''
    return '; '.join(f'{k}={v}' for k, v in sorted(tags.items(), key=lambda kv: kv[0].lower()))


def main():
    config = load_json(os.path.join(ROOT, 'config.json'))
    vms = load_json(os.path.join(DATA, 'vms.json'))
    submap = load_json(os.path.join(DATA, 'submap.json'))
    tag_columns = config.get('tagColumns', [])

    retire_rows, price_rows = [], []
    price_subs, retire_subs = set(), set()

    for v in vms:
        size = v.get('vmSize')
        family, category = cls.classify(size)
        if not category:
            continue
        subid = v.get('subscriptionId')
        tags = v.get('tags')
        row = {
            'Subscription ID': subid,
            'Subscription Name': submap.get(subid, ''),
            'Resource Group': v.get('resourceGroup'),
            'VM Name': v.get('name'),
            'VM Size': size,
            'Family': family,
        }
        for col in tag_columns:
            row[col['header']] = tag_lookup(tags, col.get('keys', []))
        row['Power State'] = (v.get('powerState') or '').replace('PowerState/', '')
        row['All Tags'] = all_tags_str(tags)

        if category == 'retire':
            retire_rows.append(row)
            retire_subs.add(subid)
        else:
            row['Change Wave'] = cls.wave_of(category)
            price_rows.append(row)
            price_subs.add(subid)

    out = {
        'retire_rows': retire_rows,
        'price_rows': price_rows,
        'price_subs': sorted(price_subs),
        'retire_subs': sorted(retire_subs),
    }
    with open(os.path.join(DATA, 'rows.json'), 'w', encoding='utf-8') as f:
        json.dump(out, f, indent=1)

    print('retire rows:', len(retire_rows), '| price rows:', len(price_rows))
    print('distinct subs with retirement VMs:', len(retire_subs))
    print('distinct subs with price-impacted VMs:', len(price_subs))
    print('distinct subs needing a cost query:', len(price_subs))


if __name__ == '__main__':
    main()
