"""
classify.py — SKU classification ruleset for the Azure VM impact report.

This module encodes which VM size families are affected by a Microsoft health
advisory. It is intentionally the single place where the SKU taxonomy lives, so
you can adapt it to a *different* advisory by editing the rules below.

Two advisory "categories" are modelled:

  * retire     — families being retired / reaching end of life
  * price_v1   — "wave 1" families receiving a price change
  * price_v2   — "wave 2" families receiving a price change

The default rules reproduce an advisory covering:
  retire   : Dv3, Dsv3, Ev3, Esv3
  price v1 : Bv1, D, Ds, F, Fs, G, Gs, Ls, NP, HC
  price v2 : Av2, Amv2, Dv2, Dsv2, Fsv2, Lsv2

Nothing in this file is tenant- or organization-specific.
"""

import re

# Standard_<prefix><number>[-<constrained>]<suffix>[_v<N>][_Promo]
SIZE_RE = re.compile(r'^Standard_([A-Za-z]+?)(\d+)(?:-\d+)?([a-z]*)(?:_v(\d+))?(?:_Promo)?$')


def _parse(size):
    """Return (base, has_s, suffix, version) or None for a VM size string."""
    if not size:
        return None
    m = SIZE_RE.match(size)
    if not m:
        return None
    prefix, _num, suffix, ver = m.group(1), m.group(2), m.group(3), m.group(4)
    base = prefix.upper()
    has_s = False
    # A trailing 'S' on the letter prefix denotes premium-storage (e.g. DS -> D + s).
    if len(base) > 1 and base.endswith('S') and base not in ('HC', 'NP'):
        base = base[:-1]
        has_s = True
    return base, has_s, suffix.lower(), (int(ver) if ver else None)


def classify(size):
    """
    Map a VM size to (family_label, category).

    category is one of: 'retire', 'price_v1', 'price_v2'.
    Returns (None, None) when the size is not covered by the advisory.
    """
    parsed = _parse(size)
    if not parsed:
        return None, None
    base, has_s, suffix, version = parsed

    # --- Retirements: v3 D/E series ---
    if version == 3:
        if base == 'D':
            if suffix == '' and not has_s:
                return 'Dv3', 'retire'
            if suffix == 's' or has_s:
                return 'Dsv3', 'retire'
        if base == 'E':
            if suffix == '' and not has_s:
                return 'Ev3', 'retire'
            if suffix == 's' or has_s:
                return 'Esv3', 'retire'
        return None, None

    # --- Price wave 1: versionless (v1) families ---
    if version is None:
        if base == 'B':
            return 'Bv1', 'price_v1'
        if base in ('D', 'F', 'G', 'L') and suffix in ('', 's'):
            s = has_s or suffix == 's'
            if base == 'D':
                return ('Ds' if s else 'D'), 'price_v1'
            if base == 'F':
                return ('Fs' if s else 'F'), 'price_v1'
            if base == 'G':
                return ('Gs' if s else 'G'), 'price_v1'
            if base == 'L':
                return ('Ls', 'price_v1') if s else (None, None)
        if base == 'NP':
            return 'NP', 'price_v1'
        if base == 'HC':
            return 'HC', 'price_v1'
        return None, None

    # --- Price wave 2: v2 families ---
    if version == 2:
        if base == 'A':
            if suffix == 'm':
                return 'Amv2', 'price_v2'
            if suffix == '':
                return 'Av2', 'price_v2'
            return None, None
        if base == 'D' and suffix in ('', 's'):
            return ('Dsv2' if (has_s or suffix == 's') else 'Dv2'), 'price_v2'
        if base == 'F' and (has_s or suffix == 's'):
            return 'Fsv2', 'price_v2'
        if base == 'L' and (has_s or suffix == 's'):
            return 'Lsv2', 'price_v2'
        return None, None

    return None, None


def wave_of(category):
    """Map a category to a cost 'wave' label ('v1'/'v2'), or None for retirements."""
    if category == 'price_v1':
        return 'v1'
    if category == 'price_v2':
        return 'v2'
    return None


# ------------------------------------------------------------------
# Cost-side helpers: Azure billing meter names -> family group.
# Meters look like "D4 v2/DS4 v2", "DS3 v2 Spot", "F8s v2" etc.
# ------------------------------------------------------------------
_TOKEN_VER_RE = re.compile(r'^v\d+$')

# Azure bills several size pairs on a single combined meter (same compute rate,
# the 's' variant just adds premium storage). Collapse them to one row so the
# report does not show two confusing lines for what is really one family group.
_CANON = {
    'Dv2': 'Dv2/Dsv2', 'Dsv2': 'Dv2/Dsv2',
    'D': 'D/Ds', 'Ds': 'D/Ds',
    'F': 'F/Fs', 'Fs': 'F/Fs',
    'G': 'G/Gs', 'Gs': 'G/Gs',
}


def _token_to_size(token):
    """Turn a meter token like 'DS4 v2' into a 'Standard_DS4_v2' size string."""
    parts = token.strip().split()
    if not parts:
        return ''
    core = parts[0]
    ver = ''
    for p in parts[1:]:
        if _TOKEN_VER_RE.match(p):
            ver = '_' + p
    return 'Standard_' + core + ver


def meter_group(meter):
    """
    Map an Azure billing meter name to (clean_label, wave) when it corresponds
    to a price-impacted family, else (None, None).
    """
    clean = meter.split(' Spot')[0].split(' - ')[0].strip()
    tokens = [t for t in clean.split('/') if t.strip()]
    fams, wave = [], None
    for t in tokens:
        family, category = classify(_token_to_size(t))
        w = wave_of(category)
        if family and w:  # cost only covers price-impacted families
            if family not in fams:
                fams.append(family)
            wave = w
    if not fams:
        return None, None
    label = _CANON.get(fams[0], '/'.join(fams))
    return label, wave
