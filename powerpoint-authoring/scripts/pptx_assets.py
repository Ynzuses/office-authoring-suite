"""Local image contract shared by the bounded generator and its checker."""
import hashlib
from pathlib import Path
from urllib.parse import urlparse
from PIL import Image


def slide_asset_ids(slide):
    ids = [slide['image']] if slide.get('image') else []
    ids.extend(item['image'] for item in slide.get('items', []) if item.get('image'))
    return ids


def load_asset(spec, asset_id, base):
    asset = spec.get('assets', {}).get(asset_id)
    if not isinstance(asset, dict):
        raise ValueError(f'Missing asset definition: {asset_id}')
    unknown = set(asset) - {'path', 'kind', 'credit', 'license', 'source_url', 'sha256', 'alt'}
    if unknown:
        raise ValueError(f'Unsupported asset fields: {sorted(unknown)}')
    for key in ('path', 'credit', 'license'):
        if not isinstance(asset.get(key), str) or not asset[key].strip():
            raise ValueError(f'Asset {asset_id} requires {key}')
    if asset.get('kind') not in ('figure', 'photo', 'icon'):
        raise ValueError(f'Asset {asset_id} kind must be figure, photo or icon')
    url = asset.get('source_url', '')
    if url and urlparse(url).scheme not in ('http', 'https'):
        raise ValueError('source_url must be an HTTP(S) source page, not a local file')
    path = Path(base) / asset['path']
    data = path.read_bytes()
    sha = hashlib.sha256(data).hexdigest()
    if asset.get('sha256') and sha != asset['sha256']:
        raise ValueError(f'Asset checksum mismatch: {asset_id}')
    with Image.open(path) as im:
        if im.format not in ('PNG', 'JPEG'):
            raise ValueError('Use inspected PNG/JPEG assets; convert SVG separately and retain its original')
        size = im.size
        im.verify()
    return dict(asset, resolved_path=path, sha256=sha, pixels=size)
