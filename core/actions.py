from typing import Dict


def parse_action_line(s: str) -> Dict[str, str]:
    s = s.rstrip("\n")
    d = {
        'type':'', 'op':'', 'vk':'', 'btn':'', 'act':'',
        'x':'', 'y':'', 'nx':'', 'ny':'', 'dx':'', 'dy':'', 'ms':'', 'raw': s
    }
    if not s:
        return d
    if s.startswith('#'):
        d['type'] = '#'
        return d
    parts = s.split()
    try:
        if parts[0] == 'META':
            d['type'] = 'META'
            d['op'] = parts[1] if len(parts) > 1 else ''
            if d['op'] == 'SCREEN':
                d['x'] = parts[2] if len(parts) > 2 else ''
                d['y'] = parts[3] if len(parts) > 3 else ''
            elif d['op'] == 'START':
                d['x'] = parts[2] if len(parts) > 2 else ''
            elif d['op'] == 'RESTART':
                d['x'] = parts[2] if len(parts) > 2 else ''
            return d
        if parts[0] == 'K':
            d['type'] = 'VK'
            d['op'] = parts[1] if len(parts) > 1 else ''
            d['vk'] = parts[2] if len(parts) > 2 else ''
            d['ms'] = parts[-1] if len(parts) > 3 else ''
            return d
        if parts[0] == 'M':
            d['type'] = 'MS'
            if parts[1] == 'MOVE':
                d['op'] = 'MOVE'
                if len(parts) >= 7:
                    d['x'], d['y'], d['nx'], d['ny'], d['ms'] = parts[2], parts[3], parts[4], parts[5], parts[6]
                else:
                    d['x'], d['y'] = (parts[2] if len(parts) > 2 else ''), (parts[3] if len(parts) > 3 else '')
                    d['ms'] = parts[4] if len(parts) > 4 else ''
                return d
            if parts[1] == 'CLICK':
                d['op'] = 'CLICK'
                d['btn'] = parts[2] if len(parts) > 2 else ''
                d['act'] = parts[3] if len(parts) > 3 else ''
                if len(parts) >= 10:
                    d['x'], d['y'], d['nx'], d['ny'], d['ms'] = parts[4], parts[5], parts[6], parts[7], parts[8]
                else:
                    d['x'] = parts[4] if len(parts) > 4 else ''
                    d['y'] = parts[5] if len(parts) > 5 else ''
                    d['ms'] = parts[-1] if len(parts) > 6 else ''
                return d
            if parts[1] == 'SCROLL':
                d['op'] = 'SCROLL'
                d['dx'] = parts[2] if len(parts) > 2 else ''
                d['dy'] = parts[3] if len(parts) > 3 else ''
                d['ms'] = parts[4] if len(parts) > 4 else ''
                return d
    except Exception:
        return d
    return d


def compose_action_line(d: Dict[str, str]) -> str:
    t = (d.get('type') or '').strip()
    op = (d.get('op') or '').strip()
    if t == '#':
        return d.get('raw', '')
    if t == 'META':
        if op.upper() == 'SCREEN':
            return f"META SCREEN {d.get('x','')} {d.get('y','')}".strip()
        if op.upper() == 'START':
            return f"META START {d.get('x','')}".strip()
        if op.upper() == 'RESTART':
            return f"META RESTART {d.get('x','')}".strip()
        return d.get('raw', '')
    if t == 'VK':
        return f"K {op} {d.get('vk','')} {d.get('ms','')}".strip()
    if t == 'MS':
        if op.upper() == 'MOVE':
            x,y,nx,ny,ms = d.get('x',''), d.get('y',''), d.get('nx',''), d.get('ny',''), d.get('ms','')
            return f"M MOVE {x} {y} {nx} {ny} {ms}".strip() if nx and ny else f"M MOVE {x} {y} {ms}".strip()
        if op.upper() == 'CLICK':
            btn, act = d.get('btn',''), d.get('act','')
            x,y,nx,ny,ms = d.get('x',''), d.get('y',''), d.get('nx',''), d.get('ny',''), d.get('ms','')
            return f"M CLICK {btn} {act} {x} {y} {nx} {ny} {ms}".strip() if nx and ny else f"M CLICK {btn} {act} {x} {y} {ms}".strip()
        if op.upper() == 'SCROLL':
            dx,dy,ms = d.get('dx',''), d.get('dy',''), d.get('ms','')
            return f"M SCROLL {dx} {dy} {ms}".strip()
    return d.get('raw','') or ''


def extract_meta(lines):
    meta_screen = ''
    meta_start = ''
    meta_restart = ''
    meta_lines = []
    for s in lines:
        if s.startswith('META '):
            meta_lines.append(s)
            parts = s.split()
            if len(parts) >= 4 and parts[1] == 'SCREEN':
                meta_screen = f"{parts[2]}x{parts[3]}"
            if len(parts) >= 3 and parts[1] == 'START':
                meta_start = parts[2]
            if len(parts) >= 3 and parts[1] == 'RESTART':
                meta_restart = parts[2]
    return meta_lines, meta_screen, meta_start, meta_restart
