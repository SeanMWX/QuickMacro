# save as offset_action.py
from pathlib import Path

def add_offset(src_path: str, dst_path: str, offset_ms: int):
    src = Path(src_path)
    dst = Path(dst_path)
    out_lines = []
    for line in src.read_text(encoding='utf-8').splitlines():
        s = line.strip()
        if not s or s.startswith('#'):
            out_lines.append(line)
            continue
        parts = s.split()
        # 时间戳在最后一个字段
        try:
            parts[-1] = str(int(parts[-1]) + offset_ms)
            out_lines.append(' '.join(parts))
        except Exception:
            # 如果这一行没有合法的末尾时间戳，原样保留
            out_lines.append(line)
    dst.write_text('\n'.join(out_lines) + '\n', encoding='utf-8')

if __name__ == '__main__':
    # 示例：把 20251124-154505.action 的时间加 5000ms，写到 new.action
    add_offset('actions/20251124-153808.action', 'actions/20251124-153808-offset.action', -13000)
    print('done')
