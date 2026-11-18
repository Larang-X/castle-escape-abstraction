import os
import base64
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen
from PIL import Image, ImageDraw, ImageFont

dat12_path = r'c:\Users\Shin hee jae\Desktop\font (1)\LC_KOR12.DAT'
dat8_path = r'c:\Users\Shin hee jae\Desktop\font (1)\LC_KOR8.DAT'

with open(dat12_path, 'rb') as f:
    data12 = f.read()

num_glyphs12 = len(data12) // 24

def get_glyph_grid12(idx):
    if idx < 0 or idx >= num_glyphs12:
        return [[False]*12 for _ in range(12)]
    chunk = data12[idx*24 : (idx+1)*24]
    grid = []
    for r in range(12):
        b1 = chunk[r*2]
        b2 = chunk[r*2 + 1]
        bits = f'{b1:08b}{b2:08b}'[:12]
        grid.append([b == '1' for b in bits])
    return grid

char_to_grid = {}

# 1. ASCII 32..126
for code in range(32, 127):
    idx = 32 + code
    char_to_grid[chr(code)] = get_glyph_grid12(idx)

# 2. Hangul syllables (KS X 1001: EUC-KR 0xB0A1 .. 0xC8FE)
hangul_idx = 959
for row in range(0xB0, 0xC9):
    for col in range(0xA1, 0xFF):
        try:
            ch = bytes([row, col]).decode('euc-kr')
            char_to_grid[ch] = get_glyph_grid12(hangul_idx)
            hangul_idx += 1
        except:
            pass

# 3. Direct symbols from DAT:
char_to_grid['→'] = get_glyph_grid12(25)
char_to_grid['➔'] = get_glyph_grid12(25)
char_to_grid['←'] = get_glyph_grid12(26)
char_to_grid['↑'] = get_glyph_grid12(27)
char_to_grid['↓'] = get_glyph_grid12(28)
char_to_grid['♥'] = get_glyph_grid12(23)
char_to_grid['◆'] = get_glyph_grid12(22)
char_to_grid['■'] = get_glyph_grid12(41)

# Circled numbers from DAT
char_to_grid['①'] = get_glyph_grid12(1)
char_to_grid['②'] = get_glyph_grid12(2)
char_to_grid['③'] = get_glyph_grid12(3)
char_to_grid['④'] = get_glyph_grid12(4)
char_to_grid['⑤'] = get_glyph_grid12(5)

# 4. Custom retro pixel star '★'
star_grid = [
    [0,0,0,0,0,1,0,0,0,0,0,0],
    [0,0,0,0,1,1,1,0,0,0,0,0],
    [0,0,0,0,1,1,1,0,0,0,0,0],
    [1,1,1,1,1,1,1,1,1,1,1,0],
    [0,1,1,1,1,1,1,1,1,1,0,0],
    [0,0,1,1,1,1,1,1,1,0,0,0],
    [0,0,0,1,1,1,1,1,0,0,0,0],
    [0,0,1,1,1,0,1,1,1,0,0,0],
    [0,1,1,1,0,0,0,1,1,1,0,0],
    [0,1,1,0,0,0,0,0,1,1,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
]
char_to_grid['★'] = [[b == 1 for b in r] for r in star_grid]
char_to_grid['⭐'] = [[b == 1 for b in r] for r in star_grid]

# Hollow star '☆'
star_hollow = [
    [0,0,0,0,0,1,0,0,0,0,0,0],
    [0,0,0,0,1,0,1,0,0,0,0,0],
    [0,0,0,0,1,0,1,0,0,0,0,0],
    [1,1,1,1,1,0,1,1,1,1,1,0],
    [0,1,0,0,0,0,0,0,0,1,0,0],
    [0,0,1,0,0,0,0,0,1,0,0,0],
    [0,0,0,1,0,0,0,1,0,0,0,0],
    [0,0,1,0,1,0,1,0,1,0,0,0],
    [0,1,0,0,0,0,0,0,0,1,0,0],
    [0,1,1,0,0,0,0,0,1,1,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
]
char_to_grid['☆'] = [[b == 1 for b in r] for r in star_hollow]

# Triangle '▲', '▼', '◀', '▶'
tri_up = [
    [0,0,0,0,0,1,0,0,0,0,0,0],
    [0,0,0,0,1,1,1,0,0,0,0,0],
    [0,0,0,1,1,1,1,1,0,0,0,0],
    [0,0,1,1,1,1,1,1,1,0,0,0],
    [0,1,1,1,1,1,1,1,1,1,0,0],
    [1,1,1,1,1,1,1,1,1,1,1,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
]
char_to_grid['▲'] = [[b == 1 for b in r] for r in tri_up]

tri_down = [
    [1,1,1,1,1,1,1,1,1,1,1,0],
    [0,1,1,1,1,1,1,1,1,1,0,0],
    [0,0,1,1,1,1,1,1,1,0,0,0],
    [0,0,0,1,1,1,1,1,0,0,0,0],
    [0,0,0,0,1,1,1,0,0,0,0,0],
    [0,0,0,0,0,1,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
]
char_to_grid['▼'] = [[b == 1 for b in r] for r in tri_down]

tri_right = [
    [1,1,0,0,0,0,0,0,0,0,0,0],
    [1,1,1,1,0,0,0,0,0,0,0,0],
    [1,1,1,1,1,1,0,0,0,0,0,0],
    [1,1,1,1,1,1,1,1,0,0,0,0],
    [1,1,1,1,1,1,1,1,1,1,0,0],
    [1,1,1,1,1,1,1,1,0,0,0,0],
    [1,1,1,1,1,1,0,0,0,0,0,0],
    [1,1,1,1,0,0,0,0,0,0,0,0],
    [1,1,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
]
char_to_grid['▶'] = [[b == 1 for b in r] for r in tri_right]
char_to_grid['▸'] = [[b == 1 for b in r] for r in tri_right]

# Middle dot '·', '•'
dot_grid = [[False]*12 for _ in range(12)]
dot_grid[5][5] = True
dot_grid[5][6] = True
dot_grid[6][5] = True
dot_grid[6][6] = True
char_to_grid['·'] = dot_grid
char_to_grid['•'] = dot_grid

# Ellipsis '…'
elli_grid = [[False]*12 for _ in range(12)]
for c in [2, 6, 10]:
    elli_grid[9][c] = True
char_to_grid['…'] = elli_grid

# Dashes
dash_grid = [[False]*12 for _ in range(12)]
for c in range(1, 11):
    dash_grid[5][c] = True
char_to_grid['—'] = dash_grid
char_to_grid['–'] = dash_grid

# Cross mark '✕', '❌'
cross_grid = [
    [1,1,0,0,0,0,0,0,1,1,0,0],
    [0,1,1,0,0,0,0,1,1,0,0,0],
    [0,0,1,1,0,0,1,1,0,0,0,0],
    [0,0,0,1,1,1,1,0,0,0,0,0],
    [0,0,0,0,1,1,0,0,0,0,0,0],
    [0,0,0,1,1,1,1,0,0,0,0,0],
    [0,0,1,1,0,0,1,1,0,0,0,0],
    [0,1,1,0,0,0,0,1,1,0,0,0],
    [1,1,0,0,0,0,0,0,1,1,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
]
char_to_grid['✕'] = [[b == 1 for b in r] for r in cross_grid]
char_to_grid['❌'] = [[b == 1 for b in r] for r in cross_grid]

# Swords / cross swords icon for '⚔'
sword_grid = [
    [1,0,0,0,0,0,0,0,0,0,1,0],
    [0,1,0,0,0,0,0,0,0,1,0,0],
    [0,0,1,0,0,0,0,0,1,0,0,0],
    [0,0,0,1,0,0,0,1,0,0,0,0],
    [0,0,0,0,1,1,1,0,0,0,0,0],
    [0,0,0,1,1,1,1,0,0,0,0,0],
    [0,0,1,0,0,0,0,1,0,0,0,0],
    [0,1,0,0,0,0,0,0,1,0,0,0],
    [1,0,0,0,0,0,0,0,0,1,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
]
char_to_grid['⚔'] = [[b == 1 for b in r] for r in sword_grid]

# Checkmark '✓', '✅'
check_grid = [
    [0,0,0,0,0,0,0,0,0,1,1,0],
    [0,0,0,0,0,0,0,0,1,1,0,0],
    [0,0,0,0,0,0,0,1,1,0,0,0],
    [0,0,0,0,0,0,1,1,0,0,0,0],
    [1,1,0,0,0,1,1,0,0,0,0,0],
    [0,1,1,0,1,1,0,0,0,0,0,0],
    [0,0,1,1,1,0,0,0,0,0,0,0],
    [0,0,0,1,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
    [0,0,0,0,0,0,0,0,0,0,0,0],
]
char_to_grid['✓'] = [[b == 1 for b in r] for r in check_grid]
char_to_grid['✅'] = [[b == 1 for b in r] for r in check_grid]

# Build TrueType Font with comfortable inter-glyph spacing
PIXEL_SIZE = 100
UPM = 1200
ASCENT = 1050
DESCENT = -250

def compile_ttf(output_path):
    fb = FontBuilder(UPM, isTTF=True)
    glyph_order = ['.notdef', 'space']
    char_map = {0x20: 'space'}

    sorted_chars = sorted(char_to_grid.keys())
    for ch in sorted_chars:
        if ch == ' ': continue
        code = ord(ch)
        gname = f'uni{code:04X}' if code <= 0xFFFF else f'u{code:06X}'
        glyph_order.append(gname)
        char_map[code] = gname

    fb.setupGlyphOrder(glyph_order)
    fb.setupCharacterMap(char_map)

    glyphs = {}
    metrics = {}

    # Empty notdef
    pen = TTGlyphPen(None)
    glyphs['.notdef'] = pen.glyph()
    metrics['.notdef'] = (650, 0)

    # Space: comfortable word gap
    pen = TTGlyphPen(None)
    glyphs['space'] = pen.glyph()
    metrics['space'] = (600, 0)

    for ch in sorted_chars:
        if ch == ' ': continue
        gname = char_map[ord(ch)]
        grid = char_to_grid[ch]
        
        is_ascii = ord(ch) < 128
        used_cols = [c for r in range(12) for c in range(12) if grid[r][c]]

        if not used_cols:
            pen = TTGlyphPen(None)
            glyphs[gname] = pen.glyph()
            adv = 600 if is_ascii else 1320
            metrics[gname] = (adv, 0)
            continue

        min_c = min(used_cols)
        max_c = max(used_cols)
        w_pix = max_c - min_c + 1

        if is_ascii:
            # Shift glyph so min_c starts at 1 (100 units left bearing)
            col_shift = min_c - 1 if min_c > 0 else 0
            if ch in 'il!|:;,.':
                col_shift = min_c - 1 if min_c > 1 else 0
                adv = max(480, (w_pix + 3) * PIXEL_SIZE)
            elif ch in '[](){}':
                col_shift = min_c - 1 if min_c > 0 else 0
                adv = 620
            elif ch in 'mwMW@%#':
                adv = max(800, (w_pix + 2) * PIXEL_SIZE)
            elif ch in '0123456789':
                # Uniform tabular digits with generous breathing room
                adv = 720
            else:
                adv = 700
            x_pad = 20
        else:
            # Hangul and fullwidth symbols: add 1.2px extra advance width + 40 units left pad
            col_shift = 0
            adv = 1320
            x_pad = 40

        pen = TTGlyphPen(None)
        for r in range(12):
            y_top = (10 - r) * PIXEL_SIZE
            y_bottom = (9 - r) * PIXEL_SIZE

            c = 0
            while c < 12:
                if grid[r][c]:
                    c_start = c
                    while c < 12 and grid[r][c]:
                        c += 1
                    c_end = c - 1

                    x_left = (c_start - col_shift) * PIXEL_SIZE + x_pad
                    x_right = (c_end + 1 - col_shift) * PIXEL_SIZE + x_pad

                    pen.moveTo((x_left, y_bottom))
                    pen.lineTo((x_right, y_bottom))
                    pen.lineTo((x_right, y_top))
                    pen.lineTo((x_left, y_top))
                    pen.closePath()
                else:
                    c += 1

        glyphs[gname] = pen.glyph()
        lsb = min((c - col_shift) * PIXEL_SIZE + x_pad for c in used_cols)
        metrics[gname] = (adv, lsb)

    fb.setupGlyf(glyphs)
    fb.setupHorizontalMetrics(metrics)
    fb.setupHorizontalHeader(ascent=ASCENT, descent=DESCENT, lineGap=200)
    fb.setupNameTable({
        'familyName': 'CastlevaniaKorean',
        'styleName': 'Regular',
        'uniqueFontIdentifier': 'CastlevaniaKorean-Regular',
        'fullName': 'CastlevaniaKorean Regular',
        'psName': 'CastlevaniaKorean-Regular',
        'version': 'Version 1.1'
    })
    fb.setupOS2(sTypoAscender=ASCENT, sTypoDescender=DESCENT, sTypoLineGap=200,
                 usWinAscent=ASCENT, usWinDescent=-DESCENT,
                 sxHeight=600, sCapHeight=800)
    fb.setupPost()
    fb.save(output_path)
    print(f'Compiled TTF font to: {output_path} ({os.path.getsize(output_path)} bytes)')

os.makedirs('assets', exist_ok=True)
compile_ttf('assets/castlevania_kor.ttf')
compile_ttf('assets/castlevania_por.ttf')
compile_ttf('assets/castlevania_por_small.ttf')

# Read base64
with open('assets/castlevania_kor.ttf', 'rb') as f:
    ttf_bytes = f.read()

b64_str = base64.b64encode(ttf_bytes).decode('ascii')
print(f'Base64 length: {len(b64_str)}')

with open('scratch_font_b64.txt', 'w') as f:
    f.write(b64_str)

print('Done!')
