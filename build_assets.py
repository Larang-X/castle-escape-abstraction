import struct, os, wave, math, base64, json
import numpy as np
from PIL import Image, ImageDraw

ROM_PATH = r'c:\Users\Shin hee jae\Desktop\캐슬바니아 오더 오브 에클레시아.nds'
BASE_DIR = r'c:\Users\Shin hee jae\Desktop\최종 출품작(게임)'
SPR_DIR = os.path.join(BASE_DIR, 'assets', 'sprites')
TILE_DIR = os.path.join(BASE_DIR, 'assets', 'tiles')
AUD_DIR = os.path.join(BASE_DIR, 'assets', 'audio')

for d in (SPR_DIR, TILE_DIR, AUD_DIR):
    os.makedirs(d, exist_ok=True)

with open(ROM_PATH, 'rb') as f:
    rom = f.read()

fnt_off, fnt_size, fat_off, fat_size = struct.unpack('<IIII', rom[0x40:0x50])
ov9_off, ov9_size = struct.unpack('<II', rom[0x50:0x58])
fnt = rom[fnt_off:fnt_off+fnt_size]
fat = rom[fat_off:fat_off+fat_size]

sub_off, first_id, total_dirs = struct.unpack('<IHH', fnt[0:8])
dirs = {0xF000: ''}
files = {}
for i in range(total_dirs):
    d_id = 0xF000 + i
    s_off, f_id, parent_id = struct.unpack('<IHH', fnt[i*8:(i+1)*8])
    pos = s_off
    cur_fid = f_id
    while pos < len(fnt):
        length = fnt[pos]
        pos += 1
        if length == 0: break
        is_dir = (length & 0x80) != 0
        name_len = length & 0x7F
        name = fnt[pos:pos+name_len].decode('ascii', errors='replace')
        pos += name_len
        if is_dir:
            sub_dir_id = struct.unpack('<H', fnt[pos:pos+2])[0]
            pos += 2
            dirs[sub_dir_id] = (dirs.get(d_id, '') + '/' + name)
        else:
            start, end = struct.unpack('<II', fat[cur_fid*8:(cur_fid+1)*8])
            full = dirs.get(d_id, '') + '/' + name
            files[full] = rom[start:end]
            cur_fid += 1

def read_pal_ram(ram_addr, ov_id=None):
    if ov_id is None:
        rom_off = ram_addr - 0x02000000 + 0x4000
    else:
        _, base_ram, _, _, _, _, fid, _ = struct.unpack('<8I', rom[ov9_off + ov_id*32 : ov9_off + (ov_id+1)*32])
        st, ed = struct.unpack('<II', fat[fid*8:(fid+1)*8])
        rom_off = st + (ram_addr - base_ram)
    w0, num_pals = struct.unpack('<HH', rom[rom_off:rom_off+4])
    pals = []
    for p in range(num_pals):
        cols = struct.unpack('<16H', rom[rom_off+4+p*32 : rom_off+4+(p+1)*32])
        rgba = []
        for idx, c in enumerate(cols):
            if idx == 0:
                rgba.append((0, 0, 0, 0))
            else:
                r = (c & 0x1F) * 255 // 31
                g = ((c >> 5) & 0x1F) * 255 // 31
                b = ((c >> 10) & 0x1F) * 255 // 31
                rgba.append((r, g, b, 255))
        pals.append(rgba)
    return pals

def read_ov_pal_at_offset(ov_id, byte_offset):
    _, base_ram, _, _, _, _, fid, _ = struct.unpack('<8I', rom[ov9_off + ov_id*32 : ov9_off + (ov_id+1)*32])
    st, ed = struct.unpack('<II', fat[fid*8:(fid+1)*8])
    rom_off = st + byte_offset
    w0, num_pals = struct.unpack('<HH', rom[rom_off:rom_off+4])
    pals = []
    for p in range(num_pals):
        cols = struct.unpack('<16H', rom[rom_off+4+p*32 : rom_off+4+(p+1)*32])
        rgba = []
        for idx, c in enumerate(cols):
            if idx == 0:
                rgba.append((0, 0, 0, 0))
            else:
                r = (c & 0x1F) * 255 // 31
                g = ((c >> 5) & 0x1F) * 255 // 31
                b = ((c >> 10) & 0x1F) * 255 // 31
                rgba.append((r, g, b, 255))
        pals.append(rgba)
    return pals

def decode_sheet_4bpp(data):
    if len(data) > 8192 and len(data) % 8192 != 0:
        data = data[-8192:]
    h = (len(data) // 64)
    pixels = [[0]*128 for _ in range(h)]
    for y in range(h):
        for x_byte in range(64):
            b = data[y*64 + x_byte]
            pixels[y][x_byte*2] = b & 0x0F
            pixels[y][x_byte*2 + 1] = (b >> 4) & 0x0F
    return pixels

def extract_single_frame(p_path, f_paths, pals, frame_idx=0):
    pdata = files[p_path]
    ints = struct.unpack('<16I', pdata[:64])
    parts_off = ints[1]
    frames_off = ints[3]
    num_frames = ints[9]
    valid_frames = []
    for fi in range(num_frames):
        fe = pdata[frames_off + fi*12 : frames_off + (fi+1)*12]
        unk, nhit, nparts, hit_off, part_off = struct.unpack('<HBBII', fe)
        if nparts > 0:
            valid_frames.append(fi)
    if not valid_frames:
        return Image.new('RGBA', (16, 16), (0, 0, 0, 0))
    f_use = valid_frames[frame_idx % len(valid_frames)]
    sheets = [decode_sheet_4bpp(files[fp]) for fp in f_paths]
    f_entry = pdata[frames_off + f_use*12 : frames_off + (f_use+1)*12]
    unk, nhit, nparts, hit_off, part_off = struct.unpack('<HBBII', f_entry)
    parts = []
    min_x, min_y, max_x, max_y = 999, 999, -999, -999
    for i in range(nparts):
        poff = parts_off + part_off + i*16
        xp, yp, sx, sy, w, h, sh_fl, pal_i = struct.unpack('<hhHHHHHH', pdata[poff:poff+16])
        parts.append((xp, yp, sx, sy, w, h, sh_fl, pal_i))
        min_x = min(min_x, xp)
        min_y = min(min_y, yp)
        max_x = max(max_x, xp + w)
        max_y = max(max_y, yp + h)
    fw = max(1, max_x - min_x)
    fh = max(1, max_y - min_y)
    img = Image.new('RGBA', (fw, fh), (0, 0, 0, 0))
    px = img.load()
    for xp, yp, sx, sy, w, h, sh_fl, pal_i in reversed(parts):
        sh_idx = (sh_fl & 0xFF) % len(sheets)
        hflip = (sh_fl & 0x0100) != 0
        vflip = (sh_fl & 0x0200) != 0
        pal = pals[pal_i % len(pals)]
        sheet = sheets[sh_idx]
        for dy in range(h):
            for dx in range(w):
                src_px_x = sx + (w - 1 - dx if hflip else dx)
                src_px_y = sy + (h - 1 - dy if vflip else dy)
                if 0 <= src_px_y < len(sheet) and 0 <= src_px_x < 128:
                    cidx = sheet[src_px_y][src_px_x]
                    if cidx != 0:
                        px[xp - min_x + dx, yp - min_y + dy] = pal[cidx]
    return img

# 1. Extract key Castlevania sprites + Diverse Stage Monsters (2 frames each)
sprite_jobs = [
    ('garg_idle', '/so/p_garg00.dat', ['/sc/f_garg00.dat', '/sc/f_garg01.dat', '/sc/f_garg02.dat'], read_pal_ram(0x020CB9A4), 2),
    ('garg_fly', '/so/p_garg00.dat', ['/sc/f_garg00.dat', '/sc/f_garg01.dat', '/sc/f_garg02.dat'], read_pal_ram(0x020CB9A4), 4),
    ('coffer_closed', '/so/p_coffer00.dat', ['/sc/f_coffer00.dat'], read_pal_ram(0x020CD294), 0),
    ('coffer_open', '/so/p_coffer00.dat', ['/sc/f_coffer00.dat'], read_pal_ram(0x020CD294), 3),
    ('statue_knight', '/so/p_rstatue.dat', ['/sc/f_rstatue.dat'], read_pal_ram(0x020CDEB0), 0),
    ('pillar', '/so/p_hasira.dat', ['/sc/f_hasira.dat'], read_pal_ram(0x020CDEF8), 0),
    ('sword_glyph', '/so/p_raw00.dat', ['/sc2/f_raw00.dat'], read_pal_ram(0x020D65A8), 0),
    ('rapier_glyph', '/so/p_raw02.dat', ['/sc2/f_raw02.dat'], read_pal_ram(0x020D65F0), 0),
    ('lance_glyph', '/so/p_raw04.dat', ['/sc2/f_raw04.dat'], read_pal_ram(0x020D6678), 0),
    ('great_glyph', '/so/p_raw01.dat', ['/sc2/f_raw01.dat'], read_pal_ram(0x020D65CC), 0),
    ('gear_large', '/so/p_gear00.dat', ['/sc/f_gear00.dat'], read_pal_ram(0x022A8D44, ov_id=21), 0),
    ('fire_f1', '/so/p_fire00.dat', ['/sc/f_fire00.dat'], read_pal_ram(0x02224C9C, ov_id=18), 0),
    ('fire_f2', '/so/p_fire00.dat', ['/sc/f_fire00.dat'], read_pal_ram(0x02224C9C, ov_id=18), 2),
    ('door_closed', '/so/p_aadoor.dat', ['/sc/f_aadoor.dat'], read_pal_ram(0x022CA36C, ov_id=40), 0),
    ('door_open', '/so/p_aadoor.dat', ['/sc/f_aadoor.dat'], read_pal_ram(0x022CA36C, ov_id=40), 2),
    ('candle_f1', '/so/p_cand00.dat', ['/sc/f_cand00.dat', '/sc/f_cand01.dat'], read_pal_ram(0x022EC848, ov_id=67), 0),
    ('candle_f2', '/so/p_cand00.dat', ['/sc/f_cand00.dat', '/sc/f_cand01.dat'], read_pal_ram(0x022EC848, ov_id=67), 1),
    ('portal_f1', '/so/p_wap00.dat', ['/sc/f_wap00.dat', '/sc/f_wap01.dat', '/sc/f_wap02.dat'], read_pal_ram(0x020CE054), 0),
    ('dest_vase', '/so/p_dest00.dat', ['/sc/f_dest00.dat'], read_pal_ram(0x020CD8DC), 0),
    ('dest_chandelier', '/so/p_dest02.dat', ['/sc/f_dest02.dat'], read_pal_ram(0x020CD9A4), 0),
    ('dest_skull', '/so/p_dest04.dat', ['/sc/f_dest04.dat'], read_pal_ram(0x020CDB0C), 0),
    ('skeleton', '/so/p_ya_skl.dat', ['/sc/t_ya_skl.dat'], read_pal_ram(0x020CAE94), 0),
    ('shutter', '/so/p_shutter00.dat', ['/sc/f_shutter00.dat'], read_pal_ram(0x020D11C4), 0),
    # Diverse animated monsters per stage
    ('imp_f1', '/so/p_imp.dat', ['/sc/f_imp.dat'], read_pal_ram(0x020C9D48), 0),
    ('imp_f2', '/so/p_imp.dat', ['/sc/f_imp.dat'], read_pal_ram(0x020C9D48), 2),
    ('rskl_f1', '/so/p_rskl.dat', ['/sc/f_rskl.dat'], read_pal_ram(0x020CAFA0), 0),
    ('rskl_f2', '/so/p_rskl.dat', ['/sc/f_rskl.dat'], read_pal_ram(0x020CAFA0), 3),
    ('zbs_f1', '/so/p_zbs.dat', ['/sc/f_zbs00.dat', '/sc/f_zbs01.dat'], read_pal_ram(0x020CB898), 0),
    ('zbs_f2', '/so/p_zbs.dat', ['/sc/f_zbs00.dat', '/sc/f_zbs01.dat'], read_pal_ram(0x020CB898), 3),
    ('gho_f1', '/so/p_gho00.dat', ['/sc/f_gho00.dat', '/sc/f_gho01.dat', '/sc/f_gho02.dat', '/sc/f_gho03.dat'], read_pal_ram(0x020CB6E4), 0),
    ('gho_f2', '/so/p_gho00.dat', ['/sc/f_gho00.dat', '/sc/f_gho01.dat', '/sc/f_gho02.dat', '/sc/f_gho03.dat'], read_pal_ram(0x020CB6E4), 3),
    ('sfish_f1', '/so/p_sfish0.dat', ['/sc/f_sfish0.dat'], read_pal_ram(0x020CBA4C), 0),
    ('sfish_f2', '/so/p_sfish0.dat', ['/sc/f_sfish0.dat'], read_pal_ram(0x020CBA4C), 2),
    ('yeti_f1', '/so/p_yeti.dat', ['/sc/f_yeti0.dat', '/sc/f_yeti1.dat'], read_pal_ram(0x020CC844), 0),
    ('yeti_f2', '/so/p_yeti.dat', ['/sc/f_yeti0.dat', '/sc/f_yeti1.dat'], read_pal_ram(0x020CC844), 3),
    ('medusa_f1', '/so/p_medusa.dat', ['/sc/f_medus0.dat'], read_pal_ram(0x020CAF3C), 0),
    ('medusa_f2', '/so/p_medusa.dat', ['/sc/f_medus0.dat'], read_pal_ram(0x020CAF3C), 2),
    ('tinman_f1', '/so/p_tinman.dat', ['/sc/t_tinman.dat'], read_pal_ram(0x020CC5CC), 0),
    ('tinman_f2', '/so/p_tinman.dat', ['/sc/t_tinman.dat'], read_pal_ram(0x020CC5CC), 2),
    ('eye_f1', '/so/p_nomi00.dat', ['/sc/f_nomi00.dat', '/sc/f_nomi01.dat', '/sc/f_nomi02.dat'], read_pal_ram(0x020CA0C4), 0),
    ('eye_f2', '/so/p_nomi00.dat', ['/sc/f_nomi00.dat', '/sc/f_nomi01.dat', '/sc/f_nomi02.dat'], read_pal_ram(0x020CA0C4), 2),
    ('axarm_f1', '/so/p_axarm0.dat', ['/sc/f_axarm0.dat', '/sc/f_axarm1.dat', '/sc/f_axarm2.dat'], read_pal_ram(0x020CCF84), 0),
    ('axarm_f2', '/so/p_axarm0.dat', ['/sc/f_axarm0.dat', '/sc/f_axarm1.dat', '/sc/f_axarm2.dat'], read_pal_ram(0x020CCF84), 3),
    ('mimi_f1', '/so/p_mimi00.dat', ['/sc/f_mimi00.dat', '/sc/f_mimi01.dat'], read_pal_ram(0x020CB1BC), 0),
    ('mimi_f2', '/so/p_mimi00.dat', ['/sc/f_mimi00.dat', '/sc/f_mimi01.dat'], read_pal_ram(0x020CB1BC), 4),
    # Stage 4 High-Tier Boss Monsters: Death (사신 데스), Crimson Wraith, Necromancer, Succubus Lilith, Werewolf
    ('death_f1', '/so/p_sgami00.dat', [f'/sc/f_sgami{i:02x}.dat' for i in range(15)], read_pal_ram(0x022BEFCC, ov_id=24), 22),
    ('death_f2', '/so/p_sgami00.dat', [f'/sc/f_sgami{i:02x}.dat' for i in range(15)], read_pal_ram(0x022BEFCC, ov_id=24), 26),
    ('wraith_f1', '/so/p_sgami00.dat', [f'/sc/f_sgami{i:02x}.dat' for i in range(15)], read_pal_ram(0x022BEFCC, ov_id=24), 6),
    ('wraith_f2', '/so/p_sgami00.dat', [f'/sc/f_sgami{i:02x}.dat' for i in range(15)], read_pal_ram(0x022BEFCC, ov_id=24), 4),
    ('necro_f1', '/so/p_necro0.dat', ['/sc/f_necro0.dat', '/sc/f_necro1.dat'], read_pal_ram(0x020CB814), 0),
    ('necro_f2', '/so/p_necro0.dat', ['/sc/f_necro0.dat', '/sc/f_necro1.dat'], read_pal_ram(0x020CB814), 2),
    ('lilith_f1', '/so/p_draf00.dat', ['/sc/f_draf00.dat', '/sc/f_draf01.dat', '/sc/f_draf02.dat'], read_pal_ram(0x020CBB30), 14),
    ('lilith_f2', '/so/p_draf00.dat', ['/sc/f_draf00.dat', '/sc/f_draf01.dat', '/sc/f_draf02.dat'], read_pal_ram(0x020CBB30), 16),
    ('wwolf_f1', '/so/p_wwolf.dat', [f'/sc/f_wwolf{i}.dat' for i in range(6)], read_pal_ram(0x020CAC30), 0),
    ('wwolf_f2', '/so/p_wwolf.dat', [f'/sc/f_wwolf{i}.dat' for i in range(6)], read_pal_ram(0x020CAC30), 2),
    ('albus_f1', '/so/p_albus.dat', ['/sc2/f_albus00.dat', '/sc2/f_albus01.dat', '/sc2/f_albus02.dat'], read_pal_ram(0x020CB814), 0),
    ('albus_f2', '/so/p_albus.dat', ['/sc2/f_albus00.dat', '/sc2/f_albus01.dat', '/sc2/f_albus02.dat'], read_pal_ram(0x020CB814), 4),
    ('barlow_f1', '/so/p_bar00.dat', ['/sc/f_bar00.dat'], read_pal_ram(0x020CB814), 0),
    ('barlow_f2', '/so/p_bar00.dat', ['/sc/f_bar00.dat'], read_pal_ram(0x020CB814), 2),
]

bundle_data = {}

for name, p_path, f_paths, pals, fidx in sprite_jobs:
    try:
        im = extract_single_frame(p_path, f_paths, pals, fidx)
        if name.startswith('candle_'):
            im = im.crop((0, 0, min(32, im.width), min(32, im.height)))
        elif name.startswith('imp_'):
            # Fix upside-down Imp orientation so wings & eyes are upright!
            im = im.transpose(Image.FLIP_TOP_BOTTOM)
        out_p = os.path.join(SPR_DIR, f'{name}.png')
        im.save(out_p)
        with open(out_p, 'rb') as bf:
            bundle_data[f'spr_{name}'] = 'data:image/png;base64,' + base64.b64encode(bf.read()).decode('ascii')
        print(f'Extracted sprite {name}: {im.width}x{im.height}')
    except Exception as e:
        print(f'Warn sprite {name}: {e}')

# 2. Extract authentic Castlevania background tileset sheets + Real Platform & Floor Hazard Textures!
stage_tiles = [
    ('lava_sheet0', '/bc/f_paa00.dat', read_ov_pal_at_offset(73, 0), 2),
    ('lava_sheet1', '/bc/f_paa01.dat', read_ov_pal_at_offset(73, 0), 4),
    ('glacier_sheet0', '/bc/f_raa08.dat', read_ov_pal_at_offset(76, 1004), 3),
    ('glacier_sheet1', '/bc/f_raa08.dat', read_ov_pal_at_offset(76, 1004), 5),
    # Real Clock Tower engraved gear mechanism mural (/bc/f_paa08.dat & /bc/f_paa09.dat)!
    ('clock_sheet0', '/bc/f_paa08.dat', read_ov_pal_at_offset(73, 0), 2),
    ('clock_sheet1', '/bc/f_paa09.dat', read_ov_pal_at_offset(73, 0), 2),
    ('clock_wall', '/bc/f_paa06.dat', read_ov_pal_at_offset(73, 0), 2),
    ('castle_sheet0', '/bc/f_raa08.dat', read_ov_pal_at_offset(76, 1004), 8),
    ('castle_sheet1', '/bc/f_raa08.dat', read_ov_pal_at_offset(76, 1004), 13),
    # Real floor hazard textures (Lava Magma, Glacier Ice Rapids)
    ('haz_lava', '/bc/f_ibb05.dat', read_ov_pal_at_offset(62, 6560), 2),
    ('haz_glacier', '/bc/f_ibb01.dat', read_ov_pal_at_offset(76, 1004), 3),
]

for name, bpath, pals, pidx in stage_tiles:
    sheet = decode_sheet_4bpp(files[bpath])
    pal = pals[pidx % len(pals)]
    im = Image.new('RGBA', (128, 128), (0, 0, 0, 0))
    px = im.load()
    for y in range(min(128, len(sheet))):
        for x in range(128):
            c = sheet[y][x]
            if c != 0:
                px[x, y] = pal[c]
    out_p = os.path.join(TILE_DIR, f'{name}.png')
    im.save(out_p)
    with open(out_p, 'rb') as bf:
        bundle_data[f'tile_{name}'] = 'data:image/png;base64,' + base64.b64encode(bf.read()).decode('ascii')
    print(f'Extracted tileset {name}: 128x128')

# Extract Real NDS Stone Platform Block Strips (96x28) & Real Spikes / Guillotine Hazard Sprites
def extract_sheet_crop(name, bpath, pals, pidx, crop_box):
    sheet = decode_sheet_4bpp(files[bpath])
    pal = pals[pidx % len(pals)]
    im = Image.new('RGBA', (128, 128), (0, 0, 0, 0))
    px = im.load()
    for y in range(min(128, len(sheet))):
        for x in range(128):
            c = sheet[y][x]
            if c != 0:
                px[x, y] = pal[c]
    cropped = im.crop(crop_box)
    out_p = os.path.join(TILE_DIR, f'{name}.png')
    cropped.save(out_p)
    with open(out_p, 'rb') as bf:
        bundle_data[f'tile_{name}'] = 'data:image/png;base64,' + base64.b64encode(bf.read()).decode('ascii')
    print(f'Extracted cropped texture {name}: {cropped.width}x{cropped.height}')

extract_sheet_crop('plat_lava', '/bc/f_paa0b.dat', read_ov_pal_at_offset(73, 0), 2, (0, 0, 96, 32))
extract_sheet_crop('plat_glacier', '/bc/f_paa0b.dat', read_ov_pal_at_offset(76, 1004), 3, (0, 0, 96, 32))
extract_sheet_crop('plat_clock', '/bc/f_paa0b.dat', read_ov_pal_at_offset(73, 0), 4, (0, 0, 96, 32))
extract_sheet_crop('plat_castle', '/bc/f_paa0b.dat', read_ov_pal_at_offset(76, 1004), 8, (0, 0, 96, 32))
# Additional distinct textures per platform role/staircase tier!
extract_sheet_crop('plat_gold_arch', '/bc/f_paa06.dat', read_ov_pal_at_offset(73, 0), 2, (0, 0, 96, 32))
extract_sheet_crop('plat_stair_brick', '/bc/f_paa06.dat', read_ov_pal_at_offset(73, 0), 4, (0, 32, 96, 64))
extract_sheet_crop('plat_iron_gear', '/bc/f_paa08.dat', read_ov_pal_at_offset(73, 0), 2, (0, 0, 96, 32))
extract_sheet_crop('plat_ice_crystal', '/bc/f_ibb01.dat', read_ov_pal_at_offset(76, 1004), 5, (0, 0, 96, 32))
extract_sheet_crop('plat_magma_rock', '/bc/f_ibb05.dat', read_ov_pal_at_offset(62, 6560), 2, (0, 0, 96, 32))
extract_sheet_crop('plat_cursed_trap', '/bc/f_raa08.dat', read_ov_pal_at_offset(76, 1004), 13, (0, 64, 96, 96))

extract_sheet_crop('spikes_ice', '/sc/f_spike00.dat', read_ov_pal_at_offset(75, 72140), 1, (0, 0, 64, 16))
extract_sheet_crop('spikes_steel', '/sc/f_spike00.dat', read_ov_pal_at_offset(73, 0), 2, (0, 0, 64, 16))
extract_sheet_crop('spikes_blood', '/sc/f_spike00.dat', read_ov_pal_at_offset(76, 1004), 13, (0, 0, 64, 16))
extract_sheet_crop('ice_thorns', '/sc/f_spike00.dat', read_ov_pal_at_offset(75, 72140), 1, (64, 64, 96, 96))
extract_sheet_crop('guillotine', '/sc/f_giloti.dat', read_ov_pal_at_offset(75, 72140), 1, (12, 0, 64, 96))

print('Graphics extraction complete!')

# 3. Generate Original Pixel-Art Student Protagonist (교복 + 고딕 망토 학생 도트 스프라이트)
# 4 frames: idle, run1, run2, attack
def make_student_sprites():
    palette = {
        '.': (0, 0, 0, 0),
        'K': (20, 18, 28, 255),      # Outline dark
        'H': (42, 36, 54, 255),      # Hair dark brown/black
        'h': (78, 66, 98, 255),      # Hair highlight
        'S': (255, 224, 196, 255),   # Skin light
        's': (224, 180, 148, 255),   # Skin shadow
        'W': (240, 244, 255, 255),   # School shirt white collar
        'R': (190, 38, 51, 255),     # School necktie / cape lining crimson
        'r': (130, 22, 36, 255),     # Dark crimson cape shadow
        'B': (38, 50, 86, 255),      # School blazer navy
        'b': (58, 76, 124, 255),     # Blazer highlight
        'G': (184, 152, 72, 255),    # Gold clasp / badge
        'P': (48, 52, 68, 255),      # School trousers dark gray
        'C': (92, 28, 64, 255),      # Gothic cape outer
    }
    # 24x32 pixel art grids for student in blazer + tie + flowing Gothic cape
    base_rows = [
        "........KKKKKKKK........",
        ".......KHHHHHHHHK.......",
        "......KHhhHHHHhhHK......",
        "......KHHHHHHHHHHK......",
        "......KHSSSSSSSSHK......",
        "......KSSKSSSSKSSK......",
        "......KSSSSSSSSSSK......",
        ".......KsssSSsssK.......",
        ".......KKWWGGWWKK.......",
        "......KCCKWRRWKBBK......",
        ".....KCCrKBBBBKbbBK.....",
        "....KCCrrKBRRBKbbBK.....",
        "....KCCrrKBBBBKSSSK.....",
        "...KCCCrrKBBBBKKKK......",
        "...KCCCrrKBBBBBK........",
        "..KCCCCrrKBBBBBK........",
        "..KCCCCrrKPPPPPK........",
        "..KCCCrrrKPPPPPK........",
        "...KCCrrKPPPPPPPK.......",
        "...KKKrrKPPKKPPPK.......",
        ".....KK.KPPK.KPPK.......",
        "........KPPK..KPPK......",
        "........KKKK..KKKK......",
        "........KWWK..KWWK......"
    ]
    for mode, name in enumerate(['student_idle', 'student_run1', 'student_run2', 'student_atk']):
        im = Image.new('RGBA', (24, 28), (0, 0, 0, 0))
        px = im.load()
        for y, row in enumerate(base_rows):
            for x, ch in enumerate(row):
                if ch in palette and palette[ch][3] > 0:
                    dx, dy = x, y + 2
                    if mode == 1 and y >= 18:
                        dx += -1 if x < 12 else 1
                    elif mode == 2 and y >= 18:
                        dx += 1 if x < 12 else -1
                    elif mode == 3 and y >= 9 and x >= 14:
                        dx += 1
                    if 0 <= dx < 24 and 0 <= dy < 28:
                        px[dx, dy] = palette[ch]
        out_p = os.path.join(SPR_DIR, f'{name}.png')
        im.save(out_p)
        with open(out_p, 'rb') as bf:
            bundle_data[f'spr_{name}'] = 'data:image/png;base64,' + base64.b64encode(bf.read()).decode('ascii')
    print('Generated original pixel-art student sprites!')

make_student_sprites()

# 4. Extract Real Castlevania NDS Audio (STRM & SWAR IMA-ADPCM -> WAV)
IMA_STEP_TABLE = [
    7, 8, 9, 10, 11, 12, 13, 14, 16, 17, 19, 21, 23, 25, 28, 31,
    34, 37, 41, 45, 50, 55, 60, 66, 73, 80, 88, 97, 107, 118, 130, 143,
    157, 173, 190, 209, 230, 253, 279, 307, 337, 371, 408, 449, 494, 544, 598, 658,
    724, 796, 876, 963, 1060, 1166, 1282, 1411, 1552, 1707, 1878, 2066, 2272, 2499, 2749, 3024,
    3327, 3660, 4026, 4428, 4871, 5358, 5894, 6484, 7132, 7845, 8630, 9493, 10442, 11487, 12635, 13899,
    15289, 16818, 18500, 20350, 22385, 24623, 27086, 29794, 32767
]
IMA_INDEX_TABLE = [-1, -1, -1, -1, 2, 4, 6, 8, -1, -1, -1, -1, 2, 4, 6, 8]

def decode_ima_adpcm(raw_bytes, num_samples=None):
    if len(raw_bytes) < 4:
        return np.zeros(0, dtype=np.int16)
    pred, step_idx = struct.unpack('<hH', raw_bytes[:4])
    step_idx = max(0, min(88, step_idx & 0x7F))
    out = []
    for b in raw_bytes[4:]:
        for nib in (b & 0x0F, (b >> 4) & 0x0F):
            step = IMA_STEP_TABLE[step_idx]
            diff = step >> 3
            if nib & 1: diff += step >> 2
            if nib & 2: diff += step >> 1
            if nib & 4: diff += step
            if nib & 8:
                pred = max(-32768, pred - diff)
            else:
                pred = min(32767, pred + diff)
            step_idx = max(0, min(88, step_idx + IMA_INDEX_TABLE[nib]))
            out.append(pred)
            if num_samples and len(out) >= num_samples:
                return np.array(out, dtype=np.int16)
    return np.array(out, dtype=np.int16)

# Locate SDAT (file 922)
sdat_st, sdat_ed = struct.unpack('<II', fat[922*8:923*8])
symb_off, symb_sz, info_off, info_sz, fat_s_off, fat_s_sz = struct.unpack('<IIIIII', rom[sdat_st+16:sdat_st+40])
sdat_info = rom[sdat_st+info_off : sdat_st+info_off+info_sz]
sdat_fat = rom[sdat_st+fat_s_off : sdat_st+fat_s_off+fat_s_sz]
rec_offs = struct.unpack('<8I', sdat_info[8:40])

def extract_strm_wav(strm_idx, out_name):
    strm_info_off = rec_offs[7]
    e_off = struct.unpack('<I', sdat_info[strm_info_off+4+strm_idx*4 : strm_info_off+8+strm_idx*4])[0]
    fid = struct.unpack('<H', sdat_info[e_off:e_off+2])[0]
    f_off, f_sz = struct.unpack('<II', sdat_fat[12 + fid*16 : 20 + fid*16])
    pos = sdat_st + f_off
    sh = rom[pos:pos+68]
    wave_type, loop, channels, pad, rate, time, loop_off, num_samples, data_off, num_blocks, blk_len, blk_samples, last_len, last_samples = struct.unpack('<BBBBHHIIIIIIII', sh[24:64])
    raw_data = rom[pos + data_off : pos + f_sz]
    samples = []
    off = 0
    for b_i in range(num_blocks):
        blen = last_len if b_i == num_blocks - 1 else blk_len
        bsamp = last_samples if b_i == num_blocks - 1 else blk_samples
        blk = raw_data[off : off + blen]
        off += blen * channels
        if wave_type == 2:
            samples.append(decode_ima_adpcm(blk, bsamp))
        elif wave_type == 1:
            samples.append(np.frombuffer(blk[:bsamp*2], dtype=np.int16))
    pcm = np.concatenate(samples) if samples else np.zeros(100, dtype=np.int16)
    # Downsample slightly to 11025Hz or keep original rate for compact size
    out_path = os.path.join(AUD_DIR, f'{out_name}.wav')
    with wave.open(out_path, 'wb') as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(rate)
        wf.writeframes(pcm.tobytes())
    with open(out_path, 'rb') as bf:
        bundle_data[f'aud_{out_name}'] = 'data:audio/wav;base64,' + base64.b64encode(bf.read()).decode('ascii')
    print(f'Extracted STRM[{strm_idx}] -> {out_name}.wav ({len(pcm)} samples @{rate}Hz)')

# Extract authentic Castlevania SFX from STRM
strm_jobs = [
    (7, 'sfx_glyph'),         # STRM_SYS_GRIF_GET2 (Glyph/Abstraction absorb sound!)
    (6, 'sfx_clear'),         # STRM_SYS_QUEST_GOOD (Stage clear fanfare!)
    (8, 'sfx_chest'),         # STRM_GM_TREASURE_BOX
    (11, 'sfx_door'),         # STRM_GM_DOOR_KANE (Castle iron gate open!)
    (24, 'sfx_dispel'),       # STRM_EV01_DISPEL (Noise shatter / illusion dispel!)
    (26, 'sfx_hit'),          # STRM_EV01_MAGIC_BURN (Damage / Fire hit!)
    (120, 'sfx_jump'),        # STRM_AL_J01 (Jump!)
    (122, 'sfx_slash'),       # STRM_AL_A01 (Attack slash!)
]
for sidx, sname in strm_jobs:
    extract_strm_wav(sidx, sname)

# Also extract crisp SFX from SWAR[0] (WAVE_SE_00) and SWAR[16] (WAVE_BGM_ALL)
def decode_swar(swar_idx):
    swar_info_off = rec_offs[3]
    e_off = struct.unpack('<I', sdat_info[swar_info_off+4+swar_idx*4 : swar_info_off+8+swar_idx*4])[0]
    fid = struct.unpack('<H', sdat_info[e_off:e_off+2])[0]
    f_off, f_sz = struct.unpack('<II', sdat_fat[12 + fid*16 : 20 + fid*16])
    pos = sdat_st + f_off
    nsamp = struct.unpack('<I', rom[pos+56:pos+60])[0]
    offs = struct.unpack(f'<{nsamp}I', rom[pos+60:pos+60+nsamp*4])
    swavs = []
    for i, soff in enumerate(offs):
        end_off = offs[i+1] if i+1 < nsamp else f_sz
        wdata = rom[pos+soff : pos+end_off]
        wtype, loop, rate, time, loop_pos, non_loop_len = struct.unpack('<BBHHHI', wdata[:12])
        raw = wdata[12:]
        if wtype == 2:
            pcm = decode_ima_adpcm(raw)
        elif wtype == 1:
            pcm = np.frombuffer(raw[:(len(raw)//2)*2], dtype=np.int16)
        else:
            pcm = (np.frombuffer(raw, dtype=np.int8).astype(np.int16) << 8)
        swavs.append((rate, loop, loop_pos*4, pcm))
    return swavs

swar0 = decode_swar(0)
# Save sword hit & candle break from SWAR[0]
for idx, name in [(2, 'sfx_candle'), (5, 'sfx_sword')]:
    if idx < len(swar0):
        rate, loop, lpos, pcm = swar0[idx]
        out_path = os.path.join(AUD_DIR, f'{name}.wav')
        with wave.open(out_path, 'wb') as wf:
            wf.setnchannels(1)
            wf.setsampwidth(2)
            wf.setframerate(rate)
            wf.writeframes(pcm.tobytes())
        with open(out_path, 'rb') as bf:
            bundle_data[f'aud_{name}'] = 'data:audio/wav;base64,' + base64.b64encode(bf.read()).decode('ascii')
        print(f'Extracted SWAR0[{idx}] -> {name}.wav')

# 5. Synthesize 4 Authentic Stage BGMs using SSEQ notes + SWAR[16] (WAVE_BGM_ALL) real Castlevania instrument samples!
bgm_swavs = decode_swar(16)
print(f'Decoded {len(bgm_swavs)} real Castlevania BGM instrument samples from SWAR[16]!')

def render_sseq_bgm(sseq_idx, bfid, out_name, duration_sec=24.0):
    sseq_info_off = rec_offs[0]
    e_off = struct.unpack('<I', sdat_info[sseq_info_off+4+sseq_idx*4 : sseq_info_off+8+sseq_idx*4])[0]
    fid = struct.unpack('<H', sdat_info[e_off:e_off+2])[0]
    f_off, f_sz = struct.unpack('<II', sdat_fat[12 + fid*16 : 20 + fid*16])
    b_off, b_sz = struct.unpack('<II', sdat_fat[12 + bfid*16 : 20 + bfid*16])

    sseq_bytes = rom[sdat_st + f_off : sdat_st + f_off + f_sz]
    sbnk_bytes = rom[sdat_st + b_off : sdat_st + b_off + b_sz]

    # Parse SBNK instrument -> (swav_id, base_note)
    num_inst = struct.unpack('<I', sbnk_bytes[56:60])[0]
    inst_map = {}
    for i in range(num_inst):
        irec = sbnk_bytes[60 + i*4 : 64 + i*4]
        itype = irec[0]
        ioff = struct.unpack('<H', irec[1:3])[0]
        if itype == 1 and ioff > 0 and ioff + 10 <= len(sbnk_bytes):
            swav_id, swar_id, base_note = struct.unpack('<HHB', sbnk_bytes[ioff:ioff+5])
            inst_map[i] = (swav_id, base_note, False)
        elif itype == 16 and ioff > 0 and ioff + 14 <= len(sbnk_bytes):
            swav_id, swar_id, base_note = struct.unpack('<HHB', sbnk_bytes[ioff+4:ioff+9])
            inst_map[i] = (swav_id, base_note, True)
        elif itype == 17 and ioff > 0 and ioff + 20 <= len(sbnk_bytes):
            swav_id, swar_id, base_note = struct.unpack('<HHB', sbnk_bytes[ioff+10:ioff+15])
            inst_map[i] = (swav_id, base_note, False)

    # Parse SSEQ tracks
    data_start = struct.unpack('<I', sseq_bytes[24:28])[0]
    seq = sseq_bytes[data_start:]
    track_offs = [0]
    pos = 0
    if pos < len(seq) and seq[pos] == 0xFE:
        pos += 3
    while pos < len(seq) and seq[pos] == 0x93:
        t_idx = seq[pos+1]
        t_off = seq[pos+2] | (seq[pos+3] << 8) | (seq[pos+4] << 16)
        track_offs.append(t_off)
        pos += 5
    track_offs[0] = pos

    def read_vlq(buf, p):
        v = 0
        while p < len(buf):
            b = buf[p]
            p += 1
            v = (v << 7) | (b & 0x7F)
            if not (b & 0x80):
                break
        return v, p

    out_rate = 16000
    total_samples = int(duration_sec * out_rate)
    mix = np.zeros(total_samples, dtype=np.float32)
    tempo = 120.0

    for t_i, start_p in enumerate(track_offs[:8]):
        p = start_p
        cur_tick = 0
        prog = 0
        vol = 90
        call_stack = []
        max_steps = 4000
        steps = 0
        while p < len(seq) and steps < max_steps:
            steps += 1
            tick_sec = 60.0 / (tempo * 48.0)
            cur_time = cur_tick * tick_sec
            if cur_time >= duration_sec:
                break
            cmd = seq[p]
            p += 1
            if cmd < 0x80:
                vel = seq[p]
                p += 1
                dur, p = read_vlq(seq, p)
                # Render note cmd with instrument prog
                swav_id, base_note, is_drum = inst_map.get(prog, (min(prog, len(bgm_swavs)-1), 60, False))
                if swav_id < len(bgm_swavs):
                    s_rate, s_loop, s_lpos, s_pcm = bgm_swavs[swav_id]
                    if len(s_pcm) > 32:
                        ratio = (2.0 ** ((cmd - base_note) / 12.0)) * (s_rate / out_rate) if not is_drum else (s_rate / out_rate)
                        ratio = max(0.2, min(4.0, ratio))
                        note_len = max(int(dur * tick_sec * out_rate), int(0.08 * out_rate))
                        note_len = min(note_len, int(1.5 * out_rate))
                        start_idx = int(cur_time * out_rate)
                        if start_idx < total_samples:
                            end_idx = min(total_samples, start_idx + note_len)
                            n_out = end_idx - start_idx
                            src_indices = (np.arange(n_out) * ratio).astype(np.int32)
                            if s_loop and len(s_pcm) > s_lpos + 16:
                                loop_len = len(s_pcm) - s_lpos
                                mask = src_indices >= len(s_pcm)
                                src_indices[mask] = s_lpos + ((src_indices[mask] - s_lpos) % loop_len)
                            else:
                                src_indices = np.clip(src_indices, 0, len(s_pcm) - 1)
                            env = np.linspace(1.0, 0.25, n_out, dtype=np.float32)
                            gain = (vel / 127.0) * (vol / 127.0) * 0.25
                            mix[start_idx:end_idx] += (s_pcm[src_indices].astype(np.float32) / 32768.0) * env * gain
            elif cmd == 0x80:
                wait_t, p = read_vlq(seq, p)
                cur_tick += wait_t
            elif cmd == 0x81:
                prog, p = read_vlq(seq, p)
            elif cmd == 0x93:
                p += 4
            elif cmd == 0x94:
                dest = seq[p] | (seq[p+1] << 8) | (seq[p+2] << 16)
                p = dest
            elif cmd == 0x95:
                dest = seq[p] | (seq[p+1] << 8) | (seq[p+2] << 16)
                call_stack.append(p + 3)
                p = dest
            elif cmd == 0xFD:
                if call_stack:
                    p = call_stack.pop()
                else:
                    break
            elif cmd == 0xFF:
                break
            elif cmd == 0xE1:
                tempo = max(60.0, float(struct.unpack('<H', seq[p:p+2])[0]))
                p += 2
            elif cmd in (0xC1, 0xC2, 0xC3, 0xC4, 0xC5, 0xC6, 0xC7, 0xC8, 0xC9, 0xCA, 0xCB, 0xCC, 0xCD, 0xCE, 0xCF, 0xD0, 0xD1, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6):
                if cmd == 0xC1:
                    vol = seq[p]
                p += 1
            elif cmd in (0xE0, 0xE2, 0xE3):
                p += 2
            else:
                p += 1

    # Normalize and write WAV
    peak = np.max(np.abs(mix))
    if peak > 0.01:
        mix = (mix / peak) * 0.85
    pcm16 = (mix * 32767.0).astype(np.int16)
    out_path = os.path.join(AUD_DIR, f'{out_name}.wav')
    with wave.open(out_path, 'wb') as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(out_rate)
        wf.writeframes(pcm16.tobytes())
    with open(out_path, 'rb') as bf:
        bundle_data[f'aud_{out_name}'] = 'data:audio/wav;base64,' + base64.b64encode(bf.read()).decode('ascii')
    print(f'Synthesized SSEQ[{sseq_idx}] -> {out_name}.wav ({duration_sec}s @{out_rate}Hz)')

# Render the 4 Stage BGMs from the NDS ROM:
# Map 1 (Lava): SSEQ[9] (SEQ_010_SKELTON, bfid=86)
# Map 2 (Glacier): SSEQ[11] (SEQ_012_MOUNT2, bfid=88)
# Map 3 (Clock Tower): SSEQ[18] (SEQ_018_TOWER, bfid=89)
# Map 4 (Final Castle Escape): SSEQ[19] (SEQ_019_CASTLE, bfid=93)
render_sseq_bgm(9, 86, 'bgm_lava', 20.0)
render_sseq_bgm(11, 88, 'bgm_glacier', 20.0)
render_sseq_bgm(18, 89, 'bgm_clock', 20.0)
render_sseq_bgm(19, 93, 'bgm_castle', 20.0)
render_sseq_bgm(47, 117, 'bgm_gameover', 10.0)

# Save JS bundle so index.html can load all extracted NDS assets with zero CORS issues
bundle_js_path = os.path.join(BASE_DIR, 'assets', 'assets_bundle.js')
with open(bundle_js_path, 'w', encoding='utf-8') as f:
    f.write('window.NDS_ASSETS = ' + json.dumps(bundle_data, ensure_ascii=False) + ';\n')
print(f'Saved assets_bundle.js with {len(bundle_data)} assets!')
