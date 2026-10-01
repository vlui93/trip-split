import zlib, struct, math

STOPS = [(0.00, (0x5B,0x8C,0xFF)), (0.55, (0x2B,0x5B,0xE8)), (1.00, (0x1B,0x36,0xB8))]

def grad(t):
    t = max(0.0, min(1.0, t))
    for i in range(len(STOPS)-1):
        t0,c0 = STOPS[i]; t1,c1 = STOPS[i+1]
        if t0 <= t <= t1:
            k = 0 if t1==t0 else (t-t0)/(t1-t0)
            return tuple(c0[j] + (c1[j]-c0[j])*k for j in range(3))
    return STOPS[-1][1]

def render(size, ss=4):
    """Supersampled render. Full bleed, opaque, no corner rounding — iOS masks it."""
    S = size*ss
    cx = cy = S/2.0
    r  = S*0.295
    gap = 0.075
    off = S*0.018
    wedges = []
    for i in range(3):
        a0 = -math.pi/2 + i*(2*math.pi/3) + gap
        a1 = -math.pi/2 + (i+1)*(2*math.pi/3) - gap
        mid = (a0+a1)/2.0
        wedges.append((a0, a1, cx + math.cos(mid)*off, cy + math.sin(mid)*off))

    hx, hy, hr = S*0.28, S*0.22, S*0.75
    acc = [[[0.0,0.0,0.0] for _ in range(size)] for _ in range(size)]

    for py in range(S):
        oy = py//ss
        for px in range(S):
            # background gradient along the diagonal
            cr,cg,cb = grad((px+py)/(2.0*S))
            # soft top-left highlight
            d = math.hypot(px-hx, py-hy)
            if d < hr:
                a = 0.22*(1.0 - d/hr)
                cr = cr + (255-cr)*a; cg = cg + (255-cg)*a; cb = cb + (255-cb)*a
            # white wedges
            for a0,a1,wx,wy in wedges:
                dx, dy = px-wx, py-wy
                if dx*dx+dy*dy <= r*r:
                    ang = math.atan2(dy, dx)
                    while ang < a0: ang += 2*math.pi
                    if ang <= a1:
                        cr=cg=cb=255.0
                        break
            cell = acc[oy][px//ss]
            cell[0]+=cr; cell[1]+=cg; cell[2]+=cb

    n = float(ss*ss)
    rows = []
    for y in range(size):
        row = bytearray()
        for x in range(size):
            c = acc[y][x]
            row += bytes((int(c[0]/n+0.5), int(c[1]/n+0.5), int(c[2]/n+0.5)))
        rows.append(bytes(row))
    return rows

def write_png(path, rows, size):
    raw = b''.join(b'\x00'+r for r in rows)          # filter type 0 per row
    def chunk(t, d):
        return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t+d) & 0xffffffff)
    png = (b'\x89PNG\r\n\x1a\n'
           + chunk(b'IHDR', struct.pack('>IIBBBBB', size, size, 8, 2, 0, 0, 0))  # colortype 2 = RGB, no alpha
           + chunk(b'IDAT', zlib.compress(raw, 9))
           + chunk(b'IEND', b''))
    open(path,'wb').write(png)
    return len(png)

# Bump this whenever the artwork changes. iOS caches home-screen icons by URL and
# will not refetch the same filename, even if you delete and re-add the WebClip —
# so the version has to be in the name, not just the bytes.
VERSION = 2

for s in (180, 167, 152, 120, 512, 192, 32):
    ss = 4 if s <= 200 else 2
    rows = render(s, ss)
    name = f'icon-{s}-v{VERSION}.png'
    n = write_png(name, rows, s)
    print(f'  {name}  {n:>7} bytes')
