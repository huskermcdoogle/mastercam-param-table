"""
Takes the builder's Windows user name out of res\\vbaProject.bin.

A project with a UserForm references MSForms through Excel's cached
%TEMP%\\VBE\\MSForms.exd, so the compiled project carries
C:\\Users\\<name>\\AppData\\Local\\Temp\\VBE\\MSForms.exd - in the _VBA_PROJECT
and __SRP_ caches as plain text and in the compressed "dir" stream. That path
is only a cache hint (Excel rebuilds the .exd where it likes), so the name is
overwritten in place with a placeholder of the same length; nothing in the
file moves. In "dir" the bytes are patched where the compressor stored them
as literals, then the stream is decompressed again to prove it reads back.

Usage: python tools\\scrub_vba.py <vbaProject.bin> <user name>
"""
import struct
import sys

END = 0xFFFFFFFA


def streams(d):
    """{name: list of file offsets, one per stream byte}"""
    ss = 1 << struct.unpack_from('<H', d, 30)[0]
    mss = 1 << struct.unpack_from('<H', d, 32)[0]
    dirstart, = struct.unpack_from('<I', d, 0x30)
    cutoff, = struct.unpack_from('<I', d, 0x38)
    minifat_start, = struct.unpack_from('<I', d, 0x3c)
    ndifat, = struct.unpack_from('<I', d, 0x48)
    if ndifat:
        raise SystemExit('scrub_vba: DIFAT sectors not supported (file too big)')
    sec = lambda n: 512 + n * ss
    fat = []
    for s in struct.unpack_from('<109I', d, 0x4c):
        if s < END:
            fat += struct.unpack_from('<%dI' % (ss // 4), d, sec(s))

    def chain(s, table):
        out = []
        while s < END:
            out.append(s)
            s = table[s]
        return out

    dirdata = b''.join(d[sec(s):sec(s) + ss] for s in chain(dirstart, fat))
    minifat = []
    for s in chain(minifat_start, fat):
        minifat += struct.unpack_from('<%dI' % (ss // 4), d, sec(s))
    rootchain = None
    res = {}
    for i in range(len(dirdata) // 128):
        e = dirdata[i * 128:(i + 1) * 128]
        nl, = struct.unpack_from('<H', e, 64)
        name = e[:max(0, nl - 2)].decode('utf-16-le')
        typ = e[66]
        start, size = struct.unpack_from('<IQ', e, 116)
        size &= 0xFFFFFFFF
        if typ == 5:
            rootchain = chain(start, fat)
            continue
        if typ != 2:
            continue
        offs = []
        if size < cutoff:
            for m in chain(start, minifat):
                o = m * mss
                base = sec(rootchain[o // ss]) + o % ss
                offs += range(base, base + mss)
        else:
            for s in chain(start, fat):
                offs += range(sec(s), sec(s) + ss)
        res[name] = offs[:size]
    return res


def decompress(c):
    """MS-OVBA 2.4.1: the text, and for each output byte the input index it
    was a literal at (None when it came from a copy token)."""
    if c[0] != 1:
        raise ValueError('not a compressed container')
    out, src = bytearray(), []
    i = 1
    while i < len(c):
        hdr, = struct.unpack_from('<H', c, i)
        end = i + (hdr & 0xFFF) + 3
        i += 2
        start = len(out)
        if not hdr & 0x8000:
            out += c[i:i + 4096]
            src += range(i, i + 4096)
            i += 4096
            continue
        while i < end:
            flags = c[i]
            i += 1
            for b in range(8):
                if i >= end:
                    break
                if not (flags >> b) & 1:
                    out.append(c[i])
                    src.append(i)
                    i += 1
                else:
                    tok, = struct.unpack_from('<H', c, i)
                    i += 2
                    bc = max((len(out) - start - 1).bit_length(), 4)
                    n = (tok & (0xFFFF >> bc)) + 3
                    off = (tok >> (16 - bc)) + 1
                    for _ in range(n):
                        out.append(out[-off])
                        src.append(None)
    return bytes(out), src


def hits(b, name):
    low = b.lower()
    return low.count(name.lower().encode('latin-1')) + low.count(name.lower().encode('utf-16-le'))


def main():
    path, name = sys.argv[1], sys.argv[2]
    if len(name) < 3:
        raise SystemExit('scrub_vba: user name too short to scrub safely')
    stand_in = ('user' + 'x' * len(name))[:len(name)]
    d = bytearray(open(path, 'rb').read())
    patched = 0
    for sname, offs in streams(bytes(d)).items():
        data = bytes(d[o] for o in offs)
        if sname == 'dir':
            text, src = decompress(data)
            for enc in ('latin-1', 'utf-16-le'):
                needle, repl = name.lower().encode(enc), stand_in.encode(enc)
                at = text.lower().find(needle)
                while at >= 0:
                    for k in range(len(needle)):
                        if src[at + k] is not None:
                            d[offs[src[at + k]]] = repl[k]
                        # a copy token repeats bytes already patched earlier
                    patched += 1
                    at = text.lower().find(needle, at + 1)
        else:
            low = data.lower()
            for enc in ('latin-1', 'utf-16-le'):
                needle, repl = name.lower().encode(enc), stand_in.encode(enc)
                at = low.find(needle)
                while at >= 0:
                    for k in range(len(needle)):
                        d[offs[at + k]] = repl[k]
                    patched += 1
                    at = low.find(needle, at + 1)
    # Prove it: no stream, raw or decompressed dir, still carries the name.
    left = 0
    for sname, offs in streams(bytes(d)).items():
        data = bytes(d[o] for o in offs)
        left += hits(data, name)
        if sname == 'dir':
            left += hits(decompress(data)[0], name)
    if left:
        raise SystemExit('scrub_vba: %d copies of the user name remain' % left)
    open(path, 'wb').write(d)
    print('scrub_vba: %d copies of the user name replaced' % patched)


main()
