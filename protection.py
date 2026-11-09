from pathlib import Path
import io, re, struct, zipfile, hashlib, json
import olefile

def records(data):
    # Yield MS-XLSB records as (record type, original bytes including the header).
    # The type uses up to two bytes; length is encoded in groups of seven bits.
    # Keeping the original bytes avoids re-encoding records we do not change.
    pos = 0
    while pos < len(data):
        start = pos
        typ = data[pos]; pos += 1
        if typ & 128:
            typ = (typ & 127) | (data[pos] << 7); pos += 1
        size = 0
        for shift in (0, 7, 14, 21):
            b = data[pos]; pos += 1
            size |= (b & 127) << shift
            if not b & 128:
                break
        else:
            raise ValueError('Invalid record length')
        end = pos + size
        assert end <= len(data)
        yield typ, data[start:end]
        pos = end

def decode(value):
    # MS-OVBA 2.4.3.3: decode the obfuscation wrapper around a PROJECT value.
    # This does not recover the old password; the payload may be a password hash.
    raw = bytes.fromhex(value.decode())
    seed, ver, key = raw[:3]
    assert seed ^ ver == 2
    plain, e1, e2 = seed ^ key, key, ver
    out = bytearray()
    for enc in raw[3:]:
        b = enc ^ ((e2 + plain) & 255)
        out.append(b)
        plain, e2, e1 = b, e1, enc
    # Skip the seed-dependent padding, then read the four-byte payload length.
    out = out[(seed & 6) // 2:]
    size = int.from_bytes(out[:4], 'little')
    assert size == len(out) - 4
    return bytes(out[4:])

def encode(data, original):
    # MS-OVBA 2.4.3.2: encode a new value using the original seed and project key.
    # https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/1ad481e0-7df4-4cac-a9a4-9c29a1340123
    raw = bytes.fromhex(original.decode())
    seed, ver, key = raw[:3]
    plain, e1, e2 = seed ^ key, key, ver
    out = bytearray(raw[:3])
    payload = bytes((seed & 6) // 2) + struct.pack('<I', len(data)) + data
    for b in payload:
        enc = b ^ ((e2 + plain) & 255)
        out.append(enc)
        plain, e2, e1 = b, e1, enc
    result = out.hex().upper().encode()
    assert decode(result) == data
    return result

def unlock_vba(data):
    # vbaProject.bin is an OLE compound file. Only its PROJECT stream is edited.
    stream = io.BytesIO(data)
    with olefile.OleFileIO(stream, write_mode=True) as ole:
        original = ole.openstream('PROJECT').read()
        updated = original
        # CMG: unprotected; DPB: no password; GC: project is visible.
        for name, value in [(b'CMG', bytes(4)), (b'DPB', b'\0'), (b'GC', b'\xff')]:
            pattern = rb'(?m)^' + name + rb'="([0-9A-Fa-f]+)"\r?$'
            matches = list(re.finditer(pattern, updated))
            assert len(matches) == 1
            match = matches[0]
            decode(match[1])
            replacement = encode(value, match[1])
            updated = updated[:match.start(1)] + replacement + updated[match.end(1):]
        # olefile writes streams without resizing them. Pad the shortened PROJECT
        # stream with trailing blank lines to preserve its original byte length.
        padding = len(original) - len(updated)
        assert padding >= 0 and padding % 2 == 0
        updated += b'\r\n' * (padding // 2)
        ole.write_stream('PROJECT', updated)
    result = stream.getvalue()
    # Code, forms, and other streams must remain byte-for-byte identical.
    # Decode the new settings to verify them; this does not test macro behavior.
    with olefile.OleFileIO(io.BytesIO(data)) as old, olefile.OleFileIO(io.BytesIO(result)) as new:
        assert old.listdir() == new.listdir()
        for path in old.listdir():
            if path != ['PROJECT']:
                assert old.openstream(path).read() == new.openstream(path).read()
        project = new.openstream('PROJECT').read()
        for name, value in [(b'CMG', bytes(4)), (b'DPB', b'\0'), (b'GC', b'\xff')]:
            assert decode(re.search(rb'(?m)^' + name + rb'="([0-9A-F]+)"', project)[1]) == value
    return result


