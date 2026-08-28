"""Compare a reconstructed .py against the original .pyc, function by function.

Compiles the reconstruction with 3.12 and walks both code trees in parallel,
comparing the things that must match if the source is faithful: the opcode
sequence, the names touched, and the constants.
"""
import sys, marshal, types, dis, difflib

def load_pyc(p):
    return marshal.loads(open(p,'rb').read()[16:])

def compile_py(p):
    return compile(open(p).read(), p, 'exec')

def index(code, out=None, prefix=""):
    """Map qualified-ish name -> code object, for every nested code object."""
    if out is None: out = {}
    for c in code.co_consts:
        if isinstance(c, types.CodeType):
            key = f"{prefix}{c.co_name}:{c.co_firstlineno}"
            n, i = key, 2
            while n in out: n, i = f"{key}#{i}", i+1
            out[n] = c
            index(c, out, prefix + c.co_name + ".")
    return out

def ops(code):
    """Opcode stream, ignoring caches, line numbers and jump deltas."""
    return [i.opname for i in dis.get_instructions(code) if i.opname != 'CACHE']

def consts(code):
    return [repr(c) for c in code.co_consts if not isinstance(c, types.CodeType)]

def score(a, b):
    return difflib.SequenceMatcher(None, a, b).ratio()

def main(recon, original):
    new, old = compile_py(recon), load_pyc(original)
    ni, oi = index(new), index(old)
    # match on bare function name, ignoring line numbers, so a shifted
    # reconstruction still lines up with its original
    def by_name(idx):
        d = {}
        for k, c in idx.items():
            d.setdefault(c.co_name, []).append(c)
        return d
    nb, ob = by_name(ni), by_name(oi)

    print(f"{'function':<40} {'ops':>6} {'consts':>7}  status")
    print("-"*72)
    perfect = partial = missing = 0
    for name in sorted(ob):
        olds = ob[name]
        news = nb.get(name, [])
        for k, oc in enumerate(olds):
            if k >= len(news):
                print(f"{name:<40} {'':>6} {'':>7}  MISSING")
                missing += 1
                continue
            nc = news[k]
            so = score(ops(nc), ops(oc))
            sc = score(consts(nc), consts(oc))
            tag = "ok" if so == 1.0 and sc == 1.0 else ("close" if so > .92 else "DIFF")
            if tag == "ok": perfect += 1
            elif tag == "close": partial += 1
            print(f"{name:<40} {so:>6.3f} {sc:>7.3f}  {tag}")
    extra = sum(len(v) for v in nb.values()) - sum(len(v) for v in ob.values())
    print("-"*72)
    print(f"exact {perfect}   close {partial}   missing {missing}   extra {max(0,extra)}")
    # module level
    print(f"\nmodule body: ops {score(ops(new), ops(old)):.3f}  consts {score(consts(new), consts(old)):.3f}")

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
