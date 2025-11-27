"""
Wrapper that delegates to `scripts/temp/inject_deps_path.py`.
"""
import sys
import os
here = os.path.dirname(os.path.abspath(__file__))
delegate = os.path.join(here, 'temp', 'inject_deps_path.py')
if os.path.exists(delegate):
    # invoke the moved script with same arguments
    os.execv(sys.executable, [sys.executable, delegate] + sys.argv[1:])
else:
    # fallback: echo cwd
    if len(sys.argv) > 1:
        print(sys.argv[1])
    else:
        print(os.getcwd())
