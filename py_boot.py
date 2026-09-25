"""
py_boot.py - makes the browser's Python behave like desktop Python for the
Gaeltec tools. Imported once, before any tool module.

 * rapidfuzz: the browser has no compiled rapidfuzz, so its own built-in
   pure-Python implementation is used (same results, tested on 4,000 pairs).
 * threads: browser Python can't start threads. The tools read PDFs and
   scan folders with a ThreadPoolExecutor; here it simply runs each job
   straight away, one after another - same results, just not in parallel.
"""
import concurrent.futures as _cf
import os

os.environ.setdefault("RAPIDFUZZ_IMPLEMENTATION", "python")


class SequentialExecutor:
    """Drop-in for ThreadPoolExecutor that runs every task immediately."""

    def __init__(self, max_workers=None, *args, **kwargs):
        pass

    def submit(self, fn, *args, **kwargs):
        fut = _cf.Future()
        try:
            fut.set_result(fn(*args, **kwargs))
        except BaseException as e:  # same as a worker thread: error goes on the future
            fut.set_exception(e)
        return fut

    def map(self, fn, *iterables, **kwargs):
        return [fn(*a) for a in zip(*iterables)]

    def shutdown(self, wait=True, **kwargs):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


_cf.ThreadPoolExecutor = SequentialExecutor
