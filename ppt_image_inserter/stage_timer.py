"""Time each stage of a build and flag a stage far over its expected cost.

FF's rule (global CLAUDE.md): pipelines log time per item and flag any item far over its expected cost.

    timer = StageTimer({"posters": (0.0, 0.035)})   # expected seconds, from a measured run
    with timer.stage("PowerPoint build and save"):
        ...
    timer.report()

Each stage prints one line when it ends, so a log shows where a long build is. A stage is flagged FAR OVER when it
takes more than ``factor`` times its expected time plus ``slack_s`` (short stages jitter). A stage with no expected
time is logged as "expected: unmeasured".

``expected`` keys are stage-name prefixes (the longest matching prefix wins). A value is seconds, or (base, per_item):
base + per_item x the first number in the stage name, e.g. "posters (398 movies)" with (0.0, 0.035) expects 13.9 s.
"""
import re
import time
from contextlib import contextmanager
from typing import Dict, List, Optional, Tuple


class StageTimer:
    """Wall-clock time per named stage, printed as each stage ends and as a table at the end."""

    def __init__(self, expected: Optional[Dict[str, object]] = None, factor: float = 3.0, slack_s: float = 10.0):
        """expected: {stage-name prefix: seconds or (base, per_item)} from measured runs; factor, slack_s: the FAR
        OVER threshold."""
        self.expected = dict(expected or {})
        self.factor, self.slack_s = factor, slack_s
        self.times: List[Tuple[str, float]] = []
        self.t0 = time.perf_counter()

    def expected_s(self, name: str) -> Optional[float]:
        keys = [k for k in self.expected if name.startswith(k)]
        if not keys:
            return None
        value = self.expected[max(keys, key=len)]
        if isinstance(value, (tuple, list)):
            m = re.search(r"\d+", name)
            return value[0] + value[1] * (int(m.group()) if m else 0)
        return float(value)

    def _verdict(self, name: str, seconds: float) -> str:
        exp = self.expected_s(name)
        if exp is None:
            return "expected: unmeasured"
        flag = "  FAR OVER" if seconds > self.factor * exp + self.slack_s else ""
        return f"expected ~{exp:.0f} s{flag}"

    @contextmanager
    def stage(self, name: str):
        t = time.perf_counter()
        try:
            yield
        finally:
            dt = time.perf_counter() - t
            self.times.append((name, dt))
            print(f"STAGE {name}: {dt:.1f} s ({self._verdict(name, dt)})", flush=True)

    def far_over(self) -> List[str]:
        return [n for n, dt in self.times
                if self.expected_s(n) is not None and dt > self.factor * self.expected_s(n) + self.slack_s]

    def report(self) -> None:
        total = time.perf_counter() - self.t0
        width = max([len(n) for n, _ in self.times] + [5])
        print("STAGE TIMES")
        for n, dt in self.times:
            print(f"  {n:<{width}}  {dt:7.1f} s  {self._verdict(n, dt)}")
        print(f"  {'total':<{width}}  {total:7.1f} s" + (f"  FAR OVER: {', '.join(self.far_over())}"
                                                          if self.far_over() else ""), flush=True)
