"""Launch several `build_cellpose_qc_deck.py` jobs in parallel, one per YAML.

Each condition becomes its own Python subprocess. Multiple SMB connections
from separate processes typically scale aggregate throughput off `Y:\\`
much better than threads inside one process.

Usage:
    python run_conditions_parallel.py [config1.yaml ...]

If no configs are given, defaults to every `configs/cart_*.yaml` next to
this script. Logs are written to ``logs/<config-stem>.log``.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path
from typing import List

HERE = Path(__file__).resolve().parent
DEFAULT_CONFIG_GLOB = "configs/cart_*.yaml"
PYTHON = sys.executable
BUILD_SCRIPT = HERE / "build_cellpose_qc_deck.py"
# Logs live alongside the generated decks, NOT in the repo. Override with
# --log-dir if you want them somewhere else.
DEFAULT_LOG_DIR = Path("K:/FF/PPT/PPT_autogeneration/CART_actin_only/cellpose_qc/logs")


def discover_configs(args_configs: List[str]) -> List[Path]:
    if args_configs:
        paths = [Path(c).resolve() for c in args_configs]
    else:
        paths = sorted((HERE).glob(DEFAULT_CONFIG_GLOB))
    bad = [p for p in paths if not p.exists()]
    if bad:
        raise SystemExit(f"Missing config(s): {bad}")
    return paths


def main(argv: List[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("configs", nargs="*", help="YAML configs to run")
    parser.add_argument(
        "--log-dir", default=str(DEFAULT_LOG_DIR),
        help=f"Directory for per-condition log files (default: {DEFAULT_LOG_DIR})",
    )
    args = parser.parse_args(argv)

    configs = discover_configs(args.configs)
    log_dir = Path(args.log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)

    print(f"Launching {len(configs)} condition(s) in parallel:")
    for c in configs:
        print(f"  - {c}")
    print(f"Logs: {log_dir}")
    print()

    t0 = time.time()
    procs = []
    for cfg in configs:
        log_path = log_dir / f"{cfg.stem}.log"
        log_f = open(log_path, "w")
        cmd = [PYTHON, str(BUILD_SCRIPT), str(cfg)]
        p = subprocess.Popen(cmd, stdout=log_f, stderr=subprocess.STDOUT)
        procs.append((cfg, p, log_f, log_path, time.time()))
        print(f"  started {cfg.name}  (pid {p.pid}, log {log_path.name})")

    print(f"\nWaiting on {len(procs)} subprocess(es)...")
    results = []
    remaining = set(range(len(procs)))
    while remaining:
        for idx in list(remaining):
            cfg, p, log_f, log_path, started = procs[idx]
            if p.poll() is not None:
                log_f.close()
                elapsed = time.time() - started
                results.append((cfg, p.returncode, elapsed, log_path))
                status = "OK" if p.returncode == 0 else f"FAIL({p.returncode})"
                print(f"  [{status}] {cfg.name}  in {elapsed/60:.1f} min")
                remaining.discard(idx)
        if remaining:
            time.sleep(2.0)

    wall = time.time() - t0
    print(f"\nAll done in {wall/60:.1f} min wall.")
    print("Per-condition timings:")
    for cfg, rc, el, lp in sorted(results, key=lambda r: r[0].name):
        status = "OK" if rc == 0 else f"FAIL({rc})"
        print(f"  [{status}] {cfg.name:50s} {el/60:5.1f} min  log: {lp}")

    return 0 if all(r[1] == 0 for r in results) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
