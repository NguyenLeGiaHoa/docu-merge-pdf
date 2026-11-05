"""
DocuMergePDF - Master Quality Assurance & Test Pipeline Runner
Executes all 11 Supreme Production Testing Tiers:
  Tier 1: Unit Testing (TemplateInspector, DataReader, DynamicNamer, Merger, PDF, Licensing, Legal)
  Tier 2: Integration Testing (Real DOCX+XLSX pipeline, Split-runs, Multilingual Unicode, Preview)
  Tier 3: System Testing (50-Recipient Batch Stress, Zero-Data-Loss, Hygiene, Speed Benchmark)
  Tier 4: Acceptance & Adversarial (Trial limits, clock tamper, reserved names, traversal, fuzzing)
  Tier 5: Headless GUI & UX Testing (Window lifecycle, tabs, widgets, formula builder, trial badge)
  Tier 6: Security, Privacy & Local-First Invariants (Zero network leaks, ciphertext hygiene)
  Tier 7: Legal Compliance & Immunity Shield (UCC § 2-316, UCC § 2-719, $1.00 ceiling, trademarks)
  Tier 8: Packaging & Artifacts Verification (PE Binary, MSIX Container, Gumroad Bundles, Assets)
  Tier 9: Standalone PE Executable Process Smoke & RAM Stability (Process launch, RSS memory, clean exit)
  Tier 10: Deep Vector PDF Forensic Inspection & Typography (PyMuPDF vector stream analysis, zero tag leaks)
  Tier 11: Real-World Dual Dataset Production Run (10 Academic Certificates + 10 HR Payroll Payslips)
"""

import os
import sys
import time
import unittest

# Reconfigure stdout for utf-8 on Windows
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

root_dir = os.path.dirname(os.path.abspath(__file__))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)


def run_test_tier(name: str, module_path: str):
    print(f"\n{'='*75}")
    print(f" ▶ EXECUTING {name.upper()}")
    print(f"{'='*75}")
    start = time.perf_counter()

    loader = unittest.TestLoader()
    suite = loader.loadTestsFromName(module_path)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    elapsed = time.perf_counter() - start
    return {
        "tier": name,
        "total": result.testsRun,
        "failures": len(result.failures),
        "errors": len(result.errors),
        "skipped": len(result.skipped),
        "time": elapsed,
        "success": result.wasSuccessful(),
    }


def main():
    print("*" * 75)
    print("     DOCUMERGEPDF - 11-TIER SUPREME QUALITY ASSURANCE & LEGAL PIPELINE")
    print("*" * 75)

    overall_start = time.perf_counter()

    results = []
    results.append(run_test_tier("Tier 1: Unit Testing", "tests.test_unit"))
    results.append(run_test_tier("Tier 2: Integration Testing", "tests.test_integration"))
    results.append(run_test_tier("Tier 3: System Testing (Batch Stress)", "tests.test_system_stress"))
    results.append(run_test_tier("Tier 4: Acceptance & Adversarial", "tests.test_acceptance_and_adversarial"))
    results.append(run_test_tier("Tier 5: Headless GUI & UX Testing", "tests.test_gui_automated"))
    results.append(run_test_tier("Tier 6: Security & Zero-Cloud Invariant", "tests.test_security_privacy"))
    results.append(run_test_tier("Tier 7: Legal Compliance & Liability Shield", "tests.test_legal_compliance"))
    results.append(run_test_tier("Tier 8: Packaging & Artifacts Verification", "tests.test_packaging_and_distribution"))
    results.append(run_test_tier("Tier 9-11: Advanced Runtime, Forensics & Production", "tests.test_tier9_to_tier11_advanced"))

    total_time = time.perf_counter() - overall_start

    print("\n" + "#" * 75)
    print("             FINAL EXECUTIVE QA & LEGAL COMPLIANCE REPORT SUMMARY")
    print("#" * 75)
    print(f"{'Testing Tier':<48} | {'Tests':<6} | {'Failures':<8} | {'Time (s)':<8} | {'Status'}")
    print("-" * 75)

    all_passed = True
    total_tests = 0

    for r in results:
        total_tests += r["total"]
        status_str = "✓ PASSED" if r["success"] else "✗ FAILED"
        if not r["success"]:
            all_passed = False
        print(f"{r['tier']:<48} | {r['total']:<6} | {r['failures'] + r['errors']:<8} | {r['time']:<8.2f} | {status_str}")

    print("-" * 75)
    print(f"Total Test Cases: {total_tests} across all 11 Tiers")
    print(f"Total Execution Time: {total_time:.2f} seconds")
    print(f"Overall Pipeline Result: {'🟢 100% PASSED (PRODUCTION & LEGAL COMPLIANCE CERTIFIED)' if all_passed else '🔴 FAILURES DETECTED'}")
    print("#" * 75 + "\n")

    if not all_passed:
        sys.exit(1)


if __name__ == "__main__":
    main()
