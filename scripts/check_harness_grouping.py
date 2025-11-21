from pathlib import Path
import importlib.util


def _load_core():
    root = Path(__file__).resolve().parents[1]
    wrapper_path = root / "Application" / "eidp_term_scanner.py"
    core_path = root / "Application" / "eidp_term_scanner.core.py"

    spec = importlib.util.spec_from_file_location("eidp_wrapper", wrapper_path)
    if spec is None or spec.loader is None:
        raise SystemExit("Unable to load wrapper module")
    wrapper = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(wrapper)  # type: ignore[attr-defined]

    core = wrapper._load_core_from(core_path)  # type: ignore[attr-defined]
    wrapper._apply_patches(core)  # type: ignore[attr-defined]
    return core


def main() -> None:
    core = _load_core()

    terms_path = Path("user_inputs/terms.schema.smartsnap.xlsx")
    pdf_sn0000 = Path("Data Packages") / "FakeProgram_SV1_SN0000.pdf"
    pdf_sn1111 = Path("Data Packages") / "FakeProgram_SV1_SN1111.pdf"

    terms = core.load_terms(terms_path)
    harness_terms = [
        t
        for t in terms
        if "harness" in (getattr(t, "term", "") or "").lower()
        or "harness" in (getattr(t, "term_label", "") or "").lower()
    ]
    if not harness_terms:
        print("Harness Resistance term not found in schema")
        return
    spec = harness_terms[0]

    def run_one(pdf: Path, serial: str) -> None:
        print(f"\n=== {pdf.name} / {serial} ===")
        res = core.scan_pdf_for_term_smart(pdf, serial, spec, window_chars=500, case_sensitive=False)
        print(
            f"found={res.found} page={res.page} value={res.number!r} "
            f"units={res.units!r} text_source={res.text_source!r}"
        )

    run_one(pdf_sn0000, "SN0000")
    run_one(pdf_sn1111, "SN1111")


if __name__ == "__main__":
    main()
