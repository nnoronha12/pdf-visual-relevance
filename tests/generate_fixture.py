from pathlib import Path

from tests.test_analyzer import make_fixture_pdf


if __name__ == "__main__":
    output = Path(__file__).parent / "generated" / "visual_relevance_fixture.pdf"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(make_fixture_pdf())
    print(output)
