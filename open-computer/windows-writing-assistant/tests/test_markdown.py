"""Markdown -> word-processor formatting: parse_markdown / flatten_payload.

Locks the exact field-reported cases: bold labels in numbered lists and
checkbox summary lists arriving as raw '**...**' / '*   [ ]' text.
"""

from docd.render import flatten_payload, parse_markdown
from docd.drivers.fake import FakeDriver


def one(text):
    paras = parse_markdown(text)
    assert len(paras) == 1
    return paras[0]


class TestInline:
    def test_bold(self):
        p = one("**Summary Checklist:** please review")
        assert p["text"] == "Summary Checklist: please review"
        assert p["spans"] == [(0, 18, {"bold": True})]

    def test_italic_star_and_underscore(self):
        p = one("*emphasis* and _more_")
        assert p["text"] == "emphasis and more"
        assert p["spans"] == [(0, 8, {"italic": True}), (13, 17, {"italic": True})]

    def test_bold_italic_and_code(self):
        p = one("***key*** uses `re.sub`")
        assert p["text"] == "key uses re.sub"
        assert p["spans"] == [
            (0, 3, {"bold": True, "italic": True}),
            (9, 15, {"code": True}),
        ]

    def test_snake_case_not_italicized(self):
        p = one("call parse_markdown and flatten_payload today")
        assert p["spans"] == []
        assert "parse_markdown" in p["text"]

    def test_unmatched_markers_left_alone(self):
        p = one("2 * 3 = 6 and a ** b")
        assert p["text"] == "2 * 3 = 6 and a ** b"
        assert p["spans"] == []


class TestBlocks:
    def test_numbered_list_with_bold_label(self):
        # Exact shape from the field report.
        p = one("1.  **Maintain High Performance in Core Areas:** Focus intensely on algorithms.")
        assert p["style"] == "List Number"
        assert p["text"].startswith("Maintain High Performance")
        assert p["spans"][0][2] == {"bold": True}
        assert p["text"][p["spans"][0][0] : p["spans"][0][1]] == (
            "Maintain High Performance in Core Areas:"
        )

    def test_bullet_and_checkboxes(self):
        paras = parse_markdown(
            "**Summary Checklist:**\n"
            "*   [ ] Have I clearly articulated the problem?\n"
            "*   [x] Have I linked my past work?\n"
            "- plain bullet"
        )
        assert paras[0]["spans"] == [(0, 18, {"bold": True})]
        assert paras[1]["style"] == "List Bullet"
        assert paras[1]["text"] == "☐ Have I clearly articulated the problem?"
        assert paras[2]["text"] == "☑ Have I linked my past work?"
        assert paras[3]["style"] == "List Bullet"
        assert paras[3]["text"] == "plain bullet"

    def test_quote_and_heading(self):
        paras = parse_markdown("## Plan\n> stay focused")
        assert paras[0]["style"] == "Heading 2"
        assert paras[1] == {"text": "stay focused", "style": "Quote", "spans": []}

    def test_style_map_off_is_verbatim(self):
        p = parse_markdown("**raw** stays", style_map=False)[0]
        assert p["text"] == "**raw** stays"
        assert p["spans"] == []


class TestFlattenPayload:
    def test_offsets_absolute_across_paragraphs(self):
        paras = parse_markdown("plain first\n**bold** second")
        payload, spans = flatten_payload(paras)
        assert payload == "plain first\rbold second"
        assert spans == [(12, 16, {"bold": True})]
        # Offsets must slice the payload to the formatted text.
        s, e, _ = spans[0]
        assert payload[s:e] == "bold"


class TestFakeDriverIntegration:
    def test_no_raw_markers_reach_the_document(self, tmp_path):
        driver = FakeDriver()
        doc = driver.new_doc()["doc"]
        driver.insert(
            doc,
            "**Summary Checklist:**\n*   [ ] item one\n1. **Bold label:** rest",
            "end",
        )
        text = driver.read(doc)["text"]
        assert "**" not in text
        assert "[ ]" not in text
        assert "☐ item one" in text
        assert "Bold label: rest" in text
