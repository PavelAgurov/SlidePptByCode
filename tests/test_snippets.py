"""Snippet registry integrity."""

from src.snippets import SNIPPETS, SNIPPET_DESCRIPTIONS, SnippetId, get_code_snippet_tools


def test_every_snippet_id_has_body_and_description() -> None:
    for sid in SnippetId:
        assert sid in SNIPPETS
        assert SNIPPETS[sid].strip()
        assert sid in SNIPPET_DESCRIPTIONS
        assert SNIPPET_DESCRIPTIONS[sid].strip()


def test_table_snippet_mentions_add_table() -> None:
    body = SNIPPETS[SnippetId.TABLE_MARKDOWN_GRID]
    assert "add_table" in body
    assert "parse_pipe_markdown_table" in body


def test_strict_tool_schema_includes_table_id() -> None:
    tools = get_code_snippet_tools()
    enum_vals = tools[0]["function"]["parameters"]["properties"]["snippet_id"]["enum"]
    assert "table_markdown_grid" in enum_vals
    assert tools[0]["function"].get("strict") is True


def test_incremental_h2_skeleton_uses_placeholder_types_not_numeric_body() -> None:
    body = SNIPPETS[SnippetId.INCREMENTAL_H2_SKELETON]
    assert "PP_PLACEHOLDER_TYPE" in body
    assert "first_body_placeholder" in body
    assert "pick_title_and_content_layout" in body
    assert "placeholders[1]" not in body


def test_content_slide_snippet_avoids_placeholders_index_one_for_body() -> None:
    body = SNIPPETS[SnippetId.CONTENT_SLIDE]
    assert "PP_PLACEHOLDER_TYPE" in body
    assert "first_body_placeholder" in body
    assert "placeholders[1]" not in body
