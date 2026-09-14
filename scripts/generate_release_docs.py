#!/usr/bin/env python3
"""
generate_release_docs.py

Automates release note generation using a company .docx template:
1. Fetches the diff (commits + file changes) between two git tags from GitHub.
2. Asks Google Gemini to turn that diff into STRUCTURED JSON (not freeform text).
3. Fills a copy of your .docx template with that data — preserving the template's
   fonts, table borders, and styles exactly. Also writes a plain .md version for
   quick reading / PR diffing.

Setup:
    pip install requests python-docx

    Environment variables (or pass via CLI flags):
        GITHUB_TOKEN   - GitHub personal access token (needed for private repos / higher rate limits)
        GEMINI_API_KEY - Free API key from https://aistudio.google.com/apikey

Usage:
    python generate_release_docs.py \
        --owner myorg --repo myrepo \
        --tag1 v2.0.0 --tag2 v2.0.1 \
        --template release_template.docx \
        --output release_v2.0.1

    Produces release_v2.0.1.docx and release_v2.0.1.md

Template requirements (see release_template.docx for a working example):
    Scalar placeholders (anywhere in body text): {{VERSION}} {{DATE}} {{OVERVIEW}}
        {{CONTRIBUTORS}} {{DIFF_LINK}}
    One table row with cells: {{TYPE}} {{DESCRIPTION}} {{MODULE}} {{OWNER}}
        (this row gets duplicated once per change, then removed)
    One bullet paragraph containing {{BREAKING_CHANGE_ITEM}}
    One bullet paragraph containing {{KNOWN_ISSUE_ITEM}}
        (each duplicated once per list item, then removed)
"""

import argparse
import json
import os
import re
import sys
import textwrap
import requests

from docx_filler import fill_release_template

GITHUB_API = "https://api.github.com"
GEMINI_API = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"

MAX_PATCH_CHARS_PER_FILE = 4000
MAX_TOTAL_DIFF_CHARS = 120_000


def fetch_compare(owner, repo, tag1, tag2, token=None):
    url = f"{GITHUB_API}/repos/{owner}/{repo}/compare/{tag1}...{tag2}"
    headers = {"Accept": "application/vnd.github+json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    resp = requests.get(url, headers=headers, timeout=30)
    if resp.status_code != 200:
        sys.exit(f"GitHub API error ({resp.status_code}): {resp.text[:500]}")
    return resp.json()


def build_diff_summary(compare_data):
    lines = [f"Total commits: {compare_data.get('total_commits', 'unknown')}", "", "## Commits"]
    for c in compare_data.get("commits", []):
        msg = c["commit"]["message"].split("\n")[0]
        author = c["commit"]["author"]["name"]
        sha = c["sha"][:7]
        lines.append(f"- [{sha}] {msg} (by {author})")

    lines.append("")
    lines.append("## Changed files")
    total_chars = sum(len(l) for l in lines)
    for f in compare_data.get("files", []):
        header = f"\n### {f['filename']} ({f['status']}, +{f.get('additions', 0)}/-{f.get('deletions', 0)})\n"
        patch = f.get("patch", "")[:MAX_PATCH_CHARS_PER_FILE]
        block = header + "```diff\n" + patch + "\n```\n"
        if total_chars + len(block) > MAX_TOTAL_DIFF_CHARS:
            lines.append("\n...[remaining file diffs truncated for length]...")
            break
        lines.append(block)
        total_chars += len(block)

    return "\n".join(lines)


def _extract_json(text):
    """Gemini sometimes wraps JSON in ```json fences despite instructions; strip them."""
    text = text.strip()
    fence_match = re.match(r"^```(?:json)?\s*(.*?)\s*```$", text, re.DOTALL)
    if fence_match:
        text = fence_match.group(1)
    return json.loads(text)


def generate_structured_release_data(diff_summary, tag1, tag2, api_key, model="gemini3.6-flash"):
    prompt = textwrap.dedent(f"""
        You are a release notes writer. Based on the commit messages and file diffs below
        (comparing git tag {tag1} to {tag2}), produce a JSON object with EXACTLY this shape
        and nothing else — no markdown fences, no commentary, just the raw JSON:

        {{
          "version": "{tag2}",
          "date": "<today's date, human readable, e.g. 13 September 2026>",
          "overview": "<2-3 sentence plain-English summary of this release>",
          "changes": [
            {{"type": "Feature|Fix|Improvement|Chore", "description": "<what changed, one sentence>", "module": "<affected area, inferred from file paths>", "owner": "<commit author name>"}}
          ],
          "breaking_changes": ["<only if genuinely breaking, else empty list>"],
          "known_issues": ["<only if evident from the diff/commits, else empty list>"],
          "contributors": ["<unique author names from the commits>"]
        }}

        Rules:
        - Base every entry strictly on the commits/diff provided. Do not invent changes.
        - Merge trivial/related commits into single logical change entries where sensible.
        - Keep "module" short (e.g. "Web App", "API", "CI", "Docs") based on the changed file paths.
        - If nothing qualifies for breaking_changes or known_issues, use an empty list, not a placeholder string.

        --- DIFF SUMMARY ({tag1} -> {tag2}) ---
        {diff_summary}
    """).strip()

    url = GEMINI_API.format(model=model)
    resp = requests.post(
        url,
        params={"key": api_key},
        json={
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"response_mime_type": "application/json"},
        },
        timeout=120,
    )
    if resp.status_code != 200:
        sys.exit(f"Gemini API error ({resp.status_code}): {resp.text[:800]}")

    data = resp.json()
    try:
        raw_text = data["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError):
        sys.exit(f"Unexpected Gemini response shape: {data}")

    try:
        return _extract_json(raw_text)
    except json.JSONDecodeError as e:
        sys.exit(f"Gemini did not return valid JSON: {e}\n\nRaw output:\n{raw_text[:2000]}")


def write_markdown_summary(data, tag1, tag2, output_path):
    lines = [f"# Release Notes — v{data.get('version', tag2)}", "", f"_Released: {data.get('date', '')}_", ""]
    lines.append("## Overview")
    lines.append(data.get("overview", ""))
    lines.append("")
    lines.append("## Summary of Changes")
    lines.append("| Type | Description | Module | Owner |")
    lines.append("|---|---|---|---|")
    for c in data.get("changes", []):
        lines.append(f"| {c.get('type','')} | {c.get('description','')} | {c.get('module','')} | {c.get('owner','')} |")
    lines.append("")
    lines.append("## Breaking Changes")
    bc = data.get("breaking_changes", [])
    lines += [f"- {b}" for b in bc] if bc else ["None in this release."]
    lines.append("")
    lines.append("## Known Issues")
    ki = data.get("known_issues", [])
    lines += [f"- {k}" for k in ki] if ki else ["None reported."]
    lines.append("")
    lines.append("## Contributors")
    lines.append(", ".join(data.get("contributors", [])) or "N/A")
    lines.append("")
    lines.append(f"**Full diff:** `{tag1}...{tag2}`")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def main():
    parser = argparse.ArgumentParser(description="Generate release docs from a GitHub tag diff using Gemini + a docx template.")
    parser.add_argument("--owner", required=True)
    parser.add_argument("--repo", required=True)
    parser.add_argument("--tag1", required=True)
    parser.add_argument("--tag2", required=True)
    parser.add_argument("--template", required=True, help="Path to the .docx template with placeholder tokens")
    parser.add_argument("--output", default="release_notes", help="Output file path WITHOUT extension")
    parser.add_argument("--github-token", default=os.environ.get("GITHUB_TOKEN"))
    parser.add_argument("--gemini-key", default=os.environ.get("GEMINI_API_KEY"))
    parser.add_argument("--model", default="gemini-3.6-flash")
    args = parser.parse_args()

    if not args.gemini_key:
        sys.exit("Missing Gemini API key. Set GEMINI_API_KEY env var or pass --gemini-key.\n"
                  "Get a free key at https://aistudio.google.com/apikey")
    if not os.path.exists(args.template):
        sys.exit(f"Template not found: {args.template}")

    print(f"Fetching diff between {args.tag1} and {args.tag2} ...")
    compare_data = fetch_compare(args.owner, args.repo, args.tag1, args.tag2, args.github_token)

    print("Building diff summary ...")
    diff_summary = build_diff_summary(compare_data)

    print("Asking Gemini for structured release data ...")
    data = generate_structured_release_data(diff_summary, args.tag1, args.tag2, args.gemini_key, args.model)

    docx_path = args.output + ".docx"
    md_path = args.output + ".md"

    print(f"Filling template -> {docx_path}")
    fill_release_template(args.template, docx_path, data)

    print(f"Writing markdown summary -> {md_path}")
    write_markdown_summary(data, args.tag1, args.tag2, md_path)

    print(f"Done.\n  {docx_path}\n  {md_path}")


if __name__ == "__main__":
    main()
