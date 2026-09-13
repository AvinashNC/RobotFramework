#!/usr/bin/env python3
"""
generate_release_docs.py

Automates release note generation:
1. Fetches the diff (commits + file changes) between two git tags from GitHub.
2. Feeds that diff + a previous release doc (as a style reference) to Google Gemini.
3. Writes the generated release doc to a markdown file.

Setup:
    pip install requests

    Environment variables (or pass via CLI flags):
        GITHUB_TOKEN   - GitHub personal access token (needed for private repos / higher rate limits)
        GEMINI_API_KEY - Free API key from https://aistudio.google.com/apikey

Usage:
    python generate_release_docs.py \
        --owner myorg \
        --repo myrepo \
        --tag1 v1.2.0 \
        --tag2 v1.3.0 \
        --prev-doc ./previous_release_notes.md \
        --output ./release_notes_v1.3.0.md
"""

import argparse
import os
import sys
import textwrap
import requests

GITHUB_API = "https://api.github.com"
GEMINI_API = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"

MAX_PATCH_CHARS_PER_FILE = 4000   # truncate huge file diffs
MAX_TOTAL_DIFF_CHARS = 120_000    # keep total prompt reasonable


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
    lines = []
    lines.append(f"Total commits: {compare_data.get('total_commits', 'unknown')}")
    lines.append("")
    lines.append("## Commits")
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


def generate_release_doc(diff_summary, prev_doc_text, tag1, tag2, api_key, model="gemini-2.0-flash"):
    prompt = textwrap.dedent(f"""
        You are a release notes writer. You will be given:
        1. A PREVIOUS RELEASE DOC that shows the exact structure, tone, section headers,
           and formatting style this project uses for release notes.
        2. A DIFF SUMMARY (commit messages + changed files) between two tags: {tag1} -> {tag2}.

        Your task: write a NEW release doc for the {tag1} -> {tag2} change, matching the
        structure/style/tone of the previous doc as closely as possible. Group changes into
        sensible sections (e.g. Features, Fixes, Breaking Changes, Chores) based on the commit
        messages and diffs — infer categories the same way the previous doc appears to.
        Do not invent changes that aren't supported by the commits/diff. If something is unclear,
        summarize conservatively rather than guessing details.

        --- PREVIOUS RELEASE DOC (style reference) ---
        {prev_doc_text}

        --- DIFF SUMMARY ({tag1} -> {tag2}) ---
        {diff_summary}

        Now write the new release doc in markdown, ready to publish.
    """).strip()

    url = GEMINI_API.format(model=model)
    resp = requests.post(
        url,
        params={"key": api_key},
        json={"contents": [{"parts": [{"text": prompt}]}]},
        timeout=120,
    )
    if resp.status_code != 200:
        sys.exit(f"Gemini API error ({resp.status_code}): {resp.text[:800]}")

    data = resp.json()
    try:
        return data["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError):
        sys.exit(f"Unexpected Gemini response shape: {data}")


def main():
    parser = argparse.ArgumentParser(description="Generate release docs from a GitHub tag diff using Gemini.")
    parser.add_argument("--owner", required=True, help="GitHub repo owner/org")
    parser.add_argument("--repo", required=True, help="GitHub repo name")
    parser.add_argument("--tag1", required=True, help="Previous tag")
    parser.add_argument("--tag2", required=True, help="New tag")
    parser.add_argument("--prev-doc", required=True, help="Path to previous release doc (markdown/txt)")
    parser.add_argument("--output", default="release_notes.md", help="Output file path")
    parser.add_argument("--github-token", default=os.environ.get("GITHUB_TOKEN"))
    parser.add_argument("--gemini-key", default=os.environ.get("GEMINI_API_KEY"))
    parser.add_argument("--model", default="gemini-2.0-flash")
    args = parser.parse_args()

    if not args.gemini_key:
        sys.exit("Missing Gemini API key. Set GEMINI_API_KEY env var or pass --gemini-key.\n"
                  "Get a free key at https://aistudio.google.com/apikey")

    if not os.path.exists(args.prev_doc):
        sys.exit(f"Previous release doc not found: {args.prev_doc}")

    print(f"Fetching diff between {args.tag1} and {args.tag2} ...")
    compare_data = fetch_compare(args.owner, args.repo, args.tag1, args.tag2, args.github_token)

    print("Building diff summary ...")
    diff_summary = build_diff_summary(compare_data)

    with open(args.prev_doc, "r", encoding="utf-8") as f:
        prev_doc_text = f.read()

    print("Generating release doc with Gemini ...")
    output_text = generate_release_doc(
        diff_summary, prev_doc_text, args.tag1, args.tag2, args.gemini_key, args.model
    )

    with open(args.output, "w", encoding="utf-8") as f:
        f.write(output_text)

    print(f"Done. Release doc written to {args.output}")


if __name__ == "__main__":
    main()
