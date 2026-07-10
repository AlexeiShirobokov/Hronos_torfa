"""Qwen project bootstrap + universal RAG prompt injector.

This script is intentionally dependency-free. It reads stable project context,
retrieves relevant fragments from the user's universal Obsidian notes, injects
both into the prompt, and sends the enriched request to a local OpenAI-compatible
Qwen server such as mlx_lm.server.
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


BASE = Path(__file__).resolve().parent
DEFAULT_BASE_URL = "http://127.0.0.1:8081"
DEFAULT_MODEL = "mlx-community/Qwen3.6-35B-A3B-4bit"
DEFAULT_NOTES_DIR = Path(
    "~/Obsidian/Alexeids/Tech/ML-NLP/Agents"
).expanduser()

PROJECT_BOOTSTRAP_FILES = (
    "AGENTS.md",
    "PROJECT_STATE.md",
    "README.md",
)

PROJECT_SEARCH_GLOBS = (
    "docs/**/*.md",
    ".vscode/tasks.json",
    ".vscode/settings.json",
)

EXCLUDE_DIRS = {
    ".git",
    ".venv",
    "venv",
    "__pycache__",
    "_tmp",
    "input",
    "output",
    "logs",
    "state",
    "node_modules",
    ".mypy_cache",
    ".pytest_cache",
}

SENSITIVE_FILE_NAMES = {
    ".env",
    ".env.local",
}

SECRET_NAME_RE = re.compile(
    r"(password|passwd|secret|token|api[_-]?key|YANDEX_APP_PASSWORD|BOT_TOKEN)",
    re.IGNORECASE,
)
TOKEN_RE = re.compile(r"[A-Za-zА-Яа-яЁё0-9_+-]+")


@dataclass(frozen=True)
class Chunk:
    path: Path
    title: str
    text: str
    source_kind: str


@dataclass(frozen=True)
class SearchHit:
    chunk: Chunk
    score: float


def read_text(path: Path, max_chars: int | None = None) -> str:
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        text = path.read_text(encoding="utf-8", errors="replace")
    if max_chars is not None and len(text) > max_chars:
        return text[:max_chars] + "\n\n[... truncated ...]"
    return text


def rel(path: Path, root: Path | None = None) -> str:
    if root:
        try:
            return str(path.resolve().relative_to(root.resolve()))
        except ValueError:
            pass
    return str(path)


def normalize_token(token: str) -> str:
    return token.lower().replace("ё", "е")


def tokenize(text: str) -> list[str]:
    return [normalize_token(t) for t in TOKEN_RE.findall(text)]


def split_markdown(path: Path, text: str, source_kind: str) -> list[Chunk]:
    chunks: list[Chunk] = []
    current_title = path.name
    current: list[str] = []

    def flush() -> None:
        body = "\n".join(current).strip()
        if body:
            chunks.extend(split_long_chunk(path, current_title, body, source_kind))

    for line in text.splitlines():
        if line.startswith("#"):
            flush()
            current = [line]
            current_title = line.lstrip("#").strip() or path.name
        else:
            current.append(line)
    flush()
    return chunks


def split_long_chunk(
    path: Path,
    title: str,
    text: str,
    source_kind: str,
    max_chars: int = 2200,
) -> list[Chunk]:
    if len(text) <= max_chars:
        return [Chunk(path=path, title=title, text=text, source_kind=source_kind)]

    chunks: list[Chunk] = []
    paragraphs = re.split(r"\n\s*\n", text)
    buf: list[str] = []
    size = 0
    part = 1
    for para in paragraphs:
        para_len = len(para) + 2
        if buf and size + para_len > max_chars:
            chunks.append(
                Chunk(path=path, title=f"{title} (part {part})", text="\n\n".join(buf), source_kind=source_kind)
            )
            buf = []
            size = 0
            part += 1
        if para_len > max_chars:
            for i in range(0, len(para), max_chars):
                chunks.append(
                    Chunk(
                        path=path,
                        title=f"{title} (part {part})",
                        text=para[i : i + max_chars],
                        source_kind=source_kind,
                    )
                )
                part += 1
            continue
        buf.append(para)
        size += para_len
    if buf:
        chunks.append(
            Chunk(path=path, title=f"{title} (part {part})", text="\n\n".join(buf), source_kind=source_kind)
        )
    return chunks


def iter_markdown_files(notes_dir: Path) -> Iterable[Path]:
    if not notes_dir.exists():
        return []
    return sorted(p for p in notes_dir.rglob("*.md") if p.is_file())


def load_universal_chunks(notes_dir: Path) -> list[Chunk]:
    chunks: list[Chunk] = []
    for path in iter_markdown_files(notes_dir):
        text = read_text(path, max_chars=120_000)
        chunks.extend(split_markdown(path, text, source_kind="universal"))
    return chunks


def load_project_search_chunks(project_root: Path) -> list[Chunk]:
    chunks: list[Chunk] = []
    for pattern in PROJECT_SEARCH_GLOBS:
        for path in sorted(project_root.glob(pattern)):
            if not path.is_file() or any(part in EXCLUDE_DIRS for part in path.parts):
                continue
            text = read_text(path, max_chars=80_000)
            if path.suffix.lower() == ".md":
                chunks.extend(split_markdown(path, text, source_kind="project"))
            else:
                chunks.extend(split_long_chunk(path, path.name, text, source_kind="project"))
    return chunks


def rank_chunks(query: str, chunks: list[Chunk], top_k: int) -> list[SearchHit]:
    query_tokens = [t for t in tokenize(query) if len(t) > 1]
    if not query_tokens or not chunks:
        return []

    query_set = set(query_tokens)
    doc_tokens: list[list[str]] = [tokenize(c.title + "\n" + c.text) for c in chunks]
    df: dict[str, int] = {}
    for toks in doc_tokens:
        for tok in set(toks):
            if tok in query_set:
                df[tok] = df.get(tok, 0) + 1

    total_docs = len(chunks)
    hits: list[SearchHit] = []
    query_lower = query.lower()
    for chunk, toks in zip(chunks, doc_tokens):
        if not toks:
            continue
        counts: dict[str, int] = {}
        for tok in toks:
            if tok in query_set:
                counts[tok] = counts.get(tok, 0) + 1
        score = 0.0
        for tok, count in counts.items():
            idf = math.log((total_docs + 1) / (df.get(tok, 0) + 1)) + 1.0
            score += (1.0 + math.log(count)) * idf
        title_lower = chunk.title.lower()
        text_lower = chunk.text.lower()
        if query_lower and query_lower in text_lower:
            score += 8.0
        for tok in query_set:
            if tok in title_lower:
                score += 2.0
        if score > 0:
            hits.append(SearchHit(chunk=chunk, score=score))
    hits.sort(key=lambda h: h.score, reverse=True)
    return hits[:top_k]


def project_tree(project_root: Path, max_entries: int = 120) -> str:
    rows: list[str] = []
    for path in sorted(project_root.rglob("*")):
        try:
            relative = path.relative_to(project_root)
        except ValueError:
            continue
        if any(part in EXCLUDE_DIRS for part in relative.parts):
            continue
        if path.is_dir():
            continue
        if path.name in SENSITIVE_FILE_NAMES or SECRET_NAME_RE.search(path.name):
            continue
        rows.append(str(relative))
        if len(rows) >= max_entries:
            rows.append("[... truncated file tree ...]")
            break
    return "\n".join(rows)


def collect_project_context(project_root: Path, max_chars: int) -> str:
    blocks: list[str] = [
        f"Project root: {project_root}",
        "",
        "File tree (filtered):",
        project_tree(project_root),
    ]

    remaining = max_chars - sum(len(b) for b in blocks)
    for name in PROJECT_BOOTSTRAP_FILES:
        path = project_root / name
        if not path.exists() or not path.is_file():
            continue
        per_file_limit = max(1500, min(remaining, 9000))
        if per_file_limit <= 0:
            break
        text = read_text(path, max_chars=per_file_limit)
        blocks.append(f"\n--- {name} ---\n{text}")
        remaining = max_chars - sum(len(b) for b in blocks)
    text = "\n".join(blocks)
    if len(text) > max_chars:
        return text[:max_chars] + "\n\n[... project context truncated ...]"
    return text


def format_hits(hits: list[SearchHit], root: Path | None, max_chars: int) -> str:
    if not hits:
        return "No matching fragments found."
    blocks: list[str] = []
    used = 0
    for i, hit in enumerate(hits, start=1):
        chunk = hit.chunk
        source = rel(chunk.path, root)
        header = f"[{i}] {chunk.source_kind} score={hit.score:.2f} source={source} section={chunk.title}"
        body = chunk.text.strip()
        block = f"{header}\n{body}\n"
        if used + len(block) > max_chars:
            left = max_chars - used
            if left > 500:
                blocks.append(block[:left] + "\n[... rag context truncated ...]\n")
            break
        blocks.append(block)
        used += len(block)
    return "\n".join(blocks)


def build_messages(
    task: str,
    project_root: Path,
    notes_dir: Path,
    top_k: int,
    max_project_chars: int,
    max_rag_chars: int,
    no_rag: bool = False,
) -> tuple[list[dict[str, str]], list[SearchHit], list[SearchHit]]:
    project_context = collect_project_context(project_root, max_project_chars)

    project_hits = rank_chunks(task, load_project_search_chunks(project_root), top_k=3)
    universal_hits: list[SearchHit] = []
    if not no_rag:
        universal_hits = rank_chunks(task, load_universal_chunks(notes_dir), top_k=top_k)

    project_rag = format_hits(project_hits, project_root, max_chars=max_rag_chars // 3)
    universal_rag = format_hits(universal_hits, notes_dir, max_chars=max_rag_chars)

    system = (
        "Ты Qwen, локальный инженерный помощник. Отвечай по-русски.\n"
        "Контекст ниже автоматически собран обвязкой project bootstrap + universal RAG.\n"
        "Считай предоставленные правила и состояние проекта более надежными, чем догадки.\n"
        "Не выводи секреты, пароли, токены и содержимое .env. Если данных не хватает, "
        "скажи, что именно нужно проверить.\n"
        "Если задача требует действий вне текущего проекта, явно отметь это и запроси разрешение."
    )
    user = (
        "<project_bootstrap>\n"
        f"{project_context}\n"
        "</project_bootstrap>\n\n"
        "<project_retrieval>\n"
        f"{project_rag}\n"
        "</project_retrieval>\n\n"
        "<universal_rag>\n"
        f"{universal_rag}\n"
        "</universal_rag>\n\n"
        "<user_task>\n"
        f"{task}\n"
        "</user_task>"
    )
    return [{"role": "system", "content": system}, {"role": "user", "content": user}], project_hits, universal_hits


def call_qwen(
    messages: list[dict[str, str]],
    base_url: str,
    model: str,
    temperature: float,
    timeout: int,
    max_tokens: int,
) -> str:
    url = base_url.rstrip("/") + "/v1/chat/completions"
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            result = json.loads(resp.read().decode("utf-8"))
    except urllib.error.URLError as exc:
        raise RuntimeError(
            f"Qwen server is unavailable at {base_url}. Start it with: qwen-mlx start"
        ) from exc
    except json.JSONDecodeError as exc:
        raise RuntimeError("Qwen returned a non-JSON response") from exc

    try:
        return result["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError(f"Unexpected Qwen response shape: {result!r}") from exc


def positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be positive")
    return parsed


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Inject project context + universal RAG into a local Qwen prompt."
    )
    parser.add_argument("task", nargs="*", help="User request for Qwen")
    parser.add_argument("--project", default=os.getcwd(), help="Project root to bootstrap")
    parser.add_argument("--notes", default=str(DEFAULT_NOTES_DIR), help="Universal Obsidian notes dir")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL, help="Qwen OpenAI-compatible base URL")
    parser.add_argument("--model", default=DEFAULT_MODEL, help="Qwen model id")
    parser.add_argument("--top-k", type=positive_int, default=6, help="Universal RAG fragments")
    parser.add_argument("--max-project-chars", type=positive_int, default=24_000)
    parser.add_argument("--max-rag-chars", type=positive_int, default=18_000)
    parser.add_argument("--max-tokens", type=positive_int, default=2048)
    parser.add_argument("--temperature", type=float, default=0.2)
    parser.add_argument("--timeout", type=positive_int, default=300)
    parser.add_argument("--dry-run", action="store_true", help="Print injected messages, do not call Qwen")
    parser.add_argument("--sources", action="store_true", help="Print retrieval sources to stderr")
    parser.add_argument("--no-rag", action="store_true", help="Disable universal RAG")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    task = " ".join(args.task).strip()
    if not task:
        task = sys.stdin.read().strip()
    if not task:
        print("No task provided. Pass a prompt or pipe text to stdin.", file=sys.stderr)
        return 2

    project_root = Path(args.project).expanduser().resolve()
    notes_dir = Path(args.notes).expanduser().resolve()
    if not project_root.exists() or not project_root.is_dir():
        print(f"Project root not found: {project_root}", file=sys.stderr)
        return 2

    messages, project_hits, universal_hits = build_messages(
        task=task,
        project_root=project_root,
        notes_dir=notes_dir,
        top_k=args.top_k,
        max_project_chars=args.max_project_chars,
        max_rag_chars=args.max_rag_chars,
        no_rag=args.no_rag,
    )

    if args.sources:
        print("Project retrieval sources:", file=sys.stderr)
        for hit in project_hits:
            print(f"- {hit.score:.2f} {rel(hit.chunk.path, project_root)} :: {hit.chunk.title}", file=sys.stderr)
        print("Universal retrieval sources:", file=sys.stderr)
        for hit in universal_hits:
            print(f"- {hit.score:.2f} {rel(hit.chunk.path, notes_dir)} :: {hit.chunk.title}", file=sys.stderr)

    if args.dry_run:
        print(json.dumps({"messages": messages}, ensure_ascii=False, indent=2))
        return 0

    try:
        answer = call_qwen(
            messages=messages,
            base_url=args.base_url,
            model=args.model,
            temperature=args.temperature,
            timeout=args.timeout,
            max_tokens=args.max_tokens,
        )
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(answer)
    return 0


if __name__ == "__main__":
    sys.exit(main())
