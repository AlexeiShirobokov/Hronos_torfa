import tempfile
import unittest
from pathlib import Path

import qwen_bootstrap as qb


class TestQwenBootstrap(unittest.TestCase):
    def test_project_context_includes_state_but_not_env(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "AGENTS.md").write_text("## Rules\nОтвечай по-русски", encoding="utf-8")
            (root / "PROJECT_STATE.md").write_text("# State\nPipeline is ready", encoding="utf-8")
            (root / ".env").write_text("BOT_TOKEN=secret", encoding="utf-8")

            text = qb.collect_project_context(root, max_chars=10_000)

        self.assertIn("AGENTS.md", text)
        self.assertIn("Pipeline is ready", text)
        self.assertNotIn("BOT_TOKEN", text)
        self.assertNotIn(".env", text)

    def test_rank_chunks_returns_relevant_universal_note(self):
        chunks = [
            qb.Chunk(
                path=Path("rules.md"),
                title="DWG workflow",
                text="Для DWG сначала спроси разрешение на внешний путь и используй C#.",
                source_kind="universal",
            ),
            qb.Chunk(
                path=Path("mail.md"),
                title="Email",
                text="SMTP рассылка отчетов.",
                source_kind="universal",
            ),
        ]

        hits = qb.rank_chunks("как работать с DWG файлом", chunks, top_k=1)

        self.assertEqual(hits[0].chunk.title, "DWG workflow")

    def test_build_messages_injects_project_and_universal_context(self):
        with tempfile.TemporaryDirectory() as project_td, tempfile.TemporaryDirectory() as notes_td:
            project = Path(project_td)
            notes = Path(notes_td)
            (project / "README.md").write_text("# Demo\nProject bootstrap info", encoding="utf-8")
            (notes / "Правила.md").write_text(
                "# Универсальные правила\nQwen должен отвечать по-русски.",
                encoding="utf-8",
            )

            messages, _, universal_hits = qb.build_messages(
                task="какие правила для Qwen",
                project_root=project,
                notes_dir=notes,
                top_k=3,
                max_project_chars=10_000,
                max_rag_chars=10_000,
            )

        joined = "\n".join(m["content"] for m in messages)
        self.assertIn("<project_bootstrap>", joined)
        self.assertIn("<universal_rag>", joined)
        self.assertIn("Project bootstrap info", joined)
        self.assertTrue(universal_hits)


if __name__ == "__main__":
    unittest.main()
