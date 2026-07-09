import threading
import time

import translate_docx
from server import update_job, get_job, enforce_file_limit, config, UPLOAD_DIR


class TranslationWorker:
    def __init__(self):
        self._queue = []
        self._lock = threading.Lock()
        self._cv = threading.Condition(self._lock)
        self._running = False
        self._thread = None

    def start(self):
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self):
        self._running = False
        with self._cv:
            self._cv.notify_all()
        if self._thread:
            self._thread.join(timeout=5)

    def enqueue(self, job_id):
        with self._cv:
            self._queue.append(job_id)
            self._cv.notify()

    def _run(self):
        while self._running:
            job_id = None
            with self._cv:
                while self._running and not self._queue:
                    self._cv.wait(timeout=1)
                if self._queue:
                    job_id = self._queue.pop(0)
            if job_id and self._running:
                self._process(job_id)

    def _process(self, job_id):
        try:
            job = get_job(job_id)
            if job is None:
                return

            source_path = job["source_file"]
            lang = "Romanian" if job["language"] == "ro" else "English"
            mode = job.get("mode", "inline")
            provider_name = job.get("provider")
            model_override = job.get("model")

            provider = translate_docx.get_provider(config, provider_name)
            if model_override:
                provider["model"] = model_override

            update_job(job_id, status="running", progress=0)

            def progress_callback(done, total):
                update_job(job_id, progress=done, total=total)

            paragraphs = translate_docx.extract_paragraphs(source_path)
            originals = __import__("copy").deepcopy(paragraphs) if mode == "side-by-side" else None
            translated = translate_docx.translate_all(
                paragraphs, lang, provider, progress_callback=progress_callback
            )

            result_path = UPLOAD_DIR / f"{job_id}_result.docx"
            if mode == "side-by-side":
                translate_docx.write_side_by_side(source_path, originals, translated, str(result_path))
            else:
                translate_docx.write_inline(source_path, translated, str(result_path))

            update_job(job_id, status="done", progress=100, result_file=str(result_path))
            enforce_file_limit()

        except Exception as e:
            update_job(job_id, status="failed", error=str(e))


def test_worker_enqueue_dequeue():
    import server
    job_id = server.create_job("ro", "inline", None, None)
    from worker import TranslationWorker
    w = TranslationWorker()
    w.start()
    w.enqueue(job_id)
    import time
    time.sleep(0.5)
    job = server.get_job(job_id)
    assert job["status"] == "failed"
    w.stop()
