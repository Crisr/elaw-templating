import copy
import threading

import emplawra_docx_engine
import db


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
            job = db.get_job(job_id)
            if job is None:
                return

            source_path = job["source_file"]
            lang_code = job["language"]
            mode = job.get("mode", "inline")
            provider_name = job.get("provider")
            model_override = job.get("model")

            config = db.get_config()

            db.update_job(job_id, status="running", progress=0)

            if mode == "transform2cell":
                p = emplawra_docx_engine.get_provider(config, provider_name) if provider_name else None
                if model_override and p:
                    p["model"] = model_override
                result_path = db.UPLOAD_DIR / f"{job_id}_result.docx"
                emplawra_docx_engine.transform2cell(source_path, str(result_path), p)
                db.update_job(job_id, status="done", progress=100, result_file=str(result_path))
                db.enforce_file_limit()
                return

            def progress_callback(done, total):
                db.update_job(job_id, progress=done, total=total)

            paragraphs = emplawra_docx_engine.extract_paragraphs(source_path)
            originals = copy.deepcopy(paragraphs) if mode == "side-by-side" else None

            if lang_code == "none":
                translated = copy.deepcopy(paragraphs)
            else:
                lang = "Romanian" if lang_code == "ro" else "English"
                provider = emplawra_docx_engine.get_provider(config, provider_name)
                if model_override:
                    provider["model"] = model_override
                translated = emplawra_docx_engine.translate_all(
                    paragraphs, lang, provider, progress_callback=progress_callback
                )

            result_path = db.UPLOAD_DIR / f"{job_id}_result.docx"
            if mode == "side-by-side":
                emplawra_docx_engine.write_side_by_side(source_path, originals, translated, str(result_path))
            else:
                emplawra_docx_engine.write_inline(source_path, translated, str(result_path))

            db.update_job(job_id, status="done", progress=100, result_file=str(result_path))
            db.enforce_file_limit()

        except Exception as e:
            db.update_job(job_id, status="failed", error=str(e))


def test_worker_enqueue_dequeue():
    job_id = db.create_job("ro", "inline", None, None)
    from worker import TranslationWorker
    w = TranslationWorker()
    done = threading.Event()
    w.start()

    original_process = w._process
    def tracked_process(jid):
        original_process(jid)
        done.set()
    w._process = tracked_process

    w.enqueue(job_id)
    assert done.wait(timeout=5), "Worker did not process job within 5s"
    job = db.get_job(job_id)
    assert job["status"] == "failed"
    w.stop()
