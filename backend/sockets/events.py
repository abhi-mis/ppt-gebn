import threading
from flask_socketio import emit
from sockets.generate_ppt_images import generate_ppt_images
from sockets.socket_service import process_event_message
from sockets.refine_comments import refine_comments_in_ppt


def register_socket_events(socketio):

    @socketio.on("connect")
    def handle_connect():
        print("Client connected")
        emit("connected", {"message": "Connected to server"})

    @socketio.on("disconnect")
    def handle_disconnect():
        print("Client disconnected")

    @socketio.on("start_processing")
    def handle_start_processing(data):
        print("Processing started:", data)
        emit("processing_update", {"status": "started"})

    @socketio.on("cancel_processing")
    def handle_cancel(data):
        print("Cancel requested:", data)
        emit("processing_update", {"status": "cancelled"})

    @socketio.on("message")
    def handle_message(data):
        """
        Frontend now sends AI options alongside the message:
        { message, slideNumber, ppt_name, options: { jargons: true, risks_slide: true, ... } }
        """
        print("Message received:", data)
        response = process_event_message(data)
        emit("message_response", {"message": response, "status": "received"})

    @socketio.on("generate_slide_images")
    def handle_generate_slide_images(data):
        print("Generate slide images request received:", data)
        response = generate_ppt_images(data, emit_fn=emit)
        emit("slide_images", {"message": response, "status": "received"})

    # ──────────────────────────────────────────
    # Refine long comments using AI
    # ──────────────────────────────────────────
    @socketio.on("refine_comments")
    def handle_refine_comments(data):
        print(f"\n{'='*50}")
        print(f"  REFINE COMMENTS: {data.get('ppt_name', 'N/A')}")
        print(f"{'='*50}\n")
        emit("refine_status", {"status": "started", "message": "Starting comment refinement..."})

        def _run():
            try:
                refine_comments_in_ppt(data, emit_fn=lambda ev, pl: socketio.emit(ev, pl))
            except Exception as e:
                print(f"Refine error: {e}")
                socketio.emit("refine_error", {"error": str(e), "status": "error"})

        threading.Thread(target=_run, daemon=True).start()

    # ──────────────────────────────────────────
    # Generate slide summaries on demand
    # ──────────────────────────────────────────
    @socketio.on("generate_summaries")
    def handle_generate_summaries(data):
        import os
        from utils.process_slides_and_stream import process_slides_and_stream as _stream

        ppt_name = data.get("ppt_name")
        if not ppt_name:
            emit("summary_error", {"job_id": "", "error": "No ppt_name provided"})
            return

        ppt_path = os.path.join("generated", ppt_name)
        if not os.path.exists(ppt_path):
            emit("summary_error", {"job_id": "", "error": f"File not found: {ppt_name}"})
            return

        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        prompt_file = os.path.join(BASE_DIR, "prompts", "summarization_prompt.txt")
        try:
            with open(prompt_file, "r", encoding="utf-8") as f:
                prompt_template = f.read()
        except Exception as e:
            emit("summary_error", {"job_id": "", "error": f"Prompt file error: {e}"})
            return

        # Use ppt_name as job_id for consistency
        job_id = ppt_name.replace(".pptx", "").replace("-fixed", "")
        print(f"\n{'='*50}")
        print(f"  GENERATE SUMMARIES (on-demand): {ppt_name}")
        print(f"{'='*50}\n")

        def _run():
            try:
                _stream(ppt_path, prompt_template, job_id, socketio)
            except Exception as e:
                print(f"Summary error: {e}")
                socketio.emit("summary_error", {"job_id": job_id, "error": str(e)})

        threading.Thread(target=_run, daemon=True).start()

    # ──────────────────────────────────────────
    # Restyle PPT with user-configured colors/fonts
    # ──────────────────────────────────────────
    @socketio.on("restyle_ppt")
    def handle_restyle_ppt(data):
        """
        Frontend sends:
        {
            ppt_name: "uuid_report.pptx",
            style_config: {
                hdr_bg_color: "14366B",
                hdr_text_color: "FFFFFF",
                hdr_font: "Calibri",
                hdr_size: 11,
                data_font: "Calibri",
                data_size: 11,
                title_color: "14366B",
                title_size: 20,
                ...
            }
        }
        """
        import os
        from pptx import Presentation as Prs
        from ppt_sanitizer import sanitize_presentation

        ppt_name = data.get("ppt_name")
        style_config = data.get("style_config", {})

        print(f"\n{'='*50}")
        print(f"  RESTYLE PPT: {ppt_name}")
        print(f"  Config: {style_config}")
        print(f"{'='*50}\n")

        if not ppt_name:
            emit("restyle_error", {"error": "No ppt_name provided"})
            return

        ppt_path = os.path.join("generated", ppt_name)
        if not os.path.exists(ppt_path):
            emit("restyle_error", {"error": f"File not found: {ppt_name}"})
            return

        emit("restyle_status", {"status": "processing", "message": "Applying style changes..."})

        def _run():
            try:
                prs = Prs(ppt_path)
                sanitize_presentation(prs, style_config)
                prs.save(ppt_path)
                print(f"Restyled and saved: {ppt_path}")
                socketio.emit("restyle_complete", {
                    "status": "done",
                    "ppt_name": ppt_name,
                    "message": "PPT restyled successfully",
                })
            except Exception as e:
                print(f"Restyle error: {e}")
                socketio.emit("restyle_error", {"error": str(e), "status": "error"})

        threading.Thread(target=_run, daemon=True).start()
