from utils.ppt_slide_summary import extract_slides_text
from ai_helper.open_ai_helper import OpenAIHelper


def process_slides_and_stream(ppt_path, prompt_template, job_id, socketio):
    """
    Extracts text from each slide, summarizes it using OpenAI,
    and streams per-slide summaries via socket events.

    Socket events emitted:
        - "slide_summary"    → { job_id, slide_number, summary, progress }
        - "summary_complete" → { job_id, status: "done" }
        - "summary_error"    → { job_id, error }
    """
    try:
        ai = OpenAIHelper()
        slides = extract_slides_text(ppt_path)
        total = len(slides)

        for i, slide in enumerate(slides):
            if not slide["text"].strip():
                continue

            prompt = prompt_template.replace("{{text}}", slide["text"])

            summary = ai.ask(
                prompt=prompt,
                system_prompt="You are a professional Project Manager. Summarize concisely.",
                max_tokens=80,
                temperature=0.3
            )

            # Emit per slide
            socketio.emit("slide_summary", {
                "job_id": job_id,
                "slide_number": slide["slide_number"],
                "summary": summary,
                "progress": f"{i+1}/{total}"
            })

        # Done event
        socketio.emit("summary_complete", {
            "job_id": job_id,
            "status": "done"
        })

    except Exception as e:
        socketio.emit("summary_error", {
            "job_id": job_id,
            "error": str(e)
        })
