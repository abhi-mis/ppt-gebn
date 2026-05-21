import os
import json
import datetime
from google import genai

class LLMHelper:
    def __init__(self, api_key: str = None, log_file: str = "gemini_responses.json"):
        """
        Gemini AI helper that refines and grammar-corrects all 'Comments' fields in Excel-like JSON data.
        """
        # self.api_key = api_key or os.getenv("GEMINI_KEY") or self._load_key_from_env_file()
        self.api_key = (
            api_key or
            os.getenv("GEMINI_KEY") or
            self._load_key_from_env_file()
        )

        if not self.api_key:
            raise ValueError(
                "[ERROR] Missing GEMINI_KEY. Please set it either:\n"
                "- in .env as GEMINI_KEY=...\n"
                "- or export it as environment variable\n"
                "- or pass it directly to LLMHelper(api_key='...')"
            )

        if not self.api_key:
            raise ValueError("GEMINI_KEY not found. Please set it in environment or .env file.")

        self.client = genai.Client(api_key=self.api_key)
        self.model = "gemini-2.5-flash"
        self.log_file = log_file

        # Load the LLM prompt template
        prompt_path = os.path.join(os.path.dirname(__file__), "llm_prompt.txt")
        with open(prompt_path, "r", encoding="utf-8") as f:
            self.prompt_template = f.read().strip()

    # --------------------------
    # Internal Utility Functions
    # --------------------------
    def _load_key_from_env_file(self):
        """Reads GEMINI_KEY from a .env file or returns None."""
        env_path = os.path.join(os.getcwd(), ".env")

        if os.path.exists(env_path):
            try:
                with open(env_path, "r", encoding="utf-8") as f:
                    for line in f:
                        if line.strip().startswith("GEMINI_KEY="):
                            return line.strip().split("=", 1)[1].replace('"', '').replace("'", "")
            except:
                pass

        return None


    def _save_response(self, entries):
        """Logs batch responses to a JSON file for review."""
        try:
            data = []
            if os.path.exists(self.log_file):
                with open(self.log_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            data = []

        data.extend(entries)
        with open(self.log_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    # --------------------------
    # Core Gemini Functions
    # --------------------------
    def generate_bulk_responses(self, comments_list):
        """
        Sends multiple comments to Gemini in one go for refinement and grammar correction.
        Uses llm_prompt.txt as the prompt base.
        Ensures 1:1 mapping (no row mismatch).
        """
        if not comments_list:
            return []

        # --- Load system prompt from file ---
        prompt_file_path = os.path.join(os.path.dirname(__file__), "llm_prompt.txt")
        try:
            with open(prompt_file_path, "r", encoding="utf-8") as f:
                base_prompt = f.read().strip()
                print("LLM prompt loaded from llm_prompt.txt.")
        except Exception as e:
            print(f"[WARNING] Could not read llm_prompt.txt: {e}")
            base_prompt = (
                "You are a professional text refiner. "
                "Polish each comment for clarity and grammar while keeping the same meaning. "
                "Do not change or remove date prefixes like '23/02:' or '01/03:'. "
                "If a comment is None or blank, return an empty line."
            )

        # --- Clean and preserve comment list ---
        cleaned_comments = []
        for c in comments_list:
            if not c or str(c).strip().lower() in ["none", "nan", "null", ""]:
                cleaned_comments.append("")  # keep placeholder for blank
            else:
                cleaned_comments.append(str(c).strip())

        # --- Build full numbered prompt (keep all lines, even empty) ---
        joined_comments = "\n".join([
            f"{i+1}. {c if c else '[BLANK]'}"
            for i, c in enumerate(cleaned_comments)
        ])

        bulk_prompt = (
            f"{base_prompt}\n\n"
            "Refine and grammar-correct each numbered comment below.\n"
            "- Keep date prefixes intact.\n"
            "- If input is [BLANK] or None, return an empty line.\n"
            "- Return numbered list (1., 2., etc.) exactly as input.\n\n"
            f"{joined_comments}"
        )

        print(f"[LLM] Sending {len(cleaned_comments)} comments to Gemini (bulk mode)...")

        try:
            response = self.client.models.generate_content(
                model=self.model,
                contents=[bulk_prompt]
            )
            print("[SUCCESS] Gemini bulk response received.")
        except Exception as e:
            print(f"[ERROR] Error during Gemini request: {e}")
            return ["" for _ in cleaned_comments]

        # --- Parse output safely ---
        refined_text = response.text.strip() if response and hasattr(response, "text") else ""
        refined_lines = [""] * len(cleaned_comments)

        for line in refined_text.split("\n"):
            line = line.strip()
            if not line:
                continue
            if line[0].isdigit():
                parts = line.split(".", 1)
                if len(parts) == 2 and parts[0].isdigit():
                    idx = int(parts[0]) - 1
                    if 0 <= idx < len(refined_lines):
                        refined_lines[idx] = parts[1].strip()

        # --- Fallback: ensure same length ---
        refined_lines = (refined_lines + [""] * len(cleaned_comments))[:len(cleaned_comments)]

        # --- Save logs ---
        entries = [
            {
                "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "original": comments_list[i],
                "refined": refined_lines[i]
            }
            for i in range(len(cleaned_comments))
        ]
        self._save_response(entries)

        return refined_lines


    
    def generate_ai_response(self, comment: str) -> str:
        """
        Handles a single comment string and returns the AI-refined version.
        Falls back to blank if the response fails.
        """
        if not comment or not comment.strip():
            return ""

        try:
            response = self.client.models.generate_content(
                model=self.model,
                contents=[self.prompt_template + "\nComment to refine:\n" + comment]
            )
            refined_text = response.text.strip()
            if not refined_text:
                return ""
            
            # Save response to log file
            entry = {
                "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "original": comment,
                "refined": refined_text
            }
            self._save_response([entry])
            
            return refined_text

        except Exception as e:
            print(f"Error during single AI response: {e}")
            return ""


    def process_comments_from_json(self, excel_data):
        """
        Finds all 'Comments' fields in Excel-style JSON data,
        refines them using Gemini in bulk, and updates the data.
        """
        all_comments, comment_positions = [], []

        # Extract all comments
        for t_idx, table in enumerate(excel_data):
            for r_idx, row in enumerate(table.get("rows", [])):
                comment_val = row.get("Comments", {}).get("value")
                if comment_val and isinstance(comment_val, str) and comment_val.strip():
                    all_comments.append(comment_val.strip())
                    comment_positions.append((t_idx, r_idx))

        if not all_comments:
            print("No comments found to process.")
            return excel_data

        # Get AI-refined comments
        refined_comments = self.generate_bulk_responses(all_comments)

        # Update JSON data
        for (t_idx, r_idx), refined_text in zip(comment_positions, refined_comments):
            excel_data[t_idx]["rows"][r_idx]["Comments"]["value"] = refined_text or ""

        print(f"Updated {len(refined_comments)} comments in JSON data.")
        return excel_data
