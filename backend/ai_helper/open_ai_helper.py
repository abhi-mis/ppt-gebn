from openai import OpenAI
from dotenv import load_dotenv
import os
import re

load_dotenv()

class OpenAIHelper:
    def __init__(self):
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("API key not found in .env file")
        self.client = OpenAI(api_key=api_key)

        # Load refine comments prompt
        prompt_path = os.path.join(os.path.dirname(__file__), "prompts", "refine_comments.txt")
        if os.path.exists(prompt_path):
            with open(prompt_path, "r", encoding="utf-8") as f:
                self.refine_prompt_template = f.read().strip()
        else:
            self.refine_prompt_template = (
                "You are a professional project report editor.\n"
                "Refine the following comments for grammar, clarity, and conciseness.\n"
                "Keep the original meaning intact. Preserve date prefixes and names.\n"
                "Return ONLY a numbered list matching input numbering.\n\n"
                "Input:\n{{comments}}"
            )

    def generate_text(self, prompt):
        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=300
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            return f"Error: {str(e)}"

    def summarize_text(self, text):
        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "Summarize the following text concisely."},
                    {"role": "user", "content": text}
                ],
                max_tokens=150
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            return f"Error: {str(e)}"

    def extract_keywords(self, text):
        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "Extract important keywords from the text. Return them as a comma-separated list."},
                    {"role": "user", "content": text}
                ],
                max_tokens=100
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            return f"Error: {str(e)}"
        
    def ask(self, prompt, system_prompt="You are a helpful assistant.", max_tokens=200, temperature=0.7):
        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=max_tokens,
                temperature=temperature
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            return f"Error: {str(e)}"

    # ──────────────────────────────────────────────
    # BULK COMMENT REFINEMENT
    # ──────────────────────────────────────────────

    def refine_comments_bulk(self, comments_list, batch_size=10):
        """
        Refines a list of comments in batches using GPT-4o-mini.

        Args:
            comments_list: List of comment strings to refine.
            batch_size: Number of comments per API call (default 10).

        Returns:
            List of refined comment strings, same length and order as input.
        """
        if not comments_list:
            return []

        all_refined = [""] * len(comments_list)

        # Process in batches
        for batch_start in range(0, len(comments_list), batch_size):
            batch_end = min(batch_start + batch_size, len(comments_list))
            batch = comments_list[batch_start:batch_end]

            # Build numbered list for this batch
            numbered_comments = "\n".join(
                f"{i + 1}. {comment}" for i, comment in enumerate(batch)
            )

            prompt = self.refine_prompt_template.replace("{{comments}}", numbered_comments)

            try:
                response = self.client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {
                            "role": "system",
                            "content": (
                                "You are a professional project report editor. "
                                "Return ONLY the numbered list of refined comments. "
                                "No preamble, no markdown, no explanation."
                            )
                        },
                        {"role": "user", "content": prompt}
                    ],
                    max_tokens=2000,
                    temperature=0.3
                )

                raw_response = response.choices[0].message.content.strip()
                parsed = self._parse_numbered_response(raw_response, len(batch))

                # Place parsed results into the correct positions
                for i, refined_text in enumerate(parsed):
                    all_refined[batch_start + i] = refined_text

                print(f"  ✅ Batch {batch_start + 1}-{batch_end}: refined {len(parsed)} comments")

            except Exception as e:
                print(f"  ⚠️ Batch {batch_start + 1}-{batch_end} failed: {e}")
                # On failure, keep originals for this batch
                for i, original in enumerate(batch):
                    all_refined[batch_start + i] = original

        return all_refined

    def _parse_numbered_response(self, response_text, expected_count):
        """
        Parses a numbered list response from the AI.
        Returns a list of strings, padded/trimmed to expected_count.
        """
        result = [""] * expected_count

        for line in response_text.split("\n"):
            line = line.strip()
            if not line:
                continue

            # Match patterns like "1. text", "1) text", "1: text"
            match = re.match(r"^(\d+)\s*[.):\-]\s*(.+)$", line)
            if match:
                idx = int(match.group(1)) - 1  # Convert to 0-based
                if 0 <= idx < expected_count:
                    result[idx] = match.group(2).strip()

        return result


# Run directly for testing
if __name__ == "__main__":
    helper = OpenAIHelper()

    print("\n--- Generate Text ---")
    print(helper.generate_text("Write a short motivational quote on learning java..."))

    print("\n--- Bulk Refine Test ---")
    test_comments = [
        "Tommy (SSP PM) to expedite internally with the SSP Project team / Stakeholders and have these dependencies to be sorted",
        "We need to check with the client team about the timeline for delivery of the API specifications document which is pending from their side",
        "Short comment ok",
    ]
    refined = helper.refine_comments_bulk(test_comments)
    for orig, ref in zip(test_comments, refined):
        print(f"  Original: {orig}")
        print(f"  Refined:  {ref}\n")
