import re

from ai_helper.open_ai_helper import OpenAIHelper

helper = OpenAIHelper()

def determine_slide_from_text(user_text):
    """
    Uses the LLM to detect which slide the user is asking about.
    Returns slide_number (int) or None if not specified.
    """

    prompt = f"""
        You are an intelligent assistant helping a user with a presentation.
        The user may ask questions like:
            - "I don't understand slide 2"
            - "Explain the table in slide 3"
            - "What does the first slide say?"
        Extract the slide number they are referring to as an integer.
        If no slide is mentioned, return None.
        Respond with just the number or None.
        User question: "{user_text}"
        Slide number:
    """

    slide_number_text = helper.ask(prompt, max_tokens=10, temperature=0)

    # Remove non-digit characters and search for integer
    match = re.search(r"\d+", slide_number_text)
    if match:
        return int(match.group())
    
    # If LLM returns something like "None", "no slide", etc.
    return None