"""
prompt_templates.py
Centralized repository for LLM extraction prompts used by the grading MCP.
Now includes heuristics to detect when a student is asking for assistance.
"""

def build_excel_prompt(table_text: str) -> str:
    """
    Returns the Bedrock prompt for extracting Q/A pairs from Excel-like data,
    including confidence and reason detection.
    """
    return f"""
You are a precise data extraction and verification assistant.
You are given text from an Excel spreadsheet which is an **answer key** or student submission.
Your goal is to extract all question–answer pairs with their point values
and perform a consistency check on the total number of points.

Additionally, detect when a student's answer appears to be **seeking assistance**
—for example, comments like “I’m not sure”, “Can you explain this?”, or any text
that looks conversational toward the instructor. Such cases should be flagged.

### Rules:

1. The **very first cell (top-left corner)** in the spreadsheet should contain the total
   number of points (e.g., "4 pts" or "4pt").
   This represents the total of all question point values combined.

2. For each question:
   - The point value appears in a cell ending with "pt" or "pts".
   - The question text appears to the right of that cell.
   - The next one or more rows following the question contain the answer(s).

3. For each question–answer pair, output:
   - "points": string (e.g., "1pt")
   - "question": text
   - "answer": text (the first answer cell below the question)
   - "confidence":
       - "high" → only one clear, objective answer
       - "low" → if multiple answers, comments, or conversational tone appear
   - "reason":
       - "N/A" → if confidence is high
       - "student might be asking a question" → if answer looks like a message or request for help

4. After extracting all pairs, compute the **sum of all point values** and compare it with
   the total points in the first cell (top-left).
   Add a field `"sanity_check_passed": true` if they match exactly, otherwise `false`.

5. Return JSON in this exact format:
{{
  "total_points_cell": "4pts",
  "sanity_check_passed": true,
  "items": [
    {{
      "points": "1pt",
      "question": "How many people prefer dogs?",
      "answer": "0.4333333",
      "confidence": "high",
      "reason": "N/A"
    }},
    {{
      "points": "1pt",
      "question": "What is the probability that a person prefers cats?",
      "answer": "I'm not sure, maybe 0.4?",
      "confidence": "low",
      "reason": "student might be asking a question"
    }}
  ]
}}

Only return valid JSON. Do not include explanations or text outside the JSON.

### Spreadsheet Data:
{table_text}
"""


def build_word_prompt(document_text: str) -> str:
    """
    Returns the Bedrock prompt for extracting Q/A pairs from a Word math assignment.
    Adds logic to detect conversational or uncertain student responses.
    """
    return f"""
You are a math grading extraction assistant.
You are given text from a Word document containing questions, point values, and answers.

### Rules:

1. Each question starts with a line containing a point value, e.g. "(2 pts)".
2. The next line beginning with "Answer:" contains the numeric or short text answer.
3. Each question/answer block may be separated by blank lines.

4. When extracting, identify conversational answers:
   - Phrases like "I'm not sure", "maybe", "could it be", "I think", or direct questions to the instructor.
   - If such text appears, mark:
       "confidence": "low"
       "reason": "student might be asking a question"
   - Otherwise:
       "confidence": "high"
       "reason": "N/A"

Return JSON in this format:
{{
  "items": [
    {{
      "points": "2pt",
      "question": "Given that a person is staff, what is the probability they prefer dogs?",
      "answer": "0.35",
      "confidence": "high",
      "reason": "N/A"
    }}
  ]
}}

### Document Text:
{document_text}
"""