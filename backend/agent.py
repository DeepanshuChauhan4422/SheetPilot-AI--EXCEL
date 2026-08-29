import os
import json
import time

from dotenv import load_dotenv

from langchain_core.messages import (
    SystemMessage,
    HumanMessage,
    AIMessage,
    ToolMessage
)

from langchain_mistralai import ChatMistralAI

from tools import (
    update_cell,
    write_range,
    read_range,
    clear_range,
    format_range,
    add_formula,
    sort_range,
    create_chart
)


# ============================================================
# ENVIRONMENT
# ============================================================

load_dotenv()

MISTRAL_API_KEY = os.getenv("MISTRAL_API_KEY")

if not MISTRAL_API_KEY:
    raise RuntimeError(
        "MISTRAL_API_KEY is missing from backend/.env"
    )


# ============================================================
# MISTRAL MODEL
# ============================================================

llm = ChatMistralAI(
    model="mistral-small-latest",
    temperature=0,
    api_key=MISTRAL_API_KEY
)


# ============================================================
# EXCEL TOOLS
# ============================================================

excel_tools = [
    update_cell,
    write_range,
    read_range,
    clear_range,
    format_range,
    add_formula,
    sort_range,
    create_chart
]

llm_with_tools = llm.bind_tools(excel_tools)


# ============================================================
# SYSTEM PROMPT
# ============================================================

SYSTEM_PROMPT = """
You are SheetPilot AI, an intelligent Excel assistant.

You control Microsoft Excel through the available tools.

============================================================
CORE RULES
============================================================

1. Understand normal human language. The user does not need
   to know Excel tool names.

2. Always use Excel tools when the user asks you to read,
   write, modify, sort, format, clear, calculate, or create
   a chart in Excel.

3. Never invent Excel data, ranges, columns, or values.

4. Use the Excel context supplied by the frontend whenever
   it is relevant.

5. Keep final responses short and clear.

6. Never perform the same successful Excel operation twice.

7. Do not modify unrelated cells.

8. Preserve existing formulas, formatting, and unrelated
   data whenever possible.

9. After a successful Excel modification, stop processing
   that operation.


============================================================
NATURAL LANGUAGE
============================================================

Understand commands such as:

"revenue ko highest se lowest karo"
-> sort_range

"expenses ko low to high sort karo"
-> sort_range

"revenue ka total batao"
-> calculate from Excel context

"revenue ka total H5 mein daalo"
-> add_formula

"A1:F1 ko bold karo"
-> format_range

"header ko yellow karo"
-> format_range

"H2 ko clear karo"
-> clear_range

"A1:B4 ka chart bana do"
-> create_chart


============================================================
READING / ANALYSIS
============================================================

When the user asks a question about Excel data:

1. First inspect the Excel context supplied by the frontend.

2. If the required data is already present, analyze that
   data directly.

3. If the required data is not present, use read_range.

4. Analyze only actual Excel data.

5. Never invent values.

6. For total, average, highest, lowest, count, or similar
   questions, use ALL relevant data rows.


============================================================
COLUMN UNDERSTANDING
============================================================

Match the user's wording with the actual Excel headers.

Examples:

"expense" -> "expenses"
"expenses" -> "expenses"
"sale" -> "sales"
"sales" -> "sales"
"revenue" -> "revenue"
"profit" -> "profit"

Use the actual header name when calling a tool whenever
possible.

Never invent a column.

If the requested column does not exist, do not guess another
column.


============================================================
RANGE UNDERSTANDING
============================================================

If the user explicitly provides a range, use that range.

Example:

"Read A1:F4"
-> read_range with A1:F4

If the user does not provide a range, use the Excel context
provided by the frontend.

If the frontend identifies a current table/region, prefer
that table when appropriate.

Do not modify unrelated cells.


============================================================
SORTING
============================================================

When the user asks to sort Excel data:

1. Identify the requested column.

2. Identify the complete table/range from Excel context.

3. Call sort_range exactly once.

4. Use the actual header name when possible.

Examples:

"Sort revenue from highest to lowest"
-> column = "revenue"
-> order = "descending"

"Sort revenue from lowest to highest"
-> column = "revenue"
-> order = "ascending"

"Sort expenses highest to lowest"
-> column = "expenses"
-> order = "descending"

"Sort expenses lowest to highest"
-> column = "expenses"
-> order = "ascending"

IMPORTANT:

- Never use write_range to perform sorting.
- Never manually rearrange rows.
- Never sort only one column.
- sort_range must sort complete rows together.
- Never call sort_range twice for one request.
- After successful sorting, stop.


============================================================
FORMULAS AND CALCULATIONS
============================================================

Understand requests such as:

"total revenue"
"sum of revenue"
"average revenue"
"highest revenue"
"lowest revenue"
"total profit"
"calculate expenses"


If the user only asks for the result:

1. Use the Excel context.
2. Calculate the answer.
3. Do NOT modify Excel unnecessarily.


If the user explicitly asks to put the result/formula into
Excel:

1. Identify the actual Excel column and rows.
2. Use add_formula.
3. Write a valid Excel formula.


Examples:

If revenue is B2:B4:

total revenue
-> =SUM(B2:B4)

average revenue
-> =AVERAGE(B2:B4)

highest revenue
-> =MAX(B2:B4)

lowest revenue
-> =MIN(B2:B4)


If profit is D2:D4:

total profit
-> =SUM(D2:D4)

average profit
-> =AVERAGE(D2:D4)


NEVER generate:

=SUM(revenue)
=SUM(B)
=SUM()

Never invent row numbers.

Never use a header name directly inside an Excel formula.


============================================================
WRITING
============================================================

For a single cell:

Use update_cell.

Example:

"Put 5000 in H2"
-> update_cell


For multiple cells/ranges:

Use write_range.

Example:

"Write this table to A1:C4"
-> write_range


Do not use write_range for sorting.


============================================================
CLEARING
============================================================

When the user asks to clear/delete/remove cell contents:

Use clear_range.

Examples:

"Clear H2"
"Clear H2:H5"
"Remove data from A1:C4"

Do not modify formatting unless the user explicitly asks.


============================================================
FORMATTING
============================================================

When the user asks to format Excel:

Use format_range.

Supported operations include:

- bold
- italic
- fill_color
- font_color
- font_size
- number_format
- horizontal_alignment
- borders
- autofit_columns
- autofit_rows


Examples:

"Make A1:F1 bold"
-> format_range

"Make the header yellow"
-> format_range

"Make revenue text green"
-> format_range

"Autofit columns"
-> format_range


If the user asks for multiple formatting changes in one
request, perform the required formatting without modifying
unrelated data.


============================================================
CHARTS
============================================================

When the user asks for a chart:

Use create_chart.

Examples:

"Create a revenue chart"
"Make a chart of A1:B4"
"Create a bar chart for sales"
"Show revenue as a chart"

Do not manually write chart data unless required.


============================================================
TOOL SAFETY
============================================================

For every user request:

1. Determine the exact operation required.

2. Use the minimum number of tools necessary.

3. Never repeat a successful tool call.

4. Never perform unrelated Excel operations.

5. Never invent missing Excel information.

6. If required information is genuinely missing, ask the
   user for it instead of guessing.

7. If a modification tool succeeds, do not call another
   modification tool unless the user's original request
   explicitly requires multiple different modifications.

8. For a single operation, exactly one corresponding tool
   call should normally be made.


============================================================
FINAL RESPONSE
============================================================

Keep successful responses short.

Examples:

"Done! ✅"

"Done! Excel data sorted successfully. ✅"

"Done! Formula added successfully. ✅"

"Done! Excel formatting updated successfully. ✅"

"Done! Chart created successfully. 📊"

Do not expose internal tool calls or implementation details.
"""


# ============================================================
# TOOL CALL CONVERTER
# ============================================================

def get_tool_calls(response):

    calls = []

    for call in response.tool_calls:

        calls.append({
            "tool": call["name"],
            "tool_call_id": call["id"],
            "arguments": call["args"]
        })

    return calls


# ============================================================
# RATE LIMIT DETECTION
# ============================================================

def is_rate_limit_error(error):

    error_text = str(error).lower()

    rate_limit_keywords = [
        "rate_limited",
        "rate limit",
        "rate-limit",
        "too many requests",
        "429",
        "1300"
    ]

    return any(
        keyword in error_text
        for keyword in rate_limit_keywords
    )


# ============================================================
# SAFE LLM INVOKE
# ============================================================

def invoke_llm(messages):

    max_retries = 4

    delays = [
        1.5,
        3.0,
        6.0,
        10.0
    ]

    for attempt in range(max_retries):

        try:

            return llm_with_tools.invoke(messages)

        except Exception as error:

            print(
                f"Mistral request failed "
                f"(attempt {attempt + 1}/{max_retries})"
            )

            print(error)

            if not is_rate_limit_error(error):
                raise

            if attempt == max_retries - 1:
                raise

            delay = delays[attempt]

            print(
                f"Rate limit detected. "
                f"Retrying in {delay} seconds..."
            )

            time.sleep(delay)

    raise RuntimeError(
        "Mistral request failed after retries."
    )


# ============================================================
# FIRST AGENT REQUEST
# ============================================================

def run_agent(user_message: str):

    messages = [

        SystemMessage(
            content=SYSTEM_PROMPT
        ),

        HumanMessage(
            content=user_message
        )

    ]

    response = invoke_llm(messages)

    if response.tool_calls:

        return {
            "type": "tool_calls",
            "tool_calls": get_tool_calls(response)
        }

    return {
        "type": "message",
        "content": response.content
    }


# ============================================================
# SUCCESSFUL MODIFICATION DETECTION
# ============================================================

def get_successful_modification(
    tool_results
):

    modification_actions = {
        "update_cell",
        "write_range",
        "clear_range",
        "format_range",
        "add_formula",
        "sort_range",
        "create_chart"
    }

    for result in tool_results:

        result_data =result.get("result")

        if not isinstance(
            result_data,
            dict
        ):
            continue

        action =result_data.get("action")

        success =result_data.get("success")

        if (
            success is True
            and action in modification_actions
        ):

            return action

    return None


# ============================================================
# SUCCESS MESSAGE
# ============================================================

def get_success_message(
    action
):

    messages = {

        "update_cell":
            "Done! Cell updated successfully. ✅",

        "write_range":
            "Done! Excel data updated successfully. ✅",

        "clear_range":
            "Done! Excel range cleared successfully. ✅",

        "format_range":
            "Done! Excel formatting updated successfully. ✅",

        "add_formula":
            "Done! Formula added successfully. ✅",

        "sort_range":
            "Done! Excel data sorted successfully. ✅",

        "create_chart":
            "Done! Chart created successfully. 📊"

    }

    return messages.get(
        action,
        "Done! Excel operation completed successfully. ✅"
    )


# ============================================================
# CONTINUE AGENT
# ============================================================

def continue_agent(
    user_message: str,
    previous_response: dict,
    tool_results: list
):

    # ========================================================
    # STOP AFTER SUCCESSFUL MODIFICATION
    # ========================================================

    successful_action =get_successful_modification(
            tool_results
        )

    if successful_action:

        return {
            "type": "message",
            "content":
                get_success_message(
                    successful_action
                )
        }


    # ========================================================
    # BUILD MESSAGE HISTORY
    # ========================================================

    messages = [

        SystemMessage(
            content=SYSTEM_PROMPT
        ),

        HumanMessage(
            content=user_message
        )

    ]


    # ========================================================
    # PREVIOUS TOOL CALLS
    # ========================================================

    previous_tool_calls = []

    for call in previous_response.get(
        "tool_calls",
        []
    ):

        previous_tool_calls.append({

            "name":
                call["tool"],

            "args":
                call["arguments"],

            "id":
                call["tool_call_id"],

            "type":
                "tool_call"

        })


    if previous_tool_calls:

        messages.append(

            AIMessage(

                content="",

                tool_calls=
                    previous_tool_calls

            )

        )


    # ========================================================
    # TOOL RESULTS
    # ========================================================

    for result in tool_results:

        tool_call_id =result.get(
                "tool_call_id"
            )

        tool_result =result.get(
                "result"
            )


        if isinstance(
            tool_result,
            (dict, list)
        ):

            tool_result =json.dumps(
                    tool_result,
                    ensure_ascii=False
                )

        else:

            tool_result =str(
                    tool_result
                )


        messages.append(

            ToolMessage(

                content=
                    tool_result,

                tool_call_id=
                    tool_call_id

            )

        )


    # ========================================================
    # ASK AI AGAIN
    # ========================================================

    response =invoke_llm(
            messages
        )


    if response.tool_calls:

        return {

            "type":
                "tool_calls",

            "tool_calls":
                get_tool_calls(
                    response
                )

        }


    return {

        "type":
            "message",

        "content":
            response.content

    }