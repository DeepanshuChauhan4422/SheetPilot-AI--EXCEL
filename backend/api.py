from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional, Any

from agent import run_agent, continue_agent


# ============================================================
# FASTAPI APPLICATION
# ============================================================

app = FastAPI(
    title="SheetPilot AI API",
    description="AI backend for SheetPilot AI Excel Agent",
    version="2.0.0"
)


# ============================================================
# CORS
# ============================================================

app.add_middleware(
    CORSMiddleware,

    allow_origins=["*"],

    allow_credentials=False,

    allow_methods=["*"],

    allow_headers=["*"],
)


# ============================================================
# REQUEST MODEL
# ============================================================

class ChatRequest(BaseModel):

    message: str

    tool_results: Optional[
        list[dict[str, Any]]
    ] = None

    previous_response: Optional[
        dict[str, Any]
    ] = None


# ============================================================
# HOME
# ============================================================

@app.get("/")
def home():

    return {
        "status": "success",
        "message": "SheetPilot AI API is running 🚀"
    }


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/health")
def health():

    return {
        "status": "success",
        "message": "SheetPilot AI backend is healthy ✅"
    }


# ============================================================
# CHAT
# ============================================================

@app.post("/chat")
def chat(
    request: ChatRequest
):

    try:

        # ====================================================
        # FIRST AI REQUEST
        # ====================================================

        if (
            request.tool_results is None
            and request.previous_response is None
        ):

            result =run_agent(
                    request.message
                )

            return {

                "status":
                    "success",

                "result":
                    result
            }


        # ====================================================
        # CONTINUE AFTER EXCEL TOOL
        # ====================================================

        if (
            request.tool_results is not None
            and request.previous_response is not None
        ):

            result =continue_agent(

                    user_message=
                        request.message,

                    previous_response=
                        request.previous_response,

                    tool_results=
                        request.tool_results
                )

            return {

                "status":
                    "success",

                "result":
                    result
            }


        # ====================================================
        # INVALID CONTINUATION
        # ====================================================

        return {

            "status":
                "error",

            "message":
                "Invalid agent continuation request."
        }


    except Exception as error:

        return {

            "status":
                "error",

            "message":
                str(error)
        }