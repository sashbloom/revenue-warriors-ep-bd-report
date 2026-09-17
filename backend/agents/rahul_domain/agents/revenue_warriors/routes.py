"""
Revenue Warriors dispatch routes.
Add these to backend/agents/rahul_domain/router.py
"""

from fastapi import APIRouter, UploadFile, File, HTTPException
from pydantic import BaseModel

from backend.core.graph_client import get_graph_client
from backend.agents.rahul_domain.agents.revenue_warriors.dispatch import (
    build_all_emails,
    extract_all_sheets,
    merge_cc,
)

router = APIRouter(prefix="/rahul/revenue-warriors", tags=["revenue-warriors"])


# ── Models ──

class DraftRequest(BaseModel):
    to: list[str]
    cc: list[str] = []
    subject: str
    body_html: str


class SendRequest(BaseModel):
    message_id: str


class SendAllRequest(BaseModel):
    message_ids: list[str]


# ── Routes ──

@router.post("/parse")
async def parse_workbook(file: UploadFile = File(...), week: str = "", cc: str = ""):
    """Upload workbook → get back all email payloads ready for preview."""
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Upload an .xlsx file")

    content = await file.read()
    cc_list = [e.strip() for e in cc.split(",") if e.strip()] if cc else []
    result = build_all_emails(content, week_override=week, cc=cc_list)
    return result


@router.get("/status")
async def graph_status():
    """Whether Graph credentials are present, or we are in dry-run mode."""
    client = get_graph_client()
    return {
        "configured": client.configured,
        "dry_run": client.dry_run,
        "missing_vars": client.missing_vars,
        "sender": client.sender_email or None,
    }


@router.post("/draft")
async def create_draft(req: DraftRequest):
    """Create one draft email in Isabelle's Outlook."""
    client = get_graph_client()
    msg_id = client.create_draft(req.to, merge_cc(req.cc), req.subject, req.body_html)
    if msg_id:
        return {
            "message_id": msg_id,
            "status": "draft_created",
            "dry_run": client.dry_run,
        }
    raise HTTPException(500, "Failed to create draft")


@router.post("/send")
async def send_draft(req: SendRequest):
    """Send one draft by message_id."""
    client = get_graph_client()
    ok = client.send_draft(req.message_id)
    if ok:
        return {"status": "sent", "dry_run": client.dry_run}
    raise HTTPException(500, "Failed to send draft")


@router.post("/send-all")
async def send_all(req: SendAllRequest):
    """Send multiple drafts at once."""
    client = get_graph_client()
    results = []
    for mid in req.message_ids:
        ok = client.send_draft(mid)
        results.append({"message_id": mid, "status": "sent" if ok else "failed"})
    sent = sum(1 for r in results if r["status"] == "sent")
    return {
        "results": results,
        "sent": sent,
        "total": len(results),
        "dry_run": client.dry_run,
    }


@router.delete("/draft")
async def delete_draft(req: SendRequest):
    """Delete one draft."""
    client = get_graph_client()
    ok = client.delete_draft(req.message_id)
    return {"status": "deleted" if ok else "failed"}
