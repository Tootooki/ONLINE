import json
from pathlib import Path


WORKFLOW_IN = Path("/tmp/reimb_workflow_check1.json")
WORKFLOW_OUT = Path("/tmp/reimb_workflow_v2_payload.json")
NODE_CODE = Path("scripts/reimbursement_27_export_v2.js")


def main() -> None:
    workflow = json.loads(WORKFLOW_IN.read_text())
    code = NODE_CODE.read_text()

    found = False
    for node in workflow["nodes"]:
        if node.get("name") == "REIMBURSEMENT_27_EXPORT":
            node["parameters"]["jsCode"] = code
            node["position"] = [1040, 48]
            found = True
            break

    if not found:
        workflow["nodes"].append(
            {
                "parameters": {"jsCode": code},
                "id": "7fc89ccb-a6e2-446b-bbb8-c88f0ee9c27b",
                "name": "REIMBURSEMENT_27_EXPORT",
                "type": "n8n-nodes-base.code",
                "typeVersion": 2,
                "position": [1040, 48],
            }
        )

    workflow.setdefault("connections", {})["MERGE"] = {
        "main": [[{"node": "REIMBURSEMENT_27_EXPORT", "type": "main", "index": 0}]]
    }

    settings = workflow.get("settings", {})
    payload = {
        "name": workflow["name"],
        "nodes": workflow["nodes"],
        "connections": workflow["connections"],
        "settings": {k: settings[k] for k in ["executionOrder"] if k in settings},
    }
    if workflow.get("staticData") is not None:
        payload["staticData"] = workflow.get("staticData")
    if workflow.get("pinData") is not None:
        payload["pinData"] = workflow.get("pinData")

    WORKFLOW_OUT.write_text(json.dumps(payload, indent=2))
    print(WORKFLOW_OUT)


if __name__ == "__main__":
    main()
