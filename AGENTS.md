# Autonomous Execution & Auto-Approve Guidelines (AGENTS.md)

## Auto-Approve & Auto-Allow Policy
- **Mode**: Full Autonomous & Auto-Approve (Auto Allow All).
- **Interactive Approval**: DISABLED. Never pause to ask the user for permission.
- **Allowed Actions**:
  - Full terminal execution (pwsh / bash / cmd / python / scripts).
  - Complete filesystem operations (read, write, edit, overwrite, delete temp files).
  - Data processing and report generation end-to-end.
  - Background processes and subagent invocations.
- Always execute plans fully from start to finish and verify all results autonomously.
