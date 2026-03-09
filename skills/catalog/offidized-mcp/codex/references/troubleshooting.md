# Troubleshooting

## Bad Request: No valid session ID provided

- Ensure initialize request succeeded before tool calls.
- Ensure proxy forwards `mcp-session-id`.
- Use correct endpoint URL (avoid `/mcp/mcp`).

## Bad Request: Server not initialized

- Reconnect and re-initialize client session.
- Ensure runtime keeps session continuity for subsequent requests.

## Inspector connection failures

1. Verify endpoint: `https://<host>/mcp`
2. Verify bearer token if enabled (`MCP_SHARED_TOKEN`).
3. Verify `GET /healthz` returns success.
