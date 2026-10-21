# v1.8 Excel AI integration

Chatbox 1.20.3 remains the model client. Its provider implementations, fetch functions, authentication, Chatbox-specific headers and native request loop are unchanged. The wrapper retains the official app name/product name/version so integration branding does not leak into the native application identity. `Chatbox/provenance.json` records exact, hash-pinned integration patch anchors. Electron remains the existing portable runtime; automatic Chatbox self-update remains disabled.

## Interaction

Work scope is inserted in the actual native InputBox toolbar beside MCP and web search. No persistent context cards or prompt buttons occupy the conversation. An overflow menu contains formula help, analysis, manual drafts, undo and live window mode. Pending proposals are isolated by conversation and share one Apply action. Read and proposal tools cannot call apply/undo.

The Excel task pane reserves layout space. Chromium stays an Excel-owned top-level HWND, positioned on the pane without WS_CHILD/SetParent. This avoids Excel consuming the cross-process child window's keystrokes. Resize/move tracking never activates the window; minimized/hidden panes hide their owned window. Detaching restores original window style and ownership. The same process, renderer, draft, chat and snapshots survive mode changes.

## Transport and capabilities

Native Chatbox launches `stdio.cjs` with its native StdioClientTransport. The server uses the official MIT MCP SDK, communicates only JSON-RPC on stdout, and connects to an unpredictable per-launch named pipe with a session token. There are no TCP listeners in the production integration. The existing Excel host pipe is restricted to the current Windows user. Stdio child teardown follows stdin/pipe closure.

- MCP SDK: https://github.com/modelcontextprotocol/typescript-sdk — npm `@modelcontextprotocol/sdk@1.32.1`, lockfile integrity pinned.
- Excel MCP: https://github.com/haris-musa/excel-mcp-server — commit `41c484b6930df70d7e783be6eaa170c6e484e7a0`, MIT copyright retained. `AiRange.cs` adapts `refs.py` bounded A1 normalization and worksheet limits, and `formatting.py` hex-color validation. CellFormat field names inform the native formatting adapter. We do not ship or claim the entire upstream Python/openpyxl engine; active unsaved Excel workbooks must use the host COM adapter.
- UI references: https://support.microsoft.com/en-us/office/edit-with-copilot-in-excel-a2fd6fe4-97ac-416b-b89a-22f4d1357c7a/ — compact in-context actions and grouped review informed the design; Chatbox retains its own appearance and navigation.

Nine tools: `excel_read`, `excel_find`, `excel_statistics`, `excel_inspect`, `excel_propose_formula`, `excel_propose_values`, `excel_propose_formulas`, `excel_propose_format`, `excel_propose_chart`.

## Data and modification model

Sparse SpecialCells areas are read in blocks of at most 16,384 cells. A Find fallback handles fragmented ranges. Counts are 64-bit through the whole worksheet. JSONL snapshots use exclusive FileStreams with DeleteOnClose, bounded active retention (eight snapshots), and streaming pagination/statistics. Model context includes at most 120 sample cells plus exact statistics and full authorized range metadata. This removes the arbitrary 200k limit without allocating the worksheet's 17 billion positions. Resource exhaustion remains an explicit failure rather than silent truncation.

Proposals capture original target values/formulas or format properties. Apply preflights every proposal, rejects overlapping same-kind writes, disables Excel event handlers during application, and rolls back on a subsequent error. Only successful batches replace undo. Reverting checks for manual edits first. Table, array, merged, protected and hidden write targets remain excluded. Uniform formatting is reversible without flattening mixed styles. Graph creation uses Excel's actual chart objects. None of these paths issue model HTTP requests.

## Attachments

The official renderer expects File.path. The existing portable Electron uses webUtils.getPathForFile instead. A minimal compatibility getter delegates through the isolated preload; it supplies paths to the original native PDF/Office/text parser, including picked and dropped files. No replacement upload endpoint or request-header injection is introduced.
