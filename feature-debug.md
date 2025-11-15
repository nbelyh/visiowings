# Enhanced visiowings Remote Debugging Implementation Plan

## Objective

Build a Windows-only remote debugging system enabling full control and inspection of Visio VBA code debugging from VS Code, with breakpoints, stepping, variable inspection, call stack, and session resilience.

---

## Architecture Overview

- **VS Code Debug Adapter**: Implements Debug Adapter Protocol (DAP) using either Python (`debugpy`-based) or TypeScript for seamless integration with VS Code.
- **Python Debug Bridge**: COM automation client managing Visio VBA debug contexts with thread-safe async handling of requests/events.
- **VBA COM Layer**: Direct COM manipulation of Visio’s VBA project — injecting breakpoints, reading debug state, managing execution.

---

## Detailed Tasks and Refinements

### Task 1: VS Code Debug Adapter

- Use a proven DAP library (e.g., `debugpy`) or lightweight DAP framework.
- Handle:
  - Multiple concurrent sessions with proper disambiguation.
  - Session attach/detach with reconnect logic.
  - VS Code launch and attach configs specifying Visio file and debugging options.

### Task 2: Establish COM Connection with Visio VBA

- Use `win32com.client.Dispatch('Visio.Application')` from `pywin32`.
- Access VBA projects/modules via `visio_app.VBE.VBProjects` and `VBComponents`.
- Handle password-protected projects with:
  - User credentials prompt or documented limitation.
  - Graceful failure and messaging if protected.

### Task 3: Breakpoint Management

- Inject breakpoints by replacing code lines with `Stop` statements.
- Store original lines in-memory and optionally encrypted temp files (avoid plaintext backups).
- Address edge cases:
  - Modules locked or protected.
  - Existing `Stop` statements.
  - Concurrent edits by other processes.
- Provide fallback to VBA’s native debugger API if available (subject to Visio COM support).

### Task 4: Debugging Event Monitoring & Notifications

- Use COM event sinks or polling of `VBE.Debugger` and execution state.
- If event sinks unavailable, implement polling with exponential backoff.
- Extract call stack info via `VBAProject.Debugger` properties or runtime introspection.
- Deliver events to VS Code via DAP events.

### Task 5: Variable & Expression Inspection

- Populate variable views through:
  - Dynamic VBA Watch window manipulation via COM.
  - Evaluate expressions using `Application.Evaluate` if safe.
- Document TBD limitations: full runtime inspection might need a custom VBA tracer (optional advanced feature).
- Return structured data conforming to DAP variable format.

### Task 6: Step Execution Simulation

- Attempt COM-based step control if accessible (explore `VBE.Commands` or similar interfaces).
- Otherwise, reliably use `SendKeys` with:
  - Windows API `FindWindow`, `SetForegroundWindow` for focus.
  - Retry logic on focus failure.
- Add configurable delay/timing parameters for key sending.

### Task 7: Asynchronous Communication & Thread Safety

- Encapsulate COM calls within a dedicated thread protected by mutexes.
- Use Python `asyncio` or `queue.Queue` to mediate between DAP commands and COM operations.
- Define message JSON protocol for inter-component communication.

### Task 8: Error Handling & Recovery

- Implement timeouts on COM calls; abort and notify if no response within 5 seconds.
- On failure or exit, remove injected breakpoints and restore original code reliably.
- Provide user-friendly error messages in VS Code UI.
- Log detailed traces using Python ’logging’ module.

### Task 9: Documentation & Testing

- Include:
  - VS Code launch.json debug configurations with examples.
  - pywin32 and Visio VBA permission setup guides.
  - Troubleshooting, including common COM errors and solutions.
  - Known limitations, including guarded projects and environment constraints.

- Tests:
  - Breakpoint injection/removal edge cases.
  - Session interruptions and reconnect handling.
  - Simulated lag and COM failure recovery.

---

## Additional Considerations

- Security: Do not write unencrypted VBA backups to disk unless user consented.
- Windows Only: Explicitly note dependency on Windows and COM, fallback/hints for other OS.
- UX: Plan VS Code custom debug views for call stack and variables reflecting VBA naming conventions.

---

## Acceptance Criteria

- Breakpoints can be set, hit, and removed, reflected in Visio VBA.
- Step commands work stably with minimal delay or focus loss.
- Variables and call stack accurately reported in VS Code during paused execution.
- Sessions survive VS Code restart, with option to reconnect.
- No unhandled crashes or VBA project corruptions occur.

---

## Suggested Technologies

- Python 3.x with `pywin32` for COM interaction.
- Debug Adapter Protocol server implemented per VS Code specs (Python or Node.js).
- Threading or async event-driven architecture for COM communication.
- Optional sendkeys package or Windows API for keyboard simulation.

---

**Author:** visiowings Development Team
**Date:** 15.11.2025
