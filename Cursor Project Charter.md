# 🚀 Cursor Project Charter
(Version 1.0)

This document defines the behavior, style, and project rules Cursor AI must follow while assisting in this project.

## 1. Development Philosophy
- Prioritize **efficient, low-quota** API usage.
- Avoid unnecessary complexity; keep the architecture clean and modular.
- Default to **batch operations** over row-by-row processing unless otherwise instructed.

## 2. Code Change Process
- **ASK before making major structural changes** (e.g., refactoring, redesigning modules).
- Minor fixes and improvements are allowed without prior confirmation.

## 3. Style & Documentation
- Add **clear, practical comments** on all non-trivial logic.
- Follow consistent naming conventions (`camelCase` for variables and functions).
- Default to clear, readable structure over clever one-liners.

## 4. Quota & Resource Usage
- Be **mindful of quota-heavy operations** (e.g., GmailApp.search, SpreadsheetApp.getRange loops).
- Suggest optimizations if a quota-heavy operation is detected.
- If quota risk is moderate or high, alert me immediately.

## 5. Error Handling
- Wrap all major API calls in **try/catch blocks** unless speed outweighs reliability.
- Log meaningful error messages to help debug quickly.

## 6. Confirmation Triggers
- For **destructive operations** (e.g., mass delete, irreversible edits), STOP and seek explicit confirmation.
- For large/batch updates (> 1000 rows or > 100 emails), notify before proceeding.

## 7. Performance and Caching
- Default to **caching** results of expensive lookups within a session if applicable.
- Always recommend a caching layer if response speed is critical.

## 8. Security Best Practices
- NEVER hardcode API keys, tokens, or sensitive credentials directly in the code.
- If sensitive access is needed, recommend secure storage options.

## 9. Communication Style
- Be **concise, direct, and solution-oriented**.
- Suggest improvements, but respect project scope boundaries unless expansion is requested.

## 10. Evolution
- This Charter can and should evolve.
- If a pattern of better practices is recognized, propose a Charter amendment.

---

# ✅ Working Agreement:
Cursor agrees to reference and operate under this Charter for the duration of this project.