import re

with open("gui.py", "r", encoding="utf-8") as f:
    content = f.read()

# We need to extract the data loading block (from "try:" up to "return\n        except Exception as e:")
# and the state initialization block (from "journal_state = []" up to "disabled_ar\n            })")

# Let's find the boundaries
start_data = content.find("        try:\n            wb = load_workbook(latest_file")
end_data = content.find("            self._set_status(f\"Error reading excel: {e}\", ERROR)\n            return") + len("            self._set_status(f\"Error reading excel: {e}\", ERROR)\n            return")

start_top = content.find("        top = tk.Toplevel(self)")
end_top = content.find("top.grab_set()") + len("top.grab_set()")

start_state = content.find("        # State Initialization")
end_state = content.find("disabled_ar\n            })") + len("disabled_ar\n            })")

if -1 in [start_data, end_data, start_top, end_top, start_state, end_state]:
    print("Could not find boundaries")
    exit(1)

data_code = content[start_data:end_data]
top_code = content[start_top:end_top]
state_code = content[start_state:end_state]

# We want to reorganize _on_journal so it looks like:
# top_code
#
# def _load_and_refresh():
#     nonlocal items, journal_state
#     latest_file = max(glob.glob(str(OUTPUT_DIR / "[Rr]econciliation_*.xlsx")), key=os.path.getctime)
#     ... (indented data_code, but items=[] replaced with items.clear(); items.extend(...))
#     ... (indented state_code, but journal_state=[] replaced with journal_state.clear())
#     if hasattr(self, '_render_page'): self._render_page(current_page[0])
#     if hasattr(self, '_update_pagination_buttons'): self._update_pagination_buttons()

print("Script created but not executed")
