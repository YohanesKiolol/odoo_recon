import unittest
from pathlib import Path

class TestUIModules(unittest.TestCase):
    def test_theme_exports(self):
        from ui.theme import (
            FONT_FAMILY, FONT_BODY, FONT_MONO, _EMOJI_FONT,
            BG, PANEL, SIDEBAR_BG, PREVIEW_BG, BORDER, BORDER_DARK,
            ACCENT, ACCENT_DARK, SUCCESS, ERROR, WARN, TEXT, MUTED, WHITE,
            BANK_BADGE_COLS, TYPE_BADGE_INFO, init_fonts
        )
        self.assertEqual(BG, "#F4F5F8")
        self.assertEqual(PANEL, "#FFFFFF")
        self.assertEqual(ACCENT, "#6D28D9")
        self.assertIn("BCA", BANK_BADGE_COLS)
        self.assertIn("bank_only", TYPE_BADGE_INFO)

    def test_widgets_exports(self):
        from ui.widgets import (
            _open_path, _maximize_window, _center_modal_on_parent,
            CTkDateInput, create_custom_dropdown
        )
        self.assertTrue(callable(_open_path))
        self.assertTrue(callable(_maximize_window))
        self.assertTrue(callable(_center_modal_on_parent))
        self.assertTrue(callable(create_custom_dropdown))

    def test_modals_exports(self):
        from ui.modals import (
            open_breakdown_modal, open_cleanup_modal, open_conflict_modal,
            open_sales_portal_modal, open_discrepancy_inspection_modal,
            open_manual_match_modal, open_journal_modal
        )
        self.assertTrue(callable(open_breakdown_modal))
        self.assertTrue(callable(open_cleanup_modal))
        self.assertTrue(callable(open_conflict_modal))
        self.assertTrue(callable(open_sales_portal_modal))
        self.assertTrue(callable(open_discrepancy_inspection_modal))
        self.assertTrue(callable(open_manual_match_modal))
        self.assertTrue(callable(open_journal_modal))

    def test_views_exports(self):
        from ui.views import SidebarView, LocalDashboardView, CloudDashboardView, LogConsoleView
        self.assertTrue(issubclass(SidebarView, object))
        self.assertTrue(issubclass(LocalDashboardView, object))
        self.assertTrue(issubclass(CloudDashboardView, object))
        self.assertTrue(issubclass(LogConsoleView, object))

    def test_controllers_exports(self):
        from ui.controllers import ReconController
        self.assertTrue(callable(ReconController.tag_line))
        self.assertEqual(ReconController.tag_line("✅ Done"), "ok")
        self.assertEqual(ReconController.tag_line("❌ Failed"), "err")
        self.assertEqual(ReconController.tag_line("⚠️ Warn"), "warn")

    def test_local_dashboard_render_drill_grid(self):
        import tkinter as tk
        from ui.views import LocalDashboardView
        root = tk.Tk()
        root.withdraw()
        try:
            view = LocalDashboardView(root)
            dummy_rows = [("BCA (Main)", "2026-07-01 to 2026-07-06", "Available")]
            view._render_drill_grid(dummy_rows)
            self.assertEqual(len(view.drill_grid.winfo_children()), 6) # 3 headers + 3 cells
        finally:
            root.destroy()

    def test_gui_imports(self):
        import gui
        self.assertTrue(hasattr(gui, "App"))
        self.assertTrue(hasattr(gui.App, "_show_bank_breakdown_modal"))
        self.assertTrue(hasattr(gui.App, "_on_sync_cloud"))
        self.assertTrue(hasattr(gui.App, "_on_manual_match"))
        self.assertTrue(hasattr(gui.App, "_on_journal"))


if __name__ == "__main__":
    unittest.main()


