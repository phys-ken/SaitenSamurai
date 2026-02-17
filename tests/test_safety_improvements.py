import unittest
import pandas as pd
from pathlib import Path
import tkinter as tk
from unittest.mock import MagicMock, patch
import sys
import shutil

# Add main_src to path to import modules
sys.path.append(str(Path(__file__).parent.parent / "main_src"))

from main_src.constants import ANSWER_KEY_FILE

# Mocking tkinter and other GUI elements for testing logic
class TestSafetyImprovements(unittest.TestCase):

    def setUp(self):
        # Create a temporary directory for testing
        self.test_dir = Path("tests/tmp_safety_check")
        self.test_dir.mkdir(parents=True, exist_ok=True)
        self.answer_key_path = self.test_dir / ANSWER_KEY_FILE

    def tearDown(self):
        # Clean up temporary directory
        if self.test_dir.exists():
            shutil.rmtree(self.test_dir)

    def test_empty_answer_key_detection(self):
        """Test detection of an empty (header-only) answer key."""
        # Create an empty DataFrame (only headers)
        df = pd.DataFrame(columns=["Question", "Answer", "Score"])
        df.to_excel(self.answer_key_path, index=False)

        # Logic to check if empty
        is_empty = False
        if self.answer_key_path.exists():
            df_read = pd.read_excel(self.answer_key_path)
            if len(df_read) == 0:
                is_empty = True
        
        self.assertTrue(is_empty, "Should detect empty answer key")

    def test_filled_answer_key_detection(self):
        """Test detection of a filled answer key."""
        # Create a filled DataFrame
        df = pd.DataFrame([{"Question": 1, "Answer": 1, "Score": 10}])
        df.to_excel(self.answer_key_path, index=False)

        # Logic to check if empty
        is_empty = False
        if self.answer_key_path.exists():
            df_read = pd.read_excel(self.answer_key_path)
            if len(df_read) == 0:
                is_empty = True
        
        self.assertFalse(is_empty, "Should NOT detect filled answer key as empty")

    @patch('subprocess.Popen')
    @patch('tkinter.messagebox.askyesno')
    def test_dialog_logic_flow(self, mock_askyesno, mock_popen):
        """Simulate the logic flow in _run_box_drawer_thread."""
        
        # Setup: Empty Answer Key
        df = pd.DataFrame(columns=["Question", "Answer", "Score"])
        df.to_excel(self.answer_key_path, index=False)
        
        # Mock user clicking "Yes"
        mock_askyesno.return_value = True

        # Simulated Logic from main_gui.py
        if self.answer_key_path.exists():
            df_read = pd.read_excel(self.answer_key_path)
            if len(df_read) == 0:
                if mock_askyesno("Title", "Message"):
                    import subprocess
                    subprocess.Popen(f'explorer "{self.test_dir}"')

        # Assertions
        mock_askyesno.assert_called_once()
        mock_popen.assert_called_once()
        # Verify the command opens the correct folder
        args, _ = mock_popen.call_args
        self.assertIn(str(self.test_dir), args[0])

if __name__ == '__main__':
    unittest.main()
