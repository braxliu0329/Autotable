import sys
import os
import json
import logging
from unittest.mock import MagicMock

# Add parent directory to sys.path to import Autotable modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from autotable import AutoTable

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def test_prompt_generation():
    # Mock LLM Client
    mock_llm = MagicMock()
    mock_llm.chat_completion.return_value = '{"__identity__": "Test User", "{{ID_001}}": "Value1"}'
    
    at = AutoTable("dummy_kb", "dummy_template", mock_llm)
    
    # Test Case 1: HTML Context
    html_context = """
    <table border="1">
      <tr>
        <td rowspan="2">Header 1</td>
        <td>{{ID_001}}</td>
      </tr>
      <tr>
        <td>{{ID_002}}</td>
      </tr>
    </table>
    """
    
    kb = {"Sheet1": [["Key", "Value"], ["Header 1", "Value1"]]}
    id_map = {"{{ID_001}}": "original_1", "{{ID_002}}": "original_2"}
    
    print("\n--- Testing HTML Context ---")
    at.analyze_tables_with_llm(html_context, kb, id_map)
    
    # Check the call arguments to see the prompt
    args, _ = mock_llm.chat_completion.call_args
    messages = args[0]
    prompt_content = messages[1]['content']
    
    if "HTML 结构分析指南" in prompt_content:
        print("SUCCESS: Prompt contains HTML instructions.")
    else:
        print("FAILURE: Prompt missing HTML instructions.")
        print(f"Prompt content preview: {prompt_content[:500]}")
        
    if "rowspan" in prompt_content and "colspan" in prompt_content:
        print("SUCCESS: Prompt mentions rowspan/colspan.")
    else:
        print("FAILURE: Prompt missing structural keywords.")

    # Test Case 2: Markdown Context
    md_context = """
    | Header 1 | {{ID_001}} |
    | Header 2 | {{ID_002}} |
    """
    
    print("\n--- Testing Markdown Context ---")
    at.analyze_tables_with_llm(md_context, kb, id_map)
    
    args, _ = mock_llm.chat_completion.call_args
    messages = args[0]
    prompt_content = messages[1]['content']
    
    if "Markdown 结构分析指南" in prompt_content:
        print("SUCCESS: Prompt contains Markdown instructions.")
    else:
        print("FAILURE: Prompt missing Markdown instructions.")

if __name__ == "__main__":
    test_prompt_generation()
