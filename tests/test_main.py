import os
import pandas as pd
import subprocess

# Define file paths
INPUT_FILE = "test_input.xlsx"
OUTPUT_FILE = "test_output.xlsx"

def setup_module(module):
    """Set up a dummy Excel file for testing."""
    data = {'Name': ['Alice', 'Bob', 'Charlie'],
            'Department': ['Sales', 'Marketing', 'Sales']}
    df = pd.DataFrame(data)
    df.to_excel(INPUT_FILE, index=False)

def teardown_module(module):
    """Clean up dummy Excel files after testing."""
    if os.path.exists(INPUT_FILE):
        os.remove(INPUT_FILE)
    if os.path.exists(OUTPUT_FILE):
        os.remove(OUTPUT_FILE)

def run_cli_command(action, input_file, output_file, **kwargs):
    """Helper function to run the main.py script."""
    command = ["python3", "main.py", action, "-i", input_file, "-o", output_file]
    for key, value in kwargs.items():
        command.extend([f"--{key.replace("_", "-")}", str(value)])
    
    # Run the command within the virtual environment
    result = subprocess.run(["bash", "-c", f"source venv/bin/activate && {' '.join(command)}"], capture_output=True, text=True, check=False)
    return result

def test_filter_action():
    """Test the filter action."""
    result = run_cli_command(
        "filter",
        INPUT_FILE,
        OUTPUT_FILE,
        column="Department",
        value="Sales"
    )
    assert result.returncode == 0
    assert "Action 'filter' completed successfully" in result.stdout

    # Verify output file content
    df_output = pd.read_excel(OUTPUT_FILE)
    assert len(df_output) == 2
    assert "Alice" in df_output['Name'].values
    assert "Charlie" in df_output['Name'].values
    assert "Bob" not in df_output['Name'].values
