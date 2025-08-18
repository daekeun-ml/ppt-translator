#!/usr/bin/env python3
"""
Main entry point for the PowerPoint Translator FastMCP server.
This ensures the server runs from the correct directory regardless of how it's invoked.
"""

import os
import sys
from pathlib import Path

def main():
    # Get the directory where this script is located
    script_dir = Path(__file__).parent.absolute()
    
    # Change to the script directory to ensure relative imports work
    os.chdir(script_dir)
    
    # Add the script directory to Python path
    if str(script_dir) not in sys.path:
        sys.path.insert(0, str(script_dir))
    
    # Import and run the FastMCP server
    try:
        from fastmcp_server import main as fastmcp_main
        fastmcp_main()
    except ImportError as e:
        print(f"Error importing fastmcp_server: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
