"""
Utility functions for Dochameleon.
"""

from pathlib import Path
from typing import List


def find_files(input_dir: Path, extension: str, exclude_patterns: List[str] = None) -> List[Path]:
    """Find files with given extension, excluding certain patterns."""
    if exclude_patterns is None:
        exclude_patterns = ['_style', '_temp', '.backup']
    
    files = []
    for file in input_dir.glob(f"*.{extension}"):
        skip = False
        for pattern in exclude_patterns:
            if pattern in file.stem:
                skip = True
                break
        if not skip:
            files.append(file)
    return files
