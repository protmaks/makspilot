#!/usr/bin/env python3
"""
Version Update Script for MaxPilot
Updates version information across all HTML files in the project
"""

import json
import os
import re
from pathlib import Path


def load_version_config():
    """Load version configuration from version.json"""
    try:
        with open('version.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print("Error: version.json not found!")
        return None
    except json.JSONDecodeError as e:
        print(f"Error parsing version.json: {e}")
        return None


def update_html_file(file_path, old_version, new_version):
    """Update version in a single HTML file"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Replace version in h1 tags
        pattern = r'(<h1>[^<]*?)(' + re.escape(old_version) + r')(.*?</h1>)'
        updated_content = re.sub(pattern, r'\1' + new_version + r'\3', content)
        
        # Only write if content changed
        if updated_content != content:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(updated_content)
            return True
        return False
    except Exception as e:
        print(f"Error updating {file_path}: {e}")
        return False


def find_html_files():
    """Find all HTML files that might contain version information"""
    html_files = []
    
    # Root directory files
    for file in Path('.').glob('*.html'):
        html_files.append(file)
    
    # Language directories and subdirectories
    lang_dirs = ['ar', 'de', 'es', 'ja', 'pl', 'pt', 'ru', 'zh', 'compare', 'how_use']
    
    for lang_dir in lang_dirs:
        if os.path.isdir(lang_dir):
            for file in Path(lang_dir).rglob('*.html'):
                html_files.append(file)
    
    return html_files


def get_current_version_from_files():
    """Extract current version from HTML files"""
    html_files = find_html_files()
    
    for file_path in html_files:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Look for version pattern in h1 tags
            match = re.search(r'<h1>[^<]*?(v \d+\.\d+[^<]*?)</h1>', content)
            if match:
                return match.group(1).strip()
        except Exception:
            continue
    
    return None


def update_all_files(new_version):
    """Update version in all HTML files"""
    current_version = get_current_version_from_files()
    
    if not current_version:
        print("Could not detect current version in HTML files")
        return False
    
    print(f"Current version: {current_version}")
    print(f"New version: {new_version}")
    
    if current_version == new_version:
        print("Version is already up to date!")
        return True
    
    html_files = find_html_files()
    updated_files = []
    
    for file_path in html_files:
        if update_html_file(file_path, current_version, new_version):
            updated_files.append(str(file_path))
    
    if updated_files:
        print(f"\nUpdated {len(updated_files)} files:")
        for file in updated_files:
            print(f"  - {file}")
        return True
    else:
        print("No files were updated")
        return False


def update_version_config(new_version):
    """Update version.json with new version"""
    config = load_version_config()
    if not config:
        return False
    
    old_version = config.get('version', '')
    config['version'] = new_version
    config['releaseDate'] = '2025-08-08'  # You can modify this as needed
    
    try:
        with open('version.json', 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        print(f"Updated version.json: {old_version} -> {new_version}")
        return True
    except Exception as e:
        print(f"Error updating version.json: {e}")
        return False


def main():
    """Main function"""
    print("MaxPilot Version Update Script")
    print("=" * 40)
    
    # Load current config
    config = load_version_config()
    if not config:
        return
    
    current_version = config.get('version', 'Unknown')
    print(f"Current version in config: {current_version}")
    
    # Get new version from user
    new_version = input(f"Enter new version (current: {current_version}): ").strip()
    
    if not new_version:
        print("No version entered. Exiting.")
        return
    
    if new_version == current_version:
        print("Same version entered. No changes needed.")
        return
    
    # Confirm update
    confirm = input(f"Update from '{current_version}' to '{new_version}'? (y/N): ").strip().lower()
    if confirm not in ['y', 'yes']:
        print("Update cancelled.")
        return
    
    # Update files
    print("\nUpdating HTML files...")
    if update_all_files(new_version):
        print("\nUpdating version.json...")
        update_version_config(new_version)
        print("\n✅ Version update completed successfully!")
    else:
        print("\n❌ Version update failed!")


if __name__ == "__main__":
    main()
