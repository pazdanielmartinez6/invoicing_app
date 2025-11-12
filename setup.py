#!/usr/bin/env python3
"""
Setup script for Invoice Generator
Creates necessary directories and validates setup
"""

import os
import sys
from pathlib import Path
import json


def create_directory_structure():
    """Create necessary directories if they don't exist"""
    base_dir = Path(__file__).parent
    
    directories = [
        base_dir / "templates",
        base_dir / "output",
        base_dir / "output" / "one_pager",
        base_dir / "output" / "back_up",
    ]
    
    print("Creating directory structure...")
    for directory in directories:
        directory.mkdir(parents=True, exist_ok=True)
        print(f"  ✓ {directory.relative_to(base_dir)}")
    
    print("\nDirectory structure created successfully!")


def check_dependencies():
    """Check if required packages are installed"""
    print("\nChecking dependencies...")
    
    required_packages = {
        'fitz': 'PyMuPDF',
        'pandas': 'pandas',
        'PIL': 'Pillow',
        'openpyxl': 'openpyxl'
    }
    
    missing_packages = []
    
    for import_name, package_name in required_packages.items():
        try:
            __import__(import_name)
            print(f"  ✓ {package_name}")
        except ImportError:
            print(f"  ✗ {package_name} - NOT INSTALLED")
            missing_packages.append(package_name)
    
    if missing_packages:
        print("\n❌ Missing packages detected!")
        print("Install them using:")
        print(f"  pip install {' '.join(missing_packages)}")
        print("\nOr install all requirements:")
        print("  pip install -r requirements.txt")
        return False
    else:
        print("\n✓ All dependencies installed!")
        return True


def check_config():
    """Check if config file exists and is valid"""
    print("\nChecking configuration...")
    
    config_path = Path(__file__).parent / "config.json"
    
    if not config_path.exists():
        print("  ✗ config.json not found!")
        return False
    
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        # Check required keys
        required_keys = ['paths', 'text_positions', 'pdf_settings']
        for key in required_keys:
            if key not in config:
                print(f"  ✗ Missing '{key}' in config.json")
                return False
        
        print("  ✓ config.json is valid")
        return True
    except json.JSONDecodeError:
        print("  ✗ config.json has invalid JSON format")
        return False


def check_templates():
    """Check if template files exist"""
    print("\nChecking template files...")
    
    base_dir = Path(__file__).parent
    templates_dir = base_dir / "templates"
    
    required_templates = [
        "front_pager.pdf",
        "blank_template.pdf",
        "applogo.png"
    ]
    
    all_present = True
    for template in required_templates:
        template_path = templates_dir / template
        if template_path.exists():
            print(f"  ✓ {template}")
        else:
            print(f"  ✗ {template} - NOT FOUND")
            all_present = False
    
    if not all_present:
        print("\n⚠️  Some template files are missing!")
        print("Please add your template files to the 'templates/' folder:")
        print(f"  - front_pager.pdf: Invoice front page template")
        print(f"  - blank_template.pdf: Backup page template")
        print(f"  - applogo.png: Application logo (optional)")
        return False
    else:
        print("\n✓ All template files present!")
        return True


def main():
    """Main setup function"""
    print("=" * 60)
    print("Invoice Generator - Setup Script")
    print("=" * 60)
    
    # Create directories
    create_directory_structure()
    
    # Check dependencies
    deps_ok = check_dependencies()
    
    # Check configuration
    config_ok = check_config()
    
    # Check templates
    templates_ok = check_templates()
    
    # Summary
    print("\n" + "=" * 60)
    print("Setup Summary")
    print("=" * 60)
    
    if deps_ok and config_ok:
        if templates_ok:
            print("✓ Setup complete! You're ready to run the application.")
            print("\nTo start the application, run:")
            print("  python invoice_generator.py")
        else:
            print("⚠️  Setup incomplete - template files missing")
            print("\nAdd your template files, then run:")
            print("  python invoice_generator.py")
    else:
        print("❌ Setup incomplete - please resolve the issues above")
        if not deps_ok:
            print("\n1. Install dependencies:")
            print("   pip install -r requirements.txt")
        if not config_ok:
            print("\n2. Ensure config.json is present and valid")
    
    print("=" * 60)


if __name__ == "__main__":
    main()