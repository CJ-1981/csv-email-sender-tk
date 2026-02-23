"""
Simple test script to verify the application setup
"""
import sys

def test_imports():
    """Test that all modules can be imported"""
    print("Testing imports...")
    try:
        import config
        print("  ✓ config")
        import csv_parser
        print("  ✓ csv_parser")
        import email_sender
        print("  ✓ email_sender")
        import tkinter
        print("  ✓ tkinter (built-in)")
        return True
    except ImportError as e:
        print(f"  ✗ Import error: {e}")
        return False

def test_openpyxl():
    """Test that openpyxl is available"""
    print("\nTesting openpyxl...")
    try:
        import openpyxl
        print(f"  ✓ openpyxl version {openpyxl.__version__}")
        return True
    except ImportError:
        print("  ✗ openpyxl not installed. Run: pip install openpyxl")
        return False

def test_csv_parser():
    """Test CSV parser functionality"""
    print("\nTesting CSV parser...")
    from csv_parser import CSVParser, generate_template_csv
    
    # Test template generation
    import tempfile
    import os
    
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
        template_path = f.name
    
    try:
        if generate_template_csv(template_path):
            print("  ✓ Template CSV generation")
            
            # Test parsing
            parser = CSVParser()
            if parser.parse_file(template_path):
                data = parser.get_data()
                print(f"  ✓ CSV parsing (loaded {len(data)} rows)")
                return True
            else:
                print(f"  ✗ CSV parsing failed: {parser.get_error()}")
                return False
    finally:
        if os.path.exists(template_path):
            os.remove(template_path)

def main():
    """Run all tests"""
    print("=" * 50)
    print("CSV Email Sender - Setup Verification")
    print("=" * 50)
    
    results = []
    results.append(test_imports())
    results.append(test_openpyxl())
    results.append(test_csv_parser())
    
    print("\n" + "=" * 50)
    if all(results):
        print("✓ All tests passed! Application is ready to run.")
        print("\nTo start the application, run:")
        print("  python main.py")
    else:
        print("✗ Some tests failed. Please fix the issues above.")
    print("=" * 50)

if __name__ == "__main__":
    main()
