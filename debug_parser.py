#!/usr/bin/env python3
"""
Debug test for HTML parsing
"""

from schedule_validator import ScheduleValidator

def test_html_parsing():
    """Test HTML parsing with actual HTML content"""
    validator = ScheduleValidator()
    
    # Read actual Bandra.html file
    try:
        with open('Bandra.html', 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        print("📄 Testing HTML Parsing")
        print("=" * 60)
        print(f"File size: {len(html_content)} bytes")
        
        # Parse the HTML
        validator.parse_html_schedule(html_content, 'BANDRA')
        
        print(f"\n✓ Parsed location: BANDRA")
        
        if 'BANDRA' in validator.pdf_data:
            data = validator.pdf_data['BANDRA']
            print(f"✓ Found {len(data)} time slots")
            
            # Count total classes
            total = sum(len(entries) for entries in data.values())
            print(f"✓ Total classes: {total}")
            
            # Show first few entries
            print("\n📋 Sample entries:")
            count = 0
            for key, entries in sorted(data.items())[:5]:
                location, day, time = key.split('|')
                for entry in entries:
                    print(f"   {day} {time}: {entry['class']} - {entry['trainer']}")
                    count += 1
                    if count >= 5:
                        break
                if count >= 5:
                    break
        else:
            print("❌ No data found for BANDRA")
            
    except FileNotFoundError:
        print("❌ Bandra.html not found")
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

def test_kemps_html():
    """Test Kemps HTML parsing"""
    validator = ScheduleValidator()
    
    try:
        with open('Kemps.html', 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        print("\n\n📄 Testing Kemps HTML Parsing")
        print("=" * 60)
        print(f"File size: {len(html_content)} bytes")
        
        validator.parse_html_schedule(html_content, 'KEMPS')
        
        if 'KEMPS' in validator.pdf_data:
            data = validator.pdf_data['KEMPS']
            total = sum(len(entries) for entries in data.values())
            print(f"✓ Found {len(data)} time slots")
            print(f"✓ Total classes: {total}")
            
            # Show first few entries
            print("\n📋 Sample entries:")
            count = 0
            for key, entries in sorted(data.items())[:5]:
                location, day, time = key.split('|')
                for entry in entries:
                    print(f"   {day} {time}: {entry['class']} - {entry['trainer']}")
                    count += 1
                    if count >= 5:
                        break
                if count >= 5:
                    break
        else:
            print("❌ No data found for KEMPS")
            
    except FileNotFoundError:
        print("❌ Kemps.html not found")
    except Exception as e:
        print(f"❌ Error: {e}")

def test_csv_parsing():
    """Test CSV parsing"""
    validator = ScheduleValidator()
    
    try:
        with open('sample_schedule.csv', 'r', encoding='utf-8') as f:
            csv_content = f.read()
        
        print("\n\n📄 Testing CSV Parsing")
        print("=" * 60)
        
        validator.parse_csv(csv_content)
        
        for location, data in validator.csv_data.items():
            total = sum(len(entries) for entries in data.values())
            print(f"✓ {location}: {total} classes")
            
    except FileNotFoundError:
        print("❌ sample_schedule.csv not found")
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == '__main__':
    print("\n" + "=" * 60)
    print("🧪 Schedule Validator - HTML Parsing Debug Test")
    print("=" * 60 + "\n")
    
    test_html_parsing()
    test_kemps_html()
    test_csv_parsing()
    
    print("\n" + "=" * 60)
    print("✅ Debug test completed!")
    print("=" * 60 + "\n")
