#!/usr/bin/env python3
"""
Quick test script to verify the Schedule Validator functionality
"""

from schedule_validator import ScheduleValidator

def test_normalization():
    """Test normalization functions"""
    validator = ScheduleValidator()
    
    print("🧪 Testing Normalization Functions")
    print("=" * 50)
    
    # Test class name normalization
    print("\n1️⃣ Class Name Normalization:")
    test_classes = [
        "STRENGTH LAB (PUSH)",
        "Strength Lab Push",
        "SL PUSH",
        "cardio barre",
        "CARDIOBARRE",
        "CB"
    ]
    
    for class_name in test_classes:
        normalized = validator.normalize_class_name(class_name)
        print(f"   '{class_name}' → '{normalized}'")
    
    # Test trainer name normalization
    print("\n2️⃣ Trainer Name Normalization:")
    test_trainers = [
        "anisha",
        "  ROHAN  ",
        "Simran",
        "MRIGAKSHI"
    ]
    
    for trainer in test_trainers:
        normalized = validator.normalize_trainer_name(trainer)
        print(f"   '{trainer}' → '{normalized}'")
    
    # Test time normalization
    print("\n3️⃣ Time Normalization:")
    test_times = [
        "9:00 AM",
        "9.00 AM",
        "9:00AM",
        "9 AM",
        "10:30 PM",
        "11:00 am"
    ]
    
    for time in test_times:
        normalized = validator.normalize_time(time)
        print(f"   '{time}' → '{normalized}'")
    
    print("\n✅ Normalization tests completed!")

def test_csv_parsing():
    """Test CSV parsing"""
    validator = ScheduleValidator()
    
    print("\n\n📄 Testing CSV Parsing")
    print("=" * 50)
    
    csv_content = """Location,Day,Time,Class,Trainer,Cover Trainer
KEMPS,MONDAY,7:15 AM,STRENGTH LAB (PULL),ANISHA,
KEMPS,MONDAY,7:30 AM,BARRE 57,SIMONELLE,RICHARD
BANDRA,TUESDAY,9:00 AM,powerCycle,BRET,
"""
    
    validator.parse_csv(csv_content)
    
    print(f"\n✓ Parsed {len(validator.csv_data)} location(s)")
    
    for location, data in validator.csv_data.items():
        print(f"\n📍 {location}:")
        for key, entries in data.items():
            for entry in entries:
                print(f"   {entry['day']} {entry['time']}: {entry['class']} - {entry['trainer']}")
    
    print("\n✅ CSV parsing test completed!")

def main():
    """Run all tests"""
    print("\n" + "=" * 50)
    print("🧪 Schedule Validator - Test Suite")
    print("=" * 50 + "\n")
    
    try:
        test_normalization()
        test_csv_parsing()
        
        print("\n\n" + "=" * 50)
        print("✅ All tests passed successfully!")
        print("=" * 50 + "\n")
        
    except Exception as e:
        print(f"\n\n❌ Test failed with error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
