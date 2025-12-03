#!/usr/bin/env python3
"""
Schedule Validator - Compare CSV schedule data with PDF files
Identifies discrepancies between uploaded schedules
"""

import csv
import re
import json
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import parse_qs
from io import StringIO, BytesIO
import PyPDF2
from datetime import datetime

class ScheduleValidator:
    """Handles schedule parsing, normalization, and comparison"""
    
    # Normalization mappings
    CLASS_ALIASES = {
        'STRENGTH LAB (PUSH)': ['STRENGTH LAB PUSH', 'SL PUSH', 'STRENGTH PUSH'],
        'STRENGTH LAB (PULL)': ['STRENGTH LAB PULL', 'SL PULL', 'STRENGTH PULL'],
        'STRENGTH LAB (FULL BODY)': ['STRENGTH LAB FULL BODY', 'SL FULL BODY', 'STRENGTH FULL BODY'],
        'CARDIO BARRE': ['CARDIOBARRE', 'CB'],
        'CARDIO BARRE PLUS': ['CARDIO BARRE+', 'CB PLUS', 'CB+'],
        'CARDIO BARRE EXPRESS': ['CARDIO BARRE EXP', 'CB EXPRESS', 'CB EXP'],
        'BARRE 57': ['BARRE57', 'B57'],
        'MAT 57': ['MAT57', 'M57'],
        'MAT 57 EXPRESS': ['MAT 57 EXP', 'MAT57 EXPRESS', 'M57 EXP'],
        'POWERCYCLE': ['POWER CYCLE', 'PC'],
        'AMPED UP!': ['AMPED UP', 'AMPED'],
        'FIT': ['FITNESS'],
        'FOUNDATIONS': ['FOUNDATION'],
        'SWEAT IN 30': ['SWEAT IN 30 MIN', 'SWEAT'],
    }
    
    DAYS = ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY']
    
    def __init__(self):
        self.csv_data = {}
        self.pdf_data = {}
        
    def normalize_class_name(self, class_name):
        """Normalize class names to standard format"""
        if not class_name:
            return ''
        
        # Clean and uppercase
        normalized = class_name.strip().upper()
        
        # Remove extra spaces
        normalized = re.sub(r'\s+', ' ', normalized)
        
        # Check aliases
        for standard, aliases in self.CLASS_ALIASES.items():
            if normalized == standard:
                return standard
            if normalized in aliases:
                return standard
                
        return normalized
    
    def normalize_trainer_name(self, trainer_name):
        """Normalize trainer names to standard format"""
        if not trainer_name:
            return ''
        
        # Clean and uppercase
        normalized = trainer_name.strip().upper()
        
        # Remove extra spaces
        normalized = re.sub(r'\s+', ' ', normalized)
        
        return normalized
    
    def normalize_time(self, time_str):
        """Normalize time format to HH:MM AM/PM"""
        if not time_str:
            return ''
        
        # Clean the time string
        time_str = time_str.strip().upper()
        
        # Remove extra spaces
        time_str = re.sub(r'\s+', ' ', time_str)
        
        # Try to parse different time formats
        patterns = [
            r'(\d{1,2}):(\d{2})\s*(AM|PM)',  # 9:00 AM
            r'(\d{1,2})\.(\d{2})\s*(AM|PM)',  # 9.00 AM
            r'(\d{1,2}):(\d{2})(AM|PM)',      # 9:00AM
            r'(\d{1,2})\s*(AM|PM)',           # 9 AM
        ]
        
        for pattern in patterns:
            match = re.search(pattern, time_str)
            if match:
                if len(match.groups()) == 3:
                    hour, minute, period = match.groups()
                    return f"{int(hour)}:{minute} {period}"
                elif len(match.groups()) == 2:
                    hour, period = match.groups()
                    return f"{int(hour)}:00 {period}"
        
        return time_str
    
    def parse_csv(self, csv_content):
        """Parse CSV schedule data"""
        reader = csv.DictReader(StringIO(csv_content))
        
        for row in reader:
            location = row.get('Location', '').strip().upper()
            day = row.get('Day', '').strip().upper()
            time = self.normalize_time(row.get('Time', ''))
            class_name = self.normalize_class_name(row.get('Class', ''))
            trainer = self.normalize_trainer_name(row.get('Trainer', ''))
            
            # Handle cover replacements if specified
            if 'Cover Trainer' in row and row['Cover Trainer'].strip():
                trainer = self.normalize_trainer_name(row['Cover Trainer'])
            
            # Skip empty rows
            if not (location and day and time and class_name and trainer):
                continue
            
            # Create unique key
            key = f"{location}|{day}|{time}"
            
            if location not in self.csv_data:
                self.csv_data[location] = {}
            
            if key not in self.csv_data[location]:
                self.csv_data[location][key] = []
            
            self.csv_data[location][key].append({
                'day': day,
                'time': time,
                'class': class_name,
                'trainer': trainer
            })
    
    def extract_text_from_pdf(self, pdf_file):
        """Extract text from PDF file"""
        try:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text = ''
            
            for page in pdf_reader.pages:
                text += page.extract_text()
            
            return text
        except Exception as e:
            print(f"Error extracting PDF text: {e}")
            return ''
    
    def parse_html_schedule(self, html_content, location):
        """Parse schedule information from HTML content"""
        if location not in self.pdf_data:
            self.pdf_data[location] = {}
        
        # Split by lines
        lines = html_content.split('\n')
        
        current_day = None
        
        for i, line in enumerate(lines):
            line = line.strip()
            
            # Check if line contains a day header
            for day in self.DAYS:
                if f'>{day} </span>' in line or f'>{day}</span>' in line:
                    current_day = day
                    break
            
            if not current_day:
                continue
            
            # Look for time entries - pattern: <span class="t v0 s23" or s9 with time
            if ('class="t v0 s23"' in line or 'class="t v0 s9"' in line) and ('AM</span>' in line or 'PM</span>' in line):
                # Extract time
                time_match = re.search(r'>(\d{1,2}:\d{2}\s*(?:AM|PM))</span>', line)
                if not time_match:
                    continue
                
                time = self.normalize_time(time_match.group(1))
                
                # Look for the class and trainer in the next span (same line or next line)
                class_trainer_text = ''
                
                # Check current line for class info
                if i < len(lines) - 1:
                    # Often the class info is on the same line
                    remaining_line = line[line.find(time_match.group(0)):]
                    class_match = re.search(r'<span[^>]*>([^<]+)</span>', remaining_line)
                    if class_match:
                        class_trainer_text = class_match.group(1)
                
                if not class_trainer_text:
                    continue
                
                # Parse class and trainer from "CLASS - TRAINER" format
                if ' - ' in class_trainer_text:
                    parts = class_trainer_text.split(' - ', 1)
                    class_name = self.normalize_class_name(parts[0])
                    trainer = self.normalize_trainer_name(parts[1])
                    
                    # Remove theme badges and other decorations
                    class_name = re.sub(r'\s*<.*?>', '', class_name).strip()
                    trainer = re.sub(r'\s*<.*?>', '', trainer).strip()
                    
                    if class_name and trainer:
                        key = f"{location}|{current_day}|{time}"
                        
                        if key not in self.pdf_data[location]:
                            self.pdf_data[location][key] = []
                        
                        self.pdf_data[location][key].append({
                            'day': current_day,
                            'time': time,
                            'class': class_name,
                            'trainer': trainer
                        })
    
    def parse_pdf_schedule(self, pdf_content, location):
        """Parse schedule - handles both PDF text and HTML content"""
        # Check if it's HTML content
        if '<html' in pdf_content.lower() or '<span' in pdf_content.lower():
            self.parse_html_schedule(pdf_content, location)
        else:
            # Original PDF text parsing
            if location not in self.pdf_data:
                self.pdf_data[location] = {}
            
            lines = pdf_content.split('\n')
            current_day = None
            
            for line in lines:
                line = line.strip()
                
                # Check if line is a day
                day_match = any(day in line.upper() for day in self.DAYS)
                if day_match:
                    for day in self.DAYS:
                        if day in line.upper():
                            current_day = day
                            break
                    continue
                
                if not current_day:
                    continue
                
                # Try to extract time, class, and trainer
                pattern = r'(\d{1,2}:\d{2}\s*(?:AM|PM))\s+(.+?)\s*-\s*(.+?)(?:\n|$)'
                match = re.search(pattern, line, re.IGNORECASE)
                
                if match:
                    time = self.normalize_time(match.group(1))
                    class_name = self.normalize_class_name(match.group(2))
                    trainer = self.normalize_trainer_name(match.group(3))
                    
                    # Remove theme badges and other decorations
                    class_name = re.sub(r'\(.*?\)', '', class_name).strip()
                    trainer = re.sub(r'\(.*?\)', '', trainer).strip()
                    
                    key = f"{location}|{current_day}|{time}"
                    
                    if key not in self.pdf_data[location]:
                        self.pdf_data[location][key] = []
                    
                    self.pdf_data[location][key].append({
                        'day': current_day,
                        'time': time,
                        'class': class_name,
                        'trainer': trainer
                    })
    
    def compare_schedules(self):
        """Compare CSV and PDF data to find discrepancies"""
        all_comparisons = []
        discrepancies = []
        matched = 0
        total = 0
        
        # Check all locations
        all_locations = set(list(self.csv_data.keys()) + list(self.pdf_data.keys()))
        
        for location in all_locations:
            csv_loc = self.csv_data.get(location, {})
            pdf_loc = self.pdf_data.get(location, {})
            
            # Get all unique keys (day|time combinations)
            all_keys = set(list(csv_loc.keys()) + list(pdf_loc.keys()))
            
            for key in sorted(all_keys):
                location_name, day, time = key.split('|')
                
                csv_entries = csv_loc.get(key, [])
                pdf_entries = pdf_loc.get(key, [])
                
                # Compare entries
                for csv_entry in csv_entries:
                    total += 1
                    
                    # Find matching PDF entry
                    match_found = False
                    matched_pdf_entry = None
                    
                    for pdf_entry in pdf_entries:
                        class_match = csv_entry['class'] == pdf_entry['class']
                        trainer_match = csv_entry['trainer'] == pdf_entry['trainer']
                        
                        if class_match and trainer_match:
                            matched += 1
                            match_found = True
                            matched_pdf_entry = pdf_entry
                            break
                    
                    # Find closest match for comparison if no exact match
                    if not matched_pdf_entry:
                        matched_pdf_entry = pdf_entries[0] if pdf_entries else {
                            'class': 'NOT FOUND',
                            'trainer': 'NOT FOUND'
                        }
                    
                    comparison = {
                        'location': location_name,
                        'day': day,
                        'time': time,
                        'csvClass': csv_entry['class'],
                        'csvTrainer': csv_entry['trainer'],
                        'pdfClass': matched_pdf_entry['class'],
                        'pdfTrainer': matched_pdf_entry['trainer'],
                        'classMatch': csv_entry['class'] == matched_pdf_entry['class'],
                        'trainerMatch': csv_entry['trainer'] == matched_pdf_entry['trainer'],
                        'isMatch': match_found
                    }
                    
                    all_comparisons.append(comparison)
                    
                    if not match_found:
                        discrepancies.append(comparison)
                
                # Check for PDF entries not in CSV
                for pdf_entry in pdf_entries:
                    if not any(csv['class'] == pdf_entry['class'] and 
                             csv['trainer'] == pdf_entry['trainer'] 
                             for csv in csv_entries):
                        total += 1
                        comparison = {
                            'location': location_name,
                            'day': day,
                            'time': time,
                            'csvClass': 'NOT IN CSV',
                            'csvTrainer': 'NOT IN CSV',
                            'pdfClass': pdf_entry['class'],
                            'pdfTrainer': pdf_entry['trainer'],
                            'classMatch': False,
                            'trainerMatch': False,
                            'isMatch': False
                        }
                        all_comparisons.append(comparison)
                        discrepancies.append(comparison)
        
        return {
            'totalClasses': total,
            'matchedClasses': matched,
            'discrepancies': discrepancies,
            'allComparisons': all_comparisons
        }


class ValidationHandler(BaseHTTPRequestHandler):
    """HTTP request handler for the validation service"""
    
    def do_GET(self):
        """Serve the HTML interface"""
        if self.path == '/' or self.path == '/index.html':
            try:
                with open('schedule-validator.html', 'r') as f:
                    content = f.read()
                
                self.send_response(200)
                self.send_header('Content-type', 'text/html')
                self.end_headers()
                self.wfile.write(content.encode())
            except FileNotFoundError:
                self.send_error(404, 'File not found')
        else:
            self.send_error(404, 'File not found')
    
    def do_POST(self):
        """Handle file upload and validation"""
        if self.path == '/validate-schedules':
            try:
                # Parse multipart form data
                content_type = self.headers['Content-Type']
                if 'multipart/form-data' not in content_type:
                    self.send_error(400, 'Invalid content type')
                    return
                
                # Get boundary
                boundary = content_type.split('boundary=')[1].encode()
                
                # Read the entire request body
                content_length = int(self.headers['Content-Length'])
                body = self.rfile.read(content_length)
                
                # Parse multipart data
                files = self.parse_multipart(body, boundary)
                
                # Validate files
                if 'csv' not in files:
                    self.send_error(400, 'CSV file is required')
                    return
                
                if 'kemps' not in files and 'bandra' not in files:
                    self.send_error(400, 'At least one PDF file is required')
                    return
                
                # Process files
                validator = ScheduleValidator()
                
                # Parse CSV
                csv_content = files['csv'].decode('utf-8')
                validator.parse_csv(csv_content)
                print(f"[DEBUG] CSV parsed: {len(validator.csv_data)} locations")
                for loc in validator.csv_data:
                    count = sum(len(e) for e in validator.csv_data[loc].values())
                    print(f"[DEBUG]   {loc}: {count} classes")
                
                # Parse PDFs or HTML files
                if 'kemps' in files:
                    file_content = files['kemps']
                    print(f"[DEBUG] Kemps file size: {len(file_content)} bytes")
                    # Check if it's HTML or PDF
                    if file_content.startswith(b'%PDF'):
                        # It's a PDF
                        print(f"[DEBUG] Processing as PDF")
                        pdf_file = BytesIO(file_content)
                        pdf_text = validator.extract_text_from_pdf(pdf_file)
                        validator.parse_pdf_schedule(pdf_text, 'KEMPS')
                    else:
                        # It's likely HTML
                        print(f"[DEBUG] Processing as HTML")
                        html_content = file_content.decode('utf-8')
                        validator.parse_html_schedule(html_content, 'KEMPS')
                    
                    if 'KEMPS' in validator.pdf_data:
                        count = sum(len(e) for e in validator.pdf_data['KEMPS'].values())
                        print(f"[DEBUG] Kemps parsed: {count} classes")
                
                if 'bandra' in files:
                    file_content = files['bandra']
                    print(f"[DEBUG] Bandra file size: {len(file_content)} bytes")
                    # Check if it's HTML or PDF
                    if file_content.startswith(b'%PDF'):
                        # It's a PDF
                        print(f"[DEBUG] Processing as PDF")
                        pdf_file = BytesIO(file_content)
                        pdf_text = validator.extract_text_from_pdf(pdf_file)
                        validator.parse_pdf_schedule(pdf_text, 'BANDRA')
                    else:
                        # It's likely HTML
                        print(f"[DEBUG] Processing as HTML")
                        html_content = file_content.decode('utf-8')
                        validator.parse_html_schedule(html_content, 'BANDRA')
                    
                    if 'BANDRA' in validator.pdf_data:
                        count = sum(len(e) for e in validator.pdf_data['BANDRA'].values())
                        print(f"[DEBUG] Bandra parsed: {count} classes")
                
                # Compare schedules
                results = validator.compare_schedules()
                
                # Send response
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps(results).encode())
                
            except Exception as e:
                self.send_error(500, f'Server error: {str(e)}')
        else:
            self.send_error(404, 'Endpoint not found')
    
    def parse_multipart(self, body, boundary):
        """Parse multipart form data"""
        files = {}
        parts = body.split(b'--' + boundary)
        
        for part in parts:
            if b'Content-Disposition' in part:
                # Extract filename
                disposition = part.split(b'\r\n')[1].decode()
                filename_match = re.search(r'name="([^"]+)"', disposition)
                
                if filename_match:
                    field_name = filename_match.group(1)
                    
                    # Extract file content
                    content_start = part.find(b'\r\n\r\n') + 4
                    content_end = part.rfind(b'\r\n')
                    
                    if content_start > 3 and content_end > content_start:
                        files[field_name] = part[content_start:content_end]
        
        return files
    
    def log_message(self, format, *args):
        """Custom logging"""
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {format % args}")


def main():
    """Start the validation server"""
    port = 8080
    server_address = ('', port)
    httpd = HTTPServer(server_address, ValidationHandler)
    
    print(f"🚀 Schedule Validator Server Starting...")
    print(f"📍 Server running at: http://localhost:{port}")
    print(f"🌐 Open your browser and navigate to the URL above")
    print(f"⏹️  Press Ctrl+C to stop the server\n")
    
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n\n⏹️  Server stopped")
        httpd.shutdown()


if __name__ == '__main__':
    main()
