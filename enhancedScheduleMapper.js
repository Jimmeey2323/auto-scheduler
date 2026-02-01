/**
 * Enhanced Schedule and Theme Mapping System
 * Provides dynamic location identification, class normalization, and theme extraction
 */

class EnhancedScheduleMapper {
    constructor() {
        // Dynamic location identification patterns (order matters for specificity)
        this.locationPatterns = {
            'bandra': {
                aliases: ['supreme', 'bandra', 'shq', 'supremehq', 'bandra hq', 'supreme hq', 'bandra west'],
                canonical: 'Supreme HQ, Bandra',
                shortName: 'Bandra'
            },
            'kemps': {
                aliases: ['kemps', 'kh', 'kwality', 'kemps corner', 'kwality house', 'annex'],
                canonical: 'Kwality House, Kemps Corner', 
                shortName: 'Kemps'
            }
        };

        // Dynamic class name normalization patterns
        this.classPatterns = {
            'powercycle': {
                aliases: ['powercycle', 'cycle', 'spin', 'pc', 'power cycle'],
                canonical: 'Studio PowerCycle',
                displayFormat: 'powerCycle'
            },
            'backbodyblaze': {
                aliases: ['back body blaze', 'bbb', 'back-body-blaze'],
                canonical: 'Studio Back Body Blaze',
                displayFormat: 'STUDIO BACK BODY BLAZE'
            },
            'mat57': {
                aliases: ['mat 57', 'mat57', 'mat'],
                canonical: 'Studio Mat 57',
                displayFormat: 'STUDIO MAT 57'
            },
            'barre57': {
                aliases: ['barre 57', 'barre57', 'barre'],
                canonical: 'Studio Barre 57',
                displayFormat: 'STUDIO BARRE 57'
            },
            'ampedup': {
                aliases: ['amped up', 'amped', 'amped up!'],
                canonical: 'Studio Amped Up!',
                displayFormat: 'STUDIO AMPED UP!'
            },
            'fit': {
                aliases: ['fit', 'studio fit'],
                canonical: 'Studio FIT',
                displayFormat: 'STUDIO FIT'
            }
        };

        // Power Cycle theme templates from user requirements
        this.powerCycleTemplates = {
            'bandra': [
                { day: 'Monday', time: '8:45', theme: 'Lady Gaga vs Bruno Mars' },
                { day: 'Tuesday', time: '7:15', theme: 'Rihanna + Friends' },
                { day: 'Wednesday', time: '10:30', theme: 'Rihanna + Friends' },
                { day: 'Thursday', time: '8:00', theme: 'Lady Gaga vs Bruno Mars' },
                { day: 'Friday', time: '7:15', theme: 'Lady Gaga vs Bruno Mars' },
                { day: 'Saturday', time: '11:30', theme: 'Rihanna + Friends' }
            ],
            'kemps': [
                { day: 'Monday', time: '8:00', theme: 'Rihanna + Friends' },
                { day: 'Tuesday', time: '10:00', theme: 'Lady Gaga vs Bruno Mars' },
                { day: 'Wednesday', time: '6:30', theme: 'Lady Gaga vs Bruno Mars' },
                { day: 'Thursday', time: '8:00', theme: 'Lady Gaga vs Bruno Mars' },
                { day: 'Friday', time: '7:15', theme: 'Rihanna + Friends' },
                { day: 'Sunday', time: '10:00', theme: 'Rihanna + Friends' },
                // Annex PowerCycle classes
                { day: 'Monday', time: '8:00', theme: 'Lady Gaga vs Bruno Mars' },  // Annex
                { day: 'Monday', time: '7:00', theme: 'Rihanna + Friends' },       // Annex
                { day: 'Tuesday', time: '7:30', theme: 'Lady Gaga vs Bruno Mars' }, // Annex
                { day: 'Tuesday', time: '10:00', theme: 'Rihanna + Friends' },     // Annex
                { day: 'Tuesday', time: '7:15', theme: 'Lady Gaga vs Bruno Mars' }, // Annex
                { day: 'Wednesday', time: '8:00', theme: 'Rihanna + Friends' },    // Annex
                { day: 'Wednesday', time: '6:30', theme: 'Lady Gaga vs Bruno Mars' }, // Annex
                { day: 'Thursday', time: '8:00', theme: 'Lady Gaga vs Bruno Mars' }, // Annex
                { day: 'Thursday', time: '9:30', theme: 'Rihanna + Friends' },     // Annex
                { day: 'Thursday', time: '6:00', theme: 'Lady Gaga vs Bruno Mars' }, // Annex
                { day: 'Friday', time: '8:30', theme: 'Rihanna + Friends' },       // Annex
                { day: 'Friday', time: '7:15', theme: 'Lady Gaga vs Bruno Mars' }, // Annex
                { day: 'Saturday', time: '10:00', theme: 'Lady Gaga vs Bruno Mars' }, // Annex
                { day: 'Saturday', time: '11:30', theme: 'Rihanna + Friends' },    // Annex
                { day: 'Saturday', time: '4:30', theme: 'Lady Gaga vs Bruno Mars' }, // Annex
                { day: 'Sunday', time: '10:00', theme: 'Rihanna + Friends' },      // Annex
                { day: 'Sunday', time: '11:30', theme: 'Lady Gaga vs Bruno Mars' }, // Annex
                { day: 'Sunday', time: '5:00', theme: 'Rihanna + Friends' }        // Annex
            ]
        };
    }

    /**
     * Enhanced weekend date calculation with proper week boundary handling
     */
    calculateWeekendDates() {
        const today = new Date();
        const currentDay = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
        
        // Calculate Monday of current week (start of work week)
        const monday = new Date(today);
        const daysFromMonday = currentDay === 0 ? -6 : (1 - currentDay);
        monday.setDate(today.getDate() + daysFromMonday);
        
        // Calculate Saturday and Sunday of current week
        const saturday = new Date(monday);
        saturday.setDate(monday.getDate() + 5); // Monday + 5 = Saturday
        
        const sunday = new Date(monday);
        sunday.setDate(monday.getDate() + 6); // Monday + 6 = Sunday
        
        return {
            monday,
            saturday,
            sunday,
            weekStart: monday,
            weekEnd: sunday
        };
    }

    /**
     * Validate weekend dates against schedule data
     */
    validateWeekendDates(scheduleData) {
        const { saturday, sunday } = this.calculateWeekendDates();
        const errors = [];
        
        // Check if weekend classes have correct dates
        scheduleData.forEach((classItem, index) => {
            const day = classItem.Day?.toLowerCase();
            if (day === 'saturday' || day === 'sunday') {
                const expectedDate = day === 'saturday' ? saturday : sunday;
                const actualDate = new Date(classItem.Date);
                
                if (actualDate.toDateString() !== expectedDate.toDateString()) {
                    errors.push({
                        index,
                        day: classItem.Day,
                        expected: expectedDate.toDateString(),
                        actual: actualDate.toDateString(),
                        class: classItem
                    });
                }
            }
        });
        
        return {
            isValid: errors.length === 0,
            errors,
            correctDates: { saturday, sunday }
        };
    }

    /**
     * Dynamic location identification using fuzzy matching with priority ordering
     */
    identifyLocation(locationText) {
        if (!locationText) return null;
        
        const normalized = locationText.toLowerCase().trim();
        
        // Check for more specific patterns first (longer matches have priority)
        const matches = [];
        
        for (const [key, config] of Object.entries(this.locationPatterns)) {
            for (const alias of config.aliases) {
                if (normalized.includes(alias.toLowerCase()) || alias.toLowerCase().includes(normalized)) {
                    matches.push({
                        key,
                        config,
                        alias,
                        priority: alias.length, // Longer matches get higher priority
                        specificity: this.calculateLocationSpecificity(normalized, alias)
                    });
                }
            }
        }
        
        if (matches.length === 0) {
            return {
                key: 'unknown',
                canonical: locationText,
                shortName: locationText,
                matched: locationText
            };
        }
        
        // Sort by specificity first, then by priority (length)
        matches.sort((a, b) => {
            if (a.specificity !== b.specificity) {
                return b.specificity - a.specificity; // Higher specificity first
            }
            return b.priority - a.priority; // Longer matches first
        });
        
        const best = matches[0];
        return {
            key: best.key,
            canonical: best.config.canonical,
            shortName: best.config.shortName,
            matched: best.alias
        };
    }
    
    /**
     * Calculate location match specificity (higher = more specific/confident match)
     */
    calculateLocationSpecificity(locationText, alias) {
        const normalized = locationText.toLowerCase();
        const aliasLower = alias.toLowerCase();
        
        // Special handling for location disambiguation - context matters most
        if (normalized.includes('bandra')) {
            if (aliasLower.includes('bandra')) {
                // If the alias specifically mentions bandra, it's a very strong match
                const wordBoundaryRegex = new RegExp(`\\b${aliasLower.replace(' ', '\\s+')}\\b`);
                if (wordBoundaryRegex.test(normalized)) {
                    return 150; // Highest priority for exact Bandra matches
                }
                return 140; // Still high for partial bandra matches
            } else {
                // If "bandra" is in the text but alias doesn't mention it, penalize heavily
                return 5; // Very low priority for non-Bandra aliases when "bandra" is present
            }
        }
        
        if (normalized.includes('kemps') || normalized.includes('corner')) {
            if (aliasLower.includes('kemps') || aliasLower.includes('corner')) {
                const wordBoundaryRegex = new RegExp(`\\b${aliasLower.replace(' ', '\\s+')}\\b`);
                if (wordBoundaryRegex.test(normalized)) {
                    return 150; // Highest priority for exact Kemps matches
                }
                return 140; // Still high for partial kemps matches
            } else if (aliasLower === 'kwality' || aliasLower === 'kwality house') {
                return 70; // Medium priority for kwality when kemps indicators are present
            }
        }
        
        // Exact word match has high specificity
        const wordBoundaryRegex = new RegExp(`\\b${aliasLower.replace(' ', '\\s+')}\\b`);
        if (wordBoundaryRegex.test(normalized)) {
            return 100;
        }
        
        // Check for multi-word matches
        if (aliasLower.includes(' ')) {
            const aliasWords = aliasLower.split(' ');
            const matchedWords = aliasWords.filter(word => normalized.includes(word)).length;
            const matchPercentage = (matchedWords / aliasWords.length) * 100;
            if (matchPercentage >= 80) {
                return 80 + matchPercentage; // Higher score for better multi-word matches
            }
        }
        
        // Check if it's a partial match
        if (normalized.includes(aliasLower)) {
            return 50;
        }
        
        // Check if alias is contained in text
        if (aliasLower.includes(normalized)) {
            return 30;
        }
        
        return 10; // Default low specificity
    }

    /**
     * Intelligent class name normalization with fuzzy matching
     */
    normalizeClassName(className) {
        if (!className) return null;
        
        const normalized = className.toLowerCase().trim().replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ');
        
        for (const [key, config] of Object.entries(this.classPatterns)) {
            for (const alias of config.aliases) {
                const aliasNormalized = alias.toLowerCase().replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ');
                if (normalized.includes(aliasNormalized) || aliasNormalized.includes(normalized)) {
                    return {
                        key,
                        canonical: config.canonical,
                        displayFormat: config.displayFormat,
                        matched: alias
                    };
                }
            }
        }
        
        return {
            key: 'unknown',
            canonical: className,
            displayFormat: className.toUpperCase(),
            matched: className
        };
    }

    /**
     * Extract and parse Power Cycle themes from email thread
     */
    extractPowerCycleThemes(emailThread) {
        const themes = [];
        let foundPowerCycleSection = false;
        
        // If emailThread is a string (backward compatibility), convert to array
        const messages = Array.isArray(emailThread) ? emailThread : [{ content: emailThread }];
        
        console.log(`🚴 Scanning ${messages.length} email messages for Power Cycle themes...`);
        
        // Scan through all messages in the thread
        for (let i = 0; i < messages.length; i++) {
            const message = messages[i];
            const content = message.content || message.body || '';
            
            console.log(`📧 Scanning message ${i + 1}/${messages.length} for PowerCycle themes...`);
            
            // Look for Power Cycle theme sections in this message
            const powerCycleRegex = /POWER\s*CYCLE\s*THEMES?\s*:?\s*(.*?)(?=\n\s*(?:Thanks|Best|Regards|On\s+(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)|\n--|Covers|FIT|AMPED|\n\s*$))/is;
            const match = content.match(powerCycleRegex);
            
            if (match) {
                foundPowerCycleSection = true;
                const themeContent = match[1].trim();
                console.log(`🎯 Found Power Cycle themes in message ${i + 1}:`);
                console.log(`📋 Theme content: ${themeContent.substring(0, 300)}...`);
                
                // Parse location-based theme sections
                const locationSections = this.parseLocationBasedThemes(themeContent);
                
                // Convert to theme objects
                for (const [location, locationThemes] of Object.entries(locationSections)) {
                    console.log(`🏢 Processing ${locationThemes.length} themes for location: ${location}`);
                    for (const theme of locationThemes) {
                        themes.push({
                            location: location,
                            day: theme.day,
                            time: theme.time,
                            theme: theme.theme,
                            classType: 'PowerCycle',
                            source: `message-${i + 1}`
                        });
                    }
                }
            }
        }
        
        if (!foundPowerCycleSection) {
            console.log('⚠️ No Power Cycle themes section found in any email message');
            console.log('🔄 Using default Power Cycle themes as fallback');
            return this.getDefaultPowerCycleThemes();
        }
        
        console.log(`✅ Successfully extracted ${themes.length} Power Cycle themes from email thread`);
        return themes;
    }

    /**
     * Parse location-based theme sections from Power Cycle content
     */
    parseLocationBasedThemes(content) {
        const locationSections = {};
        const lines = content.split('\n').map(line => line.trim()).filter(line => line);
        
        let currentLocation = null;
        
        console.log(`🔍 Parsing ${lines.length} lines for location-based themes:`);
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            console.log(`📝 Line ${i + 1}: "${line}"`);
            
            // Check if line is a location header (standalone or with minimal content)
            const locationMatch = this.identifyLocation(line);
            const isLocationHeader = locationMatch && locationMatch.key !== 'unknown' && 
                                   !line.match(/\b(mon|tue|wed|thu|fri|sat|sun)/i); // Not a theme line
            
            if (isLocationHeader) {
                currentLocation = locationMatch.key;
                locationSections[currentLocation] = [];
                console.log(`🏢 Found location header: ${currentLocation} ("${line}")`);
                continue;
            }
            
            // Parse theme line if we have a current location
            if (currentLocation && line.match(/\b(mon|tue|wed|thu|fri|sat|sun)/i)) {
                const theme = this.parseThemeLine(line);
                if (theme) {
                    locationSections[currentLocation].push(theme);
                    console.log(`🎵 Added theme for ${currentLocation}: ${theme.day} ${theme.time} - ${theme.theme}`);
                } else {
                    console.log(`⚠️ Failed to parse theme line: "${line}"`);
                }
            } else if (line.match(/\b(mon|tue|wed|thu|fri|sat|sun)/i)) {
                console.log(`⚠️ Found theme line without location context: "${line}"`);
                // Try to infer location from nearby lines or set as default
                if (!currentLocation) {
                    // Look ahead for location context
                    for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
                        const nextLocationMatch = this.identifyLocation(lines[j]);
                        if (nextLocationMatch && nextLocationMatch.key !== 'unknown') {
                            currentLocation = nextLocationMatch.key;
                            locationSections[currentLocation] = locationSections[currentLocation] || [];
                            console.log(`🔍 Inferred location from context: ${currentLocation}`);
                            break;
                        }
                    }
                }
                
                if (currentLocation) {
                    const theme = this.parseThemeLine(line);
                    if (theme) {
                        locationSections[currentLocation].push(theme);
                        console.log(`🎵 Added contextual theme for ${currentLocation}: ${theme.day} ${theme.time} - ${theme.theme}`);
                    }
                }
            }
        }
        
        console.log(`📊 Parsed themes by location:`);
        for (const [location, themes] of Object.entries(locationSections)) {
            console.log(`  ${location}: ${themes.length} themes`);
        }
        
        return locationSections;
    }

    /**
     * Parse individual theme line (e.g., "Mon 8:45 am – Lady Gaga vs Bruno Mars")
     */
    parseThemeLine(line) {
        // Enhanced patterns to match various formats:
        // "Mon 8:45 am - Lady gaga Vs Bruno Mars"
        // "Monday 8am - Rihanna + Friends"
        // "Tue 7:15 pm - Rihanna + Friends"
        // "Wed 6:30 pm - Lady Gaga Vs Bruno Mars"
        const patterns = [
            // Standard format: Day Time AM/PM - Theme (capture AM/PM)
            /\b(mon|tue|wed|thu|fri|sat|sun)[a-z]*\s*(\d{1,2}:?\d{0,2})\s*([ap]m?)\s*[–\-—]\s*(.+)/i,
            // Format with explicit "am/pm" spacing
            /\b(mon|tue|wed|thu|fri|sat|sun)[a-z]*\s*(\d{1,2})\s*([ap]m)\s*[–\-—]\s*(.+)/i,
            // Alternative format without AM/PM (default to AM)
            /\b(mon|tue|wed|thu|fri|sat|sun)[a-z]*\s*(\d{1,2}:?\d{0,2})\s*[–\-—]\s*(.+)/i
        ];
        
        let match = null;
        let patternUsed = -1;
        
        for (let i = 0; i < patterns.length; i++) {
            match = line.match(patterns[i]);
            if (match) {
                patternUsed = i;
                break;
            }
        }
        
        if (!match) {
            console.log(`❌ Failed to parse theme line with any pattern: "${line}"`);
            return null;
        }
        
        console.log(`✅ Parsed theme line using pattern ${patternUsed + 1}: "${line}"`);
        
        const dayMap = {
            'mon': 'Monday', 'tue': 'Tuesday', 'wed': 'Wednesday',
            'thu': 'Thursday', 'fri': 'Friday', 'sat': 'Saturday', 'sun': 'Sunday'
        };
        
        const day = dayMap[match[1].toLowerCase().substring(0, 3)];
        let time = match[2];
        const ampm = match[3] || ''; // AM/PM if captured
        const theme = match[patternUsed < 2 ? 4 : 3].trim(); // Theme is in different position depending on pattern
        
        // Combine time with AM/PM if present
        if (ampm) {
            time = `${time} ${ampm.toUpperCase()}`;
        }
        
        // Normalize time format (this will handle AM/PM conversion)
        time = this.normalizeTime(time);
        
        console.log(`🎵 Parsed: ${day} ${time} - ${theme}`);
        
        return { day, time, theme };
    }

    /**
     * Normalize time format to consistent HH:MM AM/PM format
     */
    normalizeTime(timeStr) {
        if (!timeStr) return timeStr;
        
        let time = timeStr.toString().trim();
        
        // Check if it has AM/PM indicator
        const hasAmPm = /\b([ap])\.?m\.?\b/i.test(time);
        const isPm = /\bp\.?m\.?\b/i.test(time);
        
        // Extract the numeric time part
        let normalized = time.replace(/[^\d:.]/g, '');
        
        // Handle formats like "8.45" -> "8:45"
        if (normalized.includes('.')) {
            normalized = normalized.replace('.', ':');
        }
        
        // Handle formats like "845" -> "8:45"
        if (!normalized.includes(':') && (normalized.length === 3 || normalized.length === 4)) {
            const hours = normalized.substring(0, normalized.length - 2);
            const minutes = normalized.substring(normalized.length - 2);
            normalized = `${hours}:${minutes}`;
        }
        
        // Ensure HH:MM format and convert back to 12-hour if needed
        if (normalized.includes(':')) {
            const parts = normalized.split(':');
            if (parts.length === 2 && parts[1].length <= 2) {
                let hours = parseInt(parts[0]);
                const minutes = parts[1].padEnd(2, '0');
                let pmFlag = isPm;
                
                // If we detected AM/PM in the original string, use that
                if (hasAmPm) {
                    // Time was already in 12-hour format
                    const hourStr = hours.toString().padStart(2, '0');
                    return `${hourStr}:${minutes} ${pmFlag ? 'PM' : 'AM'}`;
                } else {
                    // Time was in 24-hour format - need to convert
                    let period = 'AM';
                    if (hours >= 12) {
                        period = 'PM';
                        if (hours > 12) {
                            hours -= 12;
                        }
                    } else if (hours === 0) {
                        hours = 12;
                    }
                    
                    const hourStr = hours.toString().padStart(2, '0');
                    return `${hourStr}:${minutes} ${period}`;
                }
            }
        }
        
        return timeStr;
    }

    /**
     * Get default Power Cycle themes from user requirements
     */
    getDefaultPowerCycleThemes() {
        const themes = [];
        
        for (const [location, locationThemes] of Object.entries(this.powerCycleTemplates)) {
            for (const template of locationThemes) {
                themes.push({
                    location,
                    day: template.day,
                    time: template.time,
                    theme: template.theme,
                    classType: 'PowerCycle'
                });
            }
        }
        
        return themes;
    }

    /**
     * Map themes to specific classes based on location, day, time, and class type
     */
    mapThemesToClasses(themes, scheduleData) {
        const mappedClasses = [...scheduleData];
        const unmappedThemes = [];
        
        for (const theme of themes) {
            let mapped = false;
            
            for (const classItem of mappedClasses) {
                if (this.isThemeMatch(theme, classItem)) {
                    classItem.theme = theme.theme;
                    mapped = true;
                    console.log(`✅ Mapped theme "${theme.theme}" to ${classItem.Day} ${classItem.Time} ${classItem.Class} at ${classItem.Location}`);
                    break;
                }
            }
            
            if (!mapped) {
                unmappedThemes.push(theme);
                console.log(`⚠️ Failed to map theme "${theme.theme}" for ${theme.day} ${theme.time} ${theme.classType} at ${theme.location}`);
            }
        }
        
        return {
            mappedClasses,
            unmappedThemes,
            mappingStats: {
                totalThemes: themes.length,
                mappedCount: themes.length - unmappedThemes.length,
                unmappedCount: unmappedThemes.length
            }
        };
    }

    /**
     * Check if a theme matches a specific class
     */
    isThemeMatch(theme, classItem) {
        // Match day
        if (theme.day.toLowerCase() !== classItem.Day?.toLowerCase()) {
            return false;
        }
        
        // Match location
        const themeLocation = this.identifyLocation(theme.location);
        const classLocation = this.identifyLocation(classItem.Location);
        if (themeLocation.key !== classLocation.key) {
            return false;
        }
        
        // Match class type
        const normalizedClass = this.normalizeClassName(classItem.Class);
        if (theme.classType.toLowerCase() !== normalizedClass.key.toLowerCase()) {
            return false;
        }
        
        // Match time (allow some flexibility)
        if (!this.isTimeMatch(theme.time, classItem.Time)) {
            return false;
        }
        
        return true;
    }

    /**
     * Check if two times match (with some flexibility for formatting differences)
     */
    isTimeMatch(time1, time2) {
        if (!time1 || !time2) return false;
        
        const normalized1 = this.normalizeTime(time1);
        const normalized2 = this.normalizeTime(time2);
        
        // Direct match
        if (normalized1 === normalized2) {
            return true;
        }
        
        // Parse both times to compare as time objects
        const parseTime = (timeStr) => {
            const parts = timeStr.split(':');
            if (parts.length === 2) {
                return {
                    hours: parseInt(parts[0]),
                    minutes: parseInt(parts[1])
                };
            }
            return null;
        };
        
        const t1 = parseTime(normalized1);
        const t2 = parseTime(normalized2);
        
        if (t1 && t2) {
            // Check if they're the same time (allowing for different formats)
            return t1.hours === t2.hours && t1.minutes === t2.minutes;
        }
        
        return false;
    }

    /**
     * Comprehensive validation of theme mapping
     */
    validateThemeMapping(scheduleData, themes) {
        const validationResult = {
            isValid: true,
            errors: [],
            warnings: [],
            stats: {
                totalClasses: scheduleData.length,
                classesWithThemes: scheduleData.filter(c => c.Theme).length,
                powerCycleClasses: scheduleData.filter(c => 
                    this.normalizeClassName(c.Class)?.key === 'powercycle').length,
                mappedPowerCycleThemes: 0
            }
        };
        
        // Check for PowerCycle classes without themes
        scheduleData.forEach((classItem, index) => {
            const normalizedClass = this.normalizeClassName(classItem.Class);
            if (normalizedClass?.key === 'powercycle') {
                if (classItem.Theme) {
                    validationResult.stats.mappedPowerCycleThemes++;
                } else {
                    validationResult.warnings.push({
                        type: 'missing_theme',
                        index,
                        message: `PowerCycle class at ${classItem.Day} ${classItem.Time} ${classItem.Location} has no theme`,
                        class: classItem
                    });
                }
            }
        });
        
        // Check for theme consistency
        const themeGroups = {};
        scheduleData.forEach((classItem, index) => {
            if (classItem.Theme) {
                const key = `${classItem.Day}-${classItem.Time}-${classItem.Location}`;
                if (!themeGroups[key]) themeGroups[key] = [];
                themeGroups[key].push({ index, class: classItem });
            }
        });
        
        // Check for duplicate or conflicting themes
        Object.entries(themeGroups).forEach(([key, group]) => {
            if (group.length > 1) {
                const themes = group.map(g => g.class.Theme);
                const uniqueThemes = [...new Set(themes)];
                if (uniqueThemes.length > 1) {
                    validationResult.errors.push({
                        type: 'conflicting_themes',
                        message: `Conflicting themes for classes at ${key}: ${uniqueThemes.join(', ')}`,
                        classes: group
                    });
                    validationResult.isValid = false;
                }
            }
        });
        
        return validationResult;
    }
}

export default EnhancedScheduleMapper;