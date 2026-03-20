#!/usr/bin/env node

/**
 * Debug script to check Recovery class extraction
 * This will help identify where Recovery classes are being lost in the processing pipeline
 */

const { KempsScheduleUpdater } = require('./updateKempsSchedule.js');

class RecoveryDebugger extends KempsScheduleUpdater {
    
    /**
     * Debug the email parsing process specifically for Recovery classes
     */
    async debugRecoveryExtraction() {
        console.log('🔍 RECOVERY CLASS DEBUG SESSION');
        console.log('='.repeat(60));
        
        try {
            // Step 1: Get recent emails
            console.log('\n📧 Step 1: Finding recent emails...');
            const messages = await this.findRecentScheduleEmails();
            
            if (!messages || messages.length === 0) {
                console.log('❌ No recent schedule emails found');
                return;
            }
            
            console.log(`✅ Found ${messages.length} recent email(s)`);
            
            // Step 2: Process each email for hosted classes
            for (let i = 0; i < messages.length; i++) {
                const message = messages[i];
                console.log(`\n📨 Processing email ${i + 1}: ${message.snippet.substring(0, 100)}...`);
                
                // Get email body
                const emailBody = await this.getEmailBody(message);
                console.log('\n📝 Email Body Preview:');
                console.log('-'.repeat(40));
                console.log(emailBody.substring(0, 2000));
                if (emailBody.length > 2000) {
                    console.log(`\n... (truncated, full length: ${emailBody.length} characters)`);
                }
                console.log('-'.repeat(40));
                
                // Check for Recovery in raw email
                const recoveryInEmail = this.checkForRecoveryInEmail(emailBody);
                
                // Parse email for schedule info
                console.log('\n🔍 Parsing email for schedule information...');
                const emailInfo = await this.parseEmailForScheduleInfo([emailBody], i === 0);
                
                console.log('\n📊 PARSING RESULTS:');
                console.log(`   Covers: ${emailInfo.covers?.length || 0}`);
                console.log(`   Themes: ${emailInfo.themes?.length || 0}`);
                console.log(`   Hosted Classes: ${emailInfo.hostedClasses?.length || 0}`);
                console.log(`   Changes: ${emailInfo.changes?.length || 0}`);
                
                // Debug hosted classes specifically
                if (emailInfo.hostedClasses && emailInfo.hostedClasses.length > 0) {
                    console.log('\n🏠 HOSTED CLASSES DETAILS:');
                    emailInfo.hostedClasses.forEach((hosted, idx) => {
                        const isRecovery = hosted.classType.toLowerCase().includes('recovery');
                        const marker = isRecovery ? '🔴 RECOVERY' : '   ';
                        console.log(`${marker} ${idx + 1}. ${hosted.day} ${hosted.time} - ${hosted.classType} - ${hosted.trainer} @ ${hosted.location}`);
                        
                        if (isRecovery) {
                            // Test normalization
                            const normalized = this.normalizeClassNameForCleaned(hosted.classType);
                            console.log(`     → Normalized to: "${normalized}"`);
                        }
                    });
                } else {
                    console.log('\n⚠️  No hosted classes found in email');
                }
                
                // Check specific regex patterns
                this.debugHostedClassRegex(emailBody);
            }
            
        } catch (error) {
            console.error('❌ Debug error:', error);
        }
    }
    
    /**
     * Check for Recovery mentions in email content
     */
    checkForRecoveryInEmail(emailBody) {
        const recoveryVariations = [
            'recovery', 'Recovery', 'RECOVERY',
            'studio recovery', 'Studio Recovery', 'STUDIO RECOVERY'
        ];
        
        console.log('\n🔍 Checking for Recovery mentions in email:');
        let found = false;
        
        for (const variation of recoveryVariations) {
            if (emailBody.includes(variation)) {
                console.log(`   ✅ Found: "${variation}"`);
                found = true;
                
                // Show context around the match
                const index = emailBody.indexOf(variation);
                const start = Math.max(0, index - 100);
                const end = Math.min(emailBody.length, index + variation.length + 100);
                const context = emailBody.substring(start, end);
                console.log(`       Context: ...${context}...`);
            }
        }
        
        if (!found) {
            console.log('   ❌ No Recovery mentions found in email content');
        }
        
        return found;
    }
    
    /**
     * Debug hosted class regex patterns
     */
    debugHostedClassRegex(emailBody) {
        console.log('\n🧪 Testing hosted class regex patterns...');
        
        // Look for the hosted classes section
        const hostedSectionRegex = /hosted\s+classes?\s*:?\s*(.*?)(?=\n\n|\n[A-Z]|$)/si;
        const hostedMatch = emailBody.match(hostedSectionRegex);
        
        if (hostedMatch) {
            console.log('\n✅ Found hosted classes section:');
            console.log('-'.repeat(30));
            console.log(hostedMatch[1]);
            console.log('-'.repeat(30));
            
            // Test individual line parsing
            const lines = hostedMatch[1].split('\n').filter(line => line.trim());
            console.log(`\n🔍 Testing ${lines.length} hosted class lines:`);
            
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();
                if (!line || line.startsWith('-')) continue;
                
                console.log(`\n   Line ${i + 1}: "${line}"`);
                const parsed = this.parseHostedLine(line);
                if (parsed) {
                    const isRecovery = parsed.classType.toLowerCase().includes('recovery');
                    const marker = isRecovery ? '🔴' : '✅';
                    console.log(`   ${marker} Parsed: ${parsed.day} ${parsed.time} - ${parsed.classType} - ${parsed.trainer}`);
                } else {
                    console.log(`   ❌ Failed to parse line`);
                }
            }
        } else {
            console.log('❌ No hosted classes section found');
        }
    }
}

// Run the debugger
async function main() {
    const recoveryDebugger = new RecoveryDebugger();
    await recoveryDebugger.debugRecoveryExtraction();
    process.exit(0);
}

if (require.main === module) {
    main().catch(console.error);
}

module.exports = RecoveryDebugger;