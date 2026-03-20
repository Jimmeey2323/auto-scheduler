#!/usr/bin/env node

/**
 * Simple Recovery class extraction checker
 * This will help identify where Recovery classes are being lost in the processing pipeline
 */

import ScheduleUpdater from './updateKempsSchedule.js';
import fs from 'fs';
import path from 'path';

async function main() {
    console.log('🔍 RECOVERY CLASS DEBUG SESSION');
    console.log('='.repeat(60));
    
    try {
        const updater = new ScheduleUpdater();
        
        // Test basic functionality
        console.log('\n✅ ScheduleUpdater initialized');
        
        // Test Recovery class normalization
        console.log('\n🧪 Testing Recovery class normalization:');
        const recoveryVariations = ['recovery', 'Recovery', 'RECOVERY', 'barre recovery'];
        for (const variation of recoveryVariations) {
            const normalized = updater.normalizeClassNameForCleaned(variation);
            console.log(`   "${variation}" → "${normalized}"`);
        }
        
        // Test class validation
        console.log('\n🧪 Testing class name validation:');
        const testClasses = ['recovery', 'Recovery', 'Studio Recovery', 'Barre Recovery'];
        for (const className of testClasses) {
            const isValid = updater.isValidClassName(className);
            console.log(`   "${className}" → ${isValid ? '✅ Valid' : '❌ Invalid'}`);
        }
        
        console.log('\n🔍 Checking for existing index files...');
        
        // Check if index 3.html contains Recovery
        const indexFile = 'index 3.html';
        if (fs.existsSync(indexFile)) {
            const content = fs.readFileSync(indexFile, 'utf8');
            const hasRecovery = content.toLowerCase().includes('recovery');
            console.log(`   ${indexFile}: ${hasRecovery ? '✅ Contains Recovery' : '❌ No Recovery found'}`);
            
            if (hasRecovery) {
                // Show Recovery context
                const lines = content.split('\n');
                lines.forEach((line, idx) => {
                    if (line.toLowerCase().includes('recovery')) {
                        console.log(`     Line ${idx + 1}: ${line.trim()}`);
                    }
                });
            }
        } else {
            console.log(`   ❌ ${indexFile} not found`);
        }
        
        console.log('\n💡 Next steps:');
        console.log('   1. Check your recent schedule emails for hosted classes');
        console.log('   2. Verify Recovery classes are in the "hosted classes:" section');
        console.log('   3. Run the main schedule update to see processing logs');
        
    } catch (error) {
        console.error('❌ Error:', error.message);
        console.log('\n💡 This might be an authentication issue.');
        console.log('   Make sure your Gmail API credentials are properly configured.');
    }
}

main().catch(console.error);