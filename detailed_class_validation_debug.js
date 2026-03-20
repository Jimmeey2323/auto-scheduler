#!/usr/bin/env node

/**
 * Detailed class validation debugger
 */

import ScheduleUpdater from './updateKempsSchedule.js';

async function main() {
    console.log('🔍 DETAILED CLASS VALIDATION DEBUG');
    console.log('='.repeat(60));
    
    const updater = new ScheduleUpdater();
    
    // Test cases
    const testClasses = [
        'recovery',
        'Recovery', 
        'RECOVERY',
        'Studio Recovery',
        'Barre Recovery',
        'studio recovery',
        'barre recovery',
        'PowerCycle',
        'Barre 57'
    ];
    
    console.log('\n🧪 Detailed validation analysis:');
    
    for (const className of testClasses) {
        console.log(`\n--- Testing: "${className}" ---`);
        
        // Step by step validation
        if (!className) {
            console.log('❌ Failed: name is empty');
            continue;
        }
        
        const val = className.toString().trim().toLowerCase();
        console.log(`   Normalized: "${val}"`);
        
        // Check invalid list
        const invalid = ['smita', 'anandita', 'cover', 'replacement', 'sakshi', 'parekh', 'taarika', 'host'];
        const hasInvalid = invalid.some(i => val.includes(i));
        if (hasInvalid) {
            console.log(`❌ Failed: contains invalid word from ${invalid}`);
            continue;
        } else {
            console.log('✅ Passed: no invalid words');
        }
        
        // Check if all digits
        if (/^\d+$/.test(val)) {
            console.log('❌ Failed: is all digits');
            continue;
        } else {
            console.log('✅ Passed: not all digits');
        }
        
        // Check short single word rule
        const words = val.split(' ');
        const wordCount = words.length;
        const length = val.length;
        console.log(`   Words: ${wordCount}, Length: ${length}`);
        
        if (wordCount === 1 && length < 3) {
            console.log('❌ Failed: single word with less than 3 characters');
            continue;
        } else {
            console.log('✅ Passed: not a short single word');
        }
        
        console.log('🎉 OVERALL RESULT: VALID');
        
        // Test actual function
        const actualResult = updater.isValidClassName(className);
        console.log(`   Function result: ${actualResult ? '✅ Valid' : '❌ Invalid'}`);
        
        if (actualResult) {
            console.log('✅ Function agrees with manual analysis');
        } else {
            console.log('🚨 FUNCTION DISAGREES! There might be a hidden issue.');
        }
    }
}

main().catch(console.error);