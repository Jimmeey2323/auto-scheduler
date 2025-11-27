// Test the cover logic
const coverRaw = "Simran";
const trainerRaw = "Richard";

console.log('coverRaw:', coverRaw);
console.log('coverRaw type:', typeof coverRaw);
console.log('coverRaw.trim():', coverRaw.trim());
console.log('toLowerCase():', coverRaw.toString().toLowerCase());
console.log('Is undefined?:', coverRaw.toString().toLowerCase() === 'undefined');

if (coverRaw && coverRaw.toString().trim() && coverRaw.toString().toLowerCase() !== 'undefined') {
  console.log('✅ Cover would be applied');
  console.log('Cover:', coverRaw, 'for', trainerRaw);
} else {
  console.log('❌ Cover would NOT be applied');
}
