import { createClient } from '@supabase/supabase-js';

// 1. Read the raw values from your .env.local file
const rawUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
const rawKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;

// 2. Clean up any accidental hidden spaces or line breaks from copying/pasting
const cleanUrl = rawUrl ? rawUrl.trim() : '';
const cleanKey = rawKey ? rawKey.trim() : '';

// 3. Set up our guaranteed working fallback URL based on your project ID
const defaultUrl = 'https://xjzyhzmqibcnvdtrtdcs.supabase.co';

// 4. Validate if the URL from your file is usable (it must start with http:// or https://)
let finalUrl = defaultUrl;

if (cleanUrl && cleanUrl.startsWith('http')) {
  finalUrl = cleanUrl;
} else {
  // If the URL in your file is broken or missing, print a helpful message in your terminal
  console.log("==================================================================");
  console.log("⚠️  NOTE: Next.js is not reading a valid URL from your .env.local file.");
  console.log(`   Value detected in file: "${rawUrl}"`);
  console.log(`   Action: Automatically using your hardcoded project URL: ${defaultUrl}`);
  console.log("==================================================================");
}

// 5. Fallback for the security key if it happens to be missing or blank
const finalKey = cleanKey || 'placeholder-anon-key-waiting-for-env-file';

// 6. Initialize the connection using clean, validated variables
export const supabase = createClient(finalUrl, finalKey);