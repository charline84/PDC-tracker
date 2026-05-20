import { createClient } from '@supabase/supabase-js';

const rawUrl = import.meta.env.VITE_SUPABASE_URL || '';
const supabaseUrl = rawUrl.startsWith('http') ? rawUrl : (rawUrl ? `https://${rawUrl}` : '');
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || '';

const isDummy = supabaseUrl.includes('your_supabase') || supabaseAnonKey.includes('your_supabase');

export const supabase = supabaseUrl && supabaseAnonKey && !isDummy
  ? createClient(supabaseUrl, supabaseAnonKey) 
  : null;
