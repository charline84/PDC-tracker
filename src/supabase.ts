import { createClient } from '@supabase/supabase-js';

const rawUrl = import.meta.env.VITE_SUPABASE_URL || '';
let supabaseUrl = rawUrl.startsWith('http') ? rawUrl : (rawUrl ? `https://${rawUrl}` : '');

// Si l'utilisateur a collé l'URL de l'API REST qui finit par /rest/v1, on nettoie
if (supabaseUrl.endsWith('/rest/v1/') || supabaseUrl.endsWith('/rest/v1')) {
  supabaseUrl = supabaseUrl.replace(/\/rest\/v1\/?$/, '');
}

const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || '';

const isDummy = supabaseUrl.includes('your_supabase') || supabaseAnonKey.includes('your_supabase');

export const supabase = (supabaseUrl && supabaseAnonKey && !isDummy)
  ? createClient(supabaseUrl, supabaseAnonKey) 
  : null;
