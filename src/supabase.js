import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = process.env.REACT_APP_SUPABASE_URL || 'https://jqmdokvsmeundmsjkjsl.supabase.co';
const SUPABASE_KEY = process.env.REACT_APP_SUPABASE_KEY || 'sb_publishable_tOx6zCi0idR6BRqiTkXi6w_X8qJL2QC';

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
