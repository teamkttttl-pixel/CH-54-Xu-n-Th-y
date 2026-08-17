import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://ovzhojzglooimtdkctzm.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_b_eQoj2TILnvdGUPL3ai7w_b-hE5Ia5";

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
