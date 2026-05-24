import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://oqjtepwhudumzrpuiegl.supabase.co";
const supabaseKey = "sb_publishable_pRk1YgSOf38ZeQ7iceTwpQ_ZNO_YuKA";

export const supabase = createClient(supabaseUrl, supabaseKey);