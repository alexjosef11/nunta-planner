import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = 'https://blqgnqfjqbsnwzzfihvb.supabase.co'
const SUPABASE_KEY = 'sb_publishable_Zh_DptBaGAymMb9x8E0MwQ_RmLSoUXV'

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY)

// ── Generic key-value helpers ─────────────────────────────────────────────────
export async function dbGet(key) {
  const { data, error } = await supabase
    .from('wedding_data')
    .select('value')
    .eq('key', key)
    .single()
  if (error || !data) return null
  return data.value
}

export async function dbSet(key, value) {
  const { error } = await supabase
    .from('wedding_data')
    .upsert({ key, value, updated_at: new Date().toISOString() }, { onConflict: 'key' })
  if (error) console.error('dbSet error:', error)
}

// ── Realtime subscription ─────────────────────────────────────────────────────
export function dbSubscribe(key, callback) {
  const channel = supabase
    .channel(`wedding_data:${key}`)
    .on('postgres_changes', {
      event: '*',
      schema: 'public',
      table: 'wedding_data',
      filter: `key=eq.${key}`
    }, (payload) => {
      if (payload.new?.value !== undefined) callback(payload.new.value)
    })
    .subscribe()
  return () => supabase.removeChannel(channel)
}
