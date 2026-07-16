/* ============================================================
   SUPABASE.JS
   Cliente compartilhado, leitura, gravação e atualização
   em tempo real do estado do aplicativo.
============================================================ */

const SUPABASE_URL =
    "https://uygdfrfyzvvygxvknfne.supabase.co";

const SUPABASE_PUBLISHABLE_KEY =
    "sb_publishable_gdm8ZWmq6MRk77KXWhEPLA_j7P-cEbB";

const SUPABASE_STATE_TABLE =
    "app_state";

const SUPABASE_STATE_ROW_ID =
    "main";

let supabaseClient = null;
let supabaseStateChannel = null;


/* ============================================================
   CLIENTE
============================================================ */

function initializeSupabaseClient() {
    if (supabaseClient) {
        return supabaseClient;
    }

    if (
        !window.supabase ||
        typeof window.supabase.createClient !== "function"
    ) {
        throw new Error(
            "A biblioteca do Supabase não foi carregada."
        );
    }

    supabaseClient =
        window.supabase.createClient(
            SUPABASE_URL,
            SUPABASE_PUBLISHABLE_KEY,
            {
                auth: {
                    persistSession: false,
                    autoRefreshToken: false,
                    detectSessionInUrl: false
                },
                db: {
                    schema: "public"
                }
            }
        );

    return supabaseClient;
}


/* ============================================================
   LEITURA
============================================================ */

async function fetchSharedAppState() {
    const client =
        initializeSupabaseClient();

    const {
        data,
        error
    } = await client
        .from(SUPABASE_STATE_TABLE)
        .select("data, updated_at")
        .eq("id", SUPABASE_STATE_ROW_ID)
        .maybeSingle();

    if (error) {
        throw error;
    }

    return data?.data || null;
}


/* ============================================================
   GRAVAÇÃO
============================================================ */

async function saveSharedAppState(appState) {
    const client =
        initializeSupabaseClient();

    const payload = {
        id: SUPABASE_STATE_ROW_ID,
        data: appState,
        updated_at:
            new Date().toISOString()
    };

    const {
        error
    } = await client
        .from(SUPABASE_STATE_TABLE)
        .upsert(
            payload,
            {
                onConflict: "id"
            }
        );

    if (error) {
        throw error;
    }

    return true;
}


/* ============================================================
   TEMPO REAL
============================================================ */

function subscribeToSharedAppState(
    onStateChange
) {
    const client =
        initializeSupabaseClient();

    if (supabaseStateChannel) {
        client.removeChannel(
            supabaseStateChannel
        );

        supabaseStateChannel = null;
    }

    supabaseStateChannel =
        client
            .channel(
                "app-state-main"
            )
            .on(
                "postgres_changes",
                {
                    event: "*",
                    schema: "public",
                    table:
                        SUPABASE_STATE_TABLE,
                    filter:
                        `id=eq.${SUPABASE_STATE_ROW_ID}`
                },
                payload => {
                    const newState =
                        payload?.new?.data;

                    if (
                        newState &&
                        typeof onStateChange === "function"
                    ) {
                        onStateChange(
                            newState,
                            payload
                        );
                    }
                }
            )
            .subscribe(status => {
                if (
                    status ===
                    "CHANNEL_ERROR"
                ) {
                    console.warn(
                        "O canal em tempo real do Supabase apresentou erro."
                    );
                }
            });

    return function unsubscribe() {
        if (!supabaseStateChannel) {
            return;
        }

        client.removeChannel(
            supabaseStateChannel
        );

        supabaseStateChannel = null;
    };
}
