const SESSION_TIMEOUT_MS = 30 * 60 * 1000; // 30 minutes
const CLEANUP_INTERVAL_MS = 5 * 60 * 1000; // 5 minutes
const MAX_SESSIONS = 1000;
export class SessionManager {
    sessions = new Map();
    cleanupTimer;
    constructor() {
        this.cleanupTimer = setInterval(() => this.cleanupStaleSessions(), CLEANUP_INTERVAL_MS);
    }
    get(sessionId) {
        const entry = this.sessions.get(sessionId);
        if (entry) {
            entry.lastActivity = Date.now();
        }
        return entry;
    }
    has(sessionId) {
        return this.sessions.has(sessionId);
    }
    set(sessionId, entry) {
        if (this.sessions.size >= MAX_SESSIONS && !this.sessions.has(sessionId)) {
            throw new Error("Maximum session limit reached");
        }
        this.sessions.set(sessionId, entry);
    }
    delete(sessionId) {
        this.sessions.delete(sessionId);
    }
    get size() {
        return this.sessions.size;
    }
    cleanupStaleSessions() {
        const now = Date.now();
        for (const [sessionId, entry] of this.sessions) {
            if (now - entry.lastActivity > SESSION_TIMEOUT_MS) {
                console.log(`Cleaning up stale session: ${sessionId}`);
                entry.transport.close().catch((err) => {
                    console.error(`Error closing stale transport for session ${sessionId}:`, err);
                });
                this.sessions.delete(sessionId);
            }
        }
    }
    async shutdown() {
        if (this.cleanupTimer) {
            clearInterval(this.cleanupTimer);
            this.cleanupTimer = undefined;
        }
        const closePromises = [];
        for (const [sessionId, entry] of this.sessions) {
            console.log(`Closing session: ${sessionId}`);
            closePromises.push(entry.transport.close().catch((err) => {
                console.error(`Error closing transport for session ${sessionId}:`, err);
            }));
        }
        await Promise.all(closePromises);
        this.sessions.clear();
    }
}
//# sourceMappingURL=session-manager.js.map