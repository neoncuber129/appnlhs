// Quản lý lưu trữ FileSystemFileHandle trong IndexedDB qua các phiên reload
const DB_NAME = 'ExcelDataEntryDB';
const DB_VERSION = 1;
const STORE_NAME = 'handles';
const KEY_ACTIVE_HANDLE = 'activeFileHandle';

function openDB() {
    return new Promise((resolve, reject) => {
        if (!('indexedDB' in window)) {
            resolve(null);
            return;
        }
        const req = indexedDB.open(DB_NAME, DB_VERSION);
        req.onupgradeneeded = () => {
            const db = req.result;
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                db.createObjectStore(STORE_NAME);
            }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}

export async function saveActiveFileHandle(handle) {
    if (!handle) return;
    try {
        const db = await openDB();
        if (!db) return;
        const tx = db.transaction(STORE_NAME, 'readwrite');
        tx.objectStore(STORE_NAME).put(handle, KEY_ACTIVE_HANDLE);
        await new Promise((res, rej) => {
            tx.oncomplete = res;
            tx.onerror = rej;
        });
    } catch (err) {
        console.warn('Không thể lưu fileHandle vào IndexedDB:', err);
    }
}

export async function getActiveFileHandle() {
    try {
        const db = await openDB();
        if (!db) return null;
        const tx = db.transaction(STORE_NAME, 'readonly');
        const req = tx.objectStore(STORE_NAME).get(KEY_ACTIVE_HANDLE);
        return await new Promise((res, rej) => {
            req.onsuccess = () => res(req.result || null);
            req.onerror = () => rej(req.error);
        });
    } catch (err) {
        console.warn('Không thể đọc fileHandle từ IndexedDB:', err);
        return null;
    }
}

export async function clearActiveFileHandle() {
    try {
        const db = await openDB();
        if (!db) return;
        const tx = db.transaction(STORE_NAME, 'readwrite');
        tx.objectStore(STORE_NAME).delete(KEY_ACTIVE_HANDLE);
        await new Promise((res, rej) => {
            tx.oncomplete = res;
            tx.onerror = rej;
        });
    } catch (err) {
        console.warn('Không thể xóa fileHandle trong IndexedDB:', err);
    }
}

/**
 * Nhận diện lỗi file đang bị khóa bởi tiến trình khác (Microsoft Excel, WPS Office, v.v.)
 */
export function isFileLockedError(err) {
    if (!err) return false;
    const msg = (err.message || (typeof err === 'string' ? err : '') || '').toLowerCase();
    const name = (err.name || '').toLowerCase();

    return (
        name === 'nomodificationallowederror' ||
        name === 'notallowederror' ||
        msg.includes('used by another process') ||
        msg.includes('being used') ||
        msg.includes('another process') ||
        msg.includes('sharing violation') ||
        msg.includes('locked') ||
        msg.includes('đang được mở') ||
        msg.includes('ứng dụng khác') ||
        msg.includes('quá trình khác') ||
        msg.includes('could not be modified')
    );
}

