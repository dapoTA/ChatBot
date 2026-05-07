import { createCipheriv, createDecipheriv, randomBytes } from "crypto";

const ALGORITHM = "aes-256-gcm";
const ENC_PREFIX = "enc:";

function getKey() {
  const hex = process.env.ENCRYPTION_KEY;
  if (!hex || hex.length !== 64) return null;
  return Buffer.from(hex, "hex");
}

/**
 * Encrypt a plain-text string.
 * Returns a string in the form  enc:<iv_hex>:<authtag_hex>:<ciphertext_hex>
 * If ENCRYPTION_KEY is not set the value is stored as-is (plain text fallback).
 */
export function encrypt(plain) {
  if (!plain) return plain;
  const key = getKey();
  if (!key) return plain; // graceful degradation — warn at startup instead

  const iv = randomBytes(16);
  const cipher = createCipheriv(ALGORITHM, key, iv);
  const ct = Buffer.concat([cipher.update(plain, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();

  return `${ENC_PREFIX}${iv.toString("hex")}:${tag.toString("hex")}:${ct.toString("hex")}`;
}

/**
 * Decrypt a value produced by encrypt().
 * If the value doesn't start with the enc: prefix it is returned as-is
 * (handles plain-text values already in the DB before encryption was added).
 */
export function decrypt(value) {
  if (!value || !value.startsWith(ENC_PREFIX)) return value;
  const key = getKey();
  if (!key) return value; // key missing — return raw ciphertext rather than crash

  const rest = value.slice(ENC_PREFIX.length);
  const [ivHex, tagHex, ctHex] = rest.split(":");
  if (!ivHex || !tagHex || !ctHex) return value; // malformed — return as-is

  try {
    const decipher = createDecipheriv(ALGORITHM, key, Buffer.from(ivHex, "hex"));
    decipher.setAuthTag(Buffer.from(tagHex, "hex"));
    const plain = Buffer.concat([
      decipher.update(Buffer.from(ctHex, "hex")),
      decipher.final(),
    ]);
    return plain.toString("utf8");
  } catch {
    // Auth tag mismatch or corrupt data — surface the error clearly
    throw new Error("Failed to decrypt credential. Check that ENCRYPTION_KEY has not changed.");
  }
}

/**
 * Log a startup warning if the encryption key is missing or malformed.
 */
export function warnIfKeyMissing() {
  const hex = process.env.ENCRYPTION_KEY;
  if (!hex) {
    console.warn("[crypto] ENCRYPTION_KEY is not set — passwords will be stored in plain text.");
  } else if (hex.length !== 64) {
    console.warn(`[crypto] ENCRYPTION_KEY must be 64 hex characters (32 bytes). Got ${hex.length} chars.`);
  } else {
    console.log("[crypto] ENCRYPTION_KEY is set — passwords encrypted with AES-256-GCM.");
  }
}
