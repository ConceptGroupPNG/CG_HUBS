// cg-staff-face-api — selfie check for CG Time Attendance (experimental)
//
// Every request must carry the employee's Microsoft sign-in (ID token):
//   Authorization: Bearer <id token issued to the CG Time Attendance app>
// The function works out WHO is asking from that token — the phone never
// says whose face it is — so nobody can enrol or verify as someone else.
//
// Actions (POST JSON):
//   status                 -> { enrolled, admin }
//   enrol   { image }      -> first selfie for the signed-in employee (once only)
//   verify  { image }      -> { verified, similarity }  does this selfie match them?
//   list                   -> (admin) every enrolled employee
//   reset   { email }      -> (admin) remove an employee's face so they can re-enrol
//
// Photos are never stored — only Rekognition face prints, labelled with the
// employee's sign-in name (with "@" written as ":" because Rekognition labels
// can't contain "@"; e.g. mscully:concept.com.pg).

import {
  RekognitionClient,
  DetectFacesCommand,
  SearchFacesByImageCommand,
  IndexFacesCommand,
  DeleteFacesCommand,
  ListFacesCommand,
} from "@aws-sdk/client-rekognition";
import crypto from "node:crypto";

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const COLLECTION_ID = process.env.COLLECTION_ID || "cg-staff";
const MATCH_THRESHOLD = Number(process.env.MATCH_THRESHOLD || 90);
const ADMIN_USERS = (process.env.ADMIN_USERS || "").toLowerCase().split(",").map((s) => s.trim()).filter(Boolean);

const rek = new RekognitionClient({});

class AuthError extends Error {}
class ForbiddenError extends Error {}
class UserError extends Error {
  constructor(message, code, status = 400) {
    super(message);
    this.code = code;
    this.status = status;
  }
}

const reply = (statusCode, obj) => ({
  statusCode,
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(obj),
});

// ---------- Microsoft token check ----------

let jwksCache = null;
let jwksFetchedAt = 0;

async function getSigningKeys(force = false) {
  const sixHours = 6 * 60 * 60 * 1000;
  if (!force && jwksCache && Date.now() - jwksFetchedAt < sixHours) return jwksCache;
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`);
  if (!res.ok) throw new Error(`Could not fetch Microsoft signing keys (${res.status})`);
  jwksCache = (await res.json()).keys;
  jwksFetchedAt = Date.now();
  return jwksCache;
}

const b64url = (s) => Buffer.from(s, "base64url");

async function verifyToken(authHeader) {
  if (!authHeader || !authHeader.startsWith("Bearer ")) throw new AuthError("Missing sign-in token");
  const parts = authHeader.slice(7).trim().split(".");
  if (parts.length !== 3) throw new AuthError("Malformed token");

  let header, payload;
  try {
    header = JSON.parse(b64url(parts[0]).toString("utf8"));
    payload = JSON.parse(b64url(parts[1]).toString("utf8"));
  } catch {
    throw new AuthError("Malformed token");
  }
  if (header.alg !== "RS256") throw new AuthError("Unexpected token algorithm");

  let jwk = (await getSigningKeys()).find((k) => k.kid === header.kid);
  if (!jwk) jwk = (await getSigningKeys(true)).find((k) => k.kid === header.kid); // keys rotate
  if (!jwk) throw new AuthError("Unknown signing key");

  const publicKey = crypto.createPublicKey({ key: { kty: jwk.kty, n: jwk.n, e: jwk.e }, format: "jwk" });
  const ok = crypto.verify("RSA-SHA256", Buffer.from(`${parts[0]}.${parts[1]}`), publicKey, b64url(parts[2]));
  if (!ok) throw new AuthError("Invalid token signature");

  const now = Math.floor(Date.now() / 1000);
  if (!payload.exp || now > payload.exp + 60) throw new AuthError("Token expired");
  if (payload.nbf && now < payload.nbf - 60) throw new AuthError("Token not yet valid");
  if (payload.aud !== CLIENT_ID) throw new AuthError("Token is for a different app");
  if (payload.tid !== TENANT_ID) throw new AuthError("Token is from a different organisation");
  if (payload.iss !== `https://login.microsoftonline.com/${TENANT_ID}/v2.0`) throw new AuthError("Unexpected token issuer");

  const email = String(payload.preferred_username || payload.email || "").trim().toLowerCase();
  if (!email) throw new AuthError("Sign-in has no user name");
  return { email, label: emailToLabel(email), admin: ADMIN_USERS.includes(email) };
}

// Rekognition labels allow letters, digits and _ . - : only.
function emailToLabel(email) {
  const label = String(email).trim().toLowerCase().replace("@", ":");
  if (!/^[a-z0-9_.\-:]{3,255}$/.test(label)) throw new UserError("This sign-in name can't be used for a face record", "BAD_NAME");
  return label;
}
const labelToEmail = (label) => String(label).replace(":", "@");

// ---------- Face helpers ----------

function imageBytes(b64) {
  if (!b64) throw new UserError("No photo supplied", "NO_IMAGE");
  const bytes = Buffer.from(String(b64).replace(/^data:image\/\w+;base64,/, ""), "base64");
  if (bytes.length < 1000) throw new UserError("Photo is empty or unreadable", "BAD_IMAGE");
  if (bytes.length > 4.5 * 1024 * 1024) throw new UserError("Photo is too large", "TOO_LARGE");
  return bytes;
}

async function requireOneFace(bytes) {
  const r = await rek.send(new DetectFacesCommand({ Image: { Bytes: bytes } }));
  const faces = (r.FaceDetails || []).filter((f) => f.Confidence >= 90);
  if (faces.length === 0) throw new UserError("No face found. Hold the phone at eye level and look at the screen.", "NO_FACE");
  if (faces.length > 1) throw new UserError("More than one face in the photo. Only you should be in the picture.", "MULTIPLE_FACES");
}

function inputError(err) {
  if (err?.name === "InvalidParameterException" || err?.name === "InvalidImageFormatException") {
    return new UserError("The photo wasn't clear enough. Please try again in better light.", "LOW_QUALITY");
  }
  return err;
}

async function allFaces() {
  const faces = [];
  let NextToken;
  do {
    const r = await rek.send(new ListFacesCommand({ CollectionId: COLLECTION_ID, MaxResults: 1000, NextToken }));
    for (const f of r.Faces || []) faces.push({ faceId: f.FaceId, label: f.ExternalImageId || "" });
    NextToken = r.NextToken;
  } while (NextToken);
  return faces;
}

async function facesFor(label) {
  return (await allFaces()).filter((f) => f.label === label);
}

async function search(bytes) {
  try {
    const r = await rek.send(new SearchFacesByImageCommand({
      CollectionId: COLLECTION_ID,
      Image: { Bytes: bytes },
      FaceMatchThreshold: MATCH_THRESHOLD,
      MaxFaces: 5,
      QualityFilter: "AUTO",
    }));
    return (r.FaceMatches || []).map((m) => ({ label: m.Face.ExternalImageId, similarity: m.Similarity }));
  } catch (err) {
    throw inputError(err);
  }
}

// ---------- Actions ----------

async function status(user) {
  const mine = await facesFor(user.label);
  return { enrolled: mine.length > 0, admin: user.admin };
}

async function enrol(user, b64) {
  if ((await facesFor(user.label)).length) {
    throw new UserError("Your face is already set up. Ask an administrator to reset it if it needs changing.", "ALREADY_ENROLLED", 409);
  }
  const bytes = imageBytes(b64);
  await requireOneFace(bytes);

  // Refuse a face that already belongs to a different employee.
  const other = (await search(bytes)).find((m) => m.label !== user.label);
  if (other) {
    throw new UserError("This face is already set up for another employee. Please see your administrator.", "FACE_IN_USE", 409);
  }

  try {
    const r = await rek.send(new IndexFacesCommand({
      CollectionId: COLLECTION_ID,
      Image: { Bytes: bytes },
      ExternalImageId: user.label,
      MaxFaces: 1,
      QualityFilter: "AUTO",
      DetectionAttributes: [],
    }));
    if (!r.FaceRecords?.length) throw new UserError("The photo wasn't clear enough to set up your face. Please try again.", "LOW_QUALITY");
    return { enrolled: true };
  } catch (err) {
    throw inputError(err);
  }
}

async function verify(user, b64) {
  if (!(await facesFor(user.label)).length) {
    throw new UserError("Your face isn't set up yet.", "NOT_ENROLLED", 409);
  }
  const bytes = imageBytes(b64);
  await requireOneFace(bytes);
  const matches = await search(bytes);
  const mine = matches.find((m) => m.label === user.label);
  const best = matches[0];
  // Verified only if this employee's face is the best match.
  if (mine && best && best.label === user.label) {
    return { verified: true, similarity: Math.round(mine.similarity * 10) / 10 };
  }
  return { verified: false };
}

async function list(user) {
  if (!user.admin) throw new ForbiddenError("Only administrators can do this");
  const counts = new Map();
  for (const f of await allFaces()) counts.set(f.label, (counts.get(f.label) || 0) + 1);
  return {
    employees: [...counts].map(([label, faces]) => ({ email: labelToEmail(label), faces }))
      .sort((a, b) => a.email.localeCompare(b.email)),
  };
}

async function reset(user, email) {
  if (!user.admin) throw new ForbiddenError("Only administrators can do this");
  const ids = (await facesFor(emailToLabel(email))).map((f) => f.faceId);
  if (ids.length) await rek.send(new DeleteFacesCommand({ CollectionId: COLLECTION_ID, FaceIds: ids }));
  return { removed: ids.length };
}

// ---------- Entry point ----------

export const handler = async (event) => {
  try {
    if (!TENANT_ID || !CLIENT_ID) throw new Error("TENANT_ID and CLIENT_ID environment variables must be set");
    if (event.requestContext?.http?.method !== "POST") return reply(405, { error: "Use POST" });

    const user = await verifyToken(event.headers?.authorization);

    let body = event.body || "{}";
    if (event.isBase64Encoded) body = Buffer.from(body, "base64").toString("utf8");
    let req;
    try {
      req = JSON.parse(body);
    } catch {
      return reply(400, { error: "Body must be JSON", code: "BAD_JSON" });
    }

    switch (req.action) {
      case "status": return reply(200, await status(user));
      case "enrol":  return reply(200, await enrol(user, req.image));
      case "verify": return reply(200, await verify(user, req.image));
      case "list":   return reply(200, await list(user));
      case "reset":  return reply(200, await reset(user, req.email));
      default:       return reply(400, { error: "Unknown action", code: "BAD_ACTION" });
    }
  } catch (err) {
    if (err instanceof AuthError) return reply(401, { error: err.message, code: "UNAUTHORISED" });
    if (err instanceof ForbiddenError) return reply(403, { error: err.message, code: "FORBIDDEN" });
    if (err instanceof UserError) return reply(err.status, { error: err.message, code: err.code });
    console.error(err);
    return reply(500, { error: "Server error", code: "SERVER_ERROR" });
  }
};
