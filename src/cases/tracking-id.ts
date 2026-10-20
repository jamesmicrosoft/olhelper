const maximumTrackingIdLength = 64;
const trackingIdPattern =
  /\bTrackingID#([A-Za-z0-9-]{1,64})(?![A-Za-z0-9-])/i;

// Fallback for real case emails that do not carry a literal "TrackingID#" token
// (for example "VDM has assigned SR 2609100050001687001 to ..."). Match a bare
// support-case / SR number of 15 to 20 digits. A 19+ digit collaboration number
// is reduced to its 16-digit main case number so every activity on the same case
// shares one folder and routes together. Mirrors the proven VBA tool behavior.
const caseNumberPattern = /(?<!\d)(\d{15,20})(?!\d)/;

function normalizeCaseNumber(raw: string): string {
  return raw.length >= 19 ? raw.slice(0, 16) : raw;
}

export function extractTrackingId(subject: string): string | null {
  const tokenMatch = trackingIdPattern.exec(subject);
  if (tokenMatch?.[1]) {
    return tokenMatch[1];
  }

  const numberMatch = caseNumberPattern.exec(subject);
  if (numberMatch?.[1]) {
    return normalizeCaseNumber(numberMatch[1]);
  }

  return null;
}

export function trackingToken(trackingId: string): string {
  return `TrackingID#${trackingId}`;
}

export function caseFolderName(trackingId: string): string {
  if (
    trackingId.length > maximumTrackingIdLength ||
    !/^[A-Za-z0-9-]+$/.test(trackingId)
  ) {
    throw new Error("Tracking ID contains unsupported folder characters.");
  }

  return trackingId;
}