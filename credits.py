"""
Credit-tracking for the 50-files-per-Rs.20 model.

Takes a `db` object shaped like a real firebase_admin.firestore.Client
(specifically: db.collection(name).document(id) returning something with
.get() -> a snapshot with .exists / .to_dict(), and .update(dict)) — this
is exactly the real API (verified against the installed firebase_admin
package's DocumentReference.get/.update signatures), so the same code
works against a mock for testing here and the real client once wired up.

Firestore document shape assumed (licenses_resultportions/{licenseKey}):
    licenseKey       (string, matches the document ID)
    active           (bool)
    expiryDate       (string, DD-MM-YYYY)
    creditsRemaining (int) -- NEW field this module adds

ASSUMPTION FLAGGED: "creditsRemaining" is a field name chosen to match the
existing camelCase style (licenseKey/expiryDate) — it doesn't exist yet in
the real collection and needs to be added (starting at 50) whenever a key
is created or topped up, e.g. via the Admin Dashboard.

Deliberately NOT using a Firestore transaction: a true transaction guards
against two simultaneous requests on the SAME license key double-spending
credits, but that's a rare race for this app's real usage pattern (one
court, one computer, sequential use) — a plain get-then-update is simpler
to reason about and test. Flagging this trade-off rather than silently
picking the more complex option.
"""


class InsufficientCredits(Exception):
    def __init__(self, remaining, requested):
        self.remaining = remaining
        self.requested = requested
        super().__init__(
            f"Not enough credits: {remaining} remaining, {requested} files uploaded. "
            f"Please purchase more credits to continue."
        )


class LicenseNotFound(Exception):
    pass


class LicenseInactive(Exception):
    pass


def check_credits(db, license_key, file_count):
    """Read-only: confirm the license exists, is active, and has enough
    credits for this batch. Does NOT deduct anything — call this early,
    before doing any expensive work, so an invalid/exhausted license fails
    fast. Pair with deduct_credits() called only after the batch actually
    finishes successfully, so a mid-batch failure (e.g. LibreOffice
    missing) never charges for work that wasn't delivered.

    Small accepted trade-off: two simultaneous requests on the same
    license key could both pass this check before either deducts, since
    the check and the deduction aren't atomic together — the same rare
    race already accepted elsewhere in this module, for the same reason
    (single court, sequential usage in practice)."""
    doc_ref = db.collection("licenses_resultportions").document(license_key)
    snapshot = doc_ref.get()

    if not snapshot.exists:
        raise LicenseNotFound(f"License key '{license_key}' not found.")

    data = snapshot.to_dict()
    if data.get("active") is False:
        raise LicenseInactive("This license key has been deactivated.")

    remaining = data.get("creditsRemaining", 0)
    if remaining < file_count:
        raise InsufficientCredits(remaining, file_count)

    return remaining


def deduct_credits(db, license_key, file_count):
    """Actually deduct file_count credits. Call only after check_credits()
    passed AND the batch has actually finished — see check_credits'
    docstring for why the two are kept separate rather than combined."""
    doc_ref = db.collection("licenses_resultportions").document(license_key)
    snapshot = doc_ref.get()
    if not snapshot.exists:
        raise LicenseNotFound(f"License key '{license_key}' not found.")
    remaining = snapshot.to_dict().get("creditsRemaining", 0)
    new_remaining = max(0, remaining - file_count)  # never go below 0, even under the race above
    doc_ref.update({"creditsRemaining": new_remaining})
    return new_remaining


def check_and_reserve_credits(db, license_key, file_count):
    """Convenience wrapper combining check + immediate deduct, for callers
    that don't need the two-phase check-early/deduct-late split (e.g. a
    simple script or the Admin Dashboard). The /api/process route uses
    check_credits() and deduct_credits() separately instead — see there
    for why."""
    check_credits(db, license_key, file_count)
    return deduct_credits(db, license_key, file_count)


# ---------------------------------------------------------------------------
# Free trial
#
# A trial is stored as an ordinary licence record in the SAME collection, so
# check_credits() / deduct_credits() above work on it with no changes at all.
# The record id (= the licence key) is derived from the computer's browser
# fingerprint, which is what makes it "one trial per computer": asking again
# from the same computer finds the existing record instead of creating a new
# one, and a used-up trial stays used up.
# ---------------------------------------------------------------------------
import re
from datetime import datetime, timedelta, timezone

_IST = timezone(timedelta(hours=5, minutes=30))
_CID_RE = re.compile(r"^CID-[0-9A-F]{1,8}$")


class InvalidComputerId(Exception):
    pass


class TrialUsedUp(Exception):
    pass


class TrialCapReached(Exception):
    pass


def trial_key_for(cid):
    """CID-1A2B3C -> TRIAL-001A2B3C (same rule the web pages use)."""
    cid = (cid or "").strip().upper()
    if not _CID_RE.match(cid):
        raise InvalidComputerId("Could not identify this computer. Please try again.")
    return "TRIAL-" + cid[4:].rjust(8, "0")


def _is_already_exists(exc):
    return type(exc).__name__ == "AlreadyExists"


def start_trial(db, collection, cid, credits, court_name="", daily_cap=None):
    """Give this computer its free trial, or hand back the one it already has.

    Returns {"key": ..., "creditsRemaining": n, "created": bool}.
    Raises TrialUsedUp if this computer's trial has no credits left,
    LicenseInactive if the trial was switched off, TrialCapReached if
    `daily_cap` new trials were already started today (protects against
    anyone farming free trials on a feature that costs real money to run),
    InvalidComputerId for a malformed computer id.
    """
    key = trial_key_for(cid)
    ref = db.collection(collection).document(key)

    snap = ref.get()
    if snap.exists:
        return _existing_trial(snap.to_dict(), key)

    counter_ref = None
    count = 0
    if daily_cap:
        day = datetime.now(_IST).strftime("%Y%m%d")
        counter_ref = db.collection("trial_counters").document(collection + "_" + day)
        csnap = counter_ref.get()
        count = (csnap.to_dict() or {}).get("count", 0) if csnap.exists else 0
        if count >= daily_cap:
            raise TrialCapReached("Today's free trials are full.")

    court = "".join(ch for ch in str(court_name or "") if ch.isprintable()).strip()[:80]
    data = {
        "licenseKey": key,
        "courtName": court or "Free Trial",
        "active": True,
        "expiryDate": "",
        "creditsRemaining": int(credits),
        "createdDate": datetime.now(_IST).strftime("%d/%m/%Y"),
        "trial": True,
        "computerID": cid.strip().upper(),
    }
    try:
        ref.create(data)  # fails if another request created it a moment ago
    except Exception as e:
        if _is_already_exists(e):
            return _existing_trial(ref.get().to_dict(), key)
        raise

    if counter_ref is not None:
        counter_ref.set({"count": count + 1})
    return {"key": key, "creditsRemaining": int(credits), "created": True}


def _existing_trial(data, key):
    data = data or {}
    if data.get("active") is False:
        raise LicenseInactive("This trial has been switched off.")
    remaining = data.get("creditsRemaining", 0)
    if data.get("trial") is True and remaining <= 0:
        raise TrialUsedUp("Your free trial has been used.")
    return {"key": key, "creditsRemaining": remaining, "created": False}
