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
