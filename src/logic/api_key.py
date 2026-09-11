"""Validation for the Gemini API key.

Google changed the key format in 2026. The old keys were "standard keys":
they started with `AIza` and were exactly 39 characters long. The keys
Google AI Studio issues now are "auth keys", they start with `AQ.` and
they are longer.

The Gemini API stops accepting standard keys in September 2026, so anyone
generating a key today gets the new format. The app used to require
exactly 39 characters, which rejects every new key before it is ever sent.

The lesson is not "use the new length". An API key is an opaque string the
provider can rebrand whenever it likes, and the only component that can
truly validate one is Google. So we check the bare minimum needed to catch
a paste accident (an empty box, a stray sentence, a truncated copy) and
let the API be the judge of the rest.
"""

# Keys issued before the 2026 migration. Still accepted by the app, since
# some of them keep working until Google finishes the rollout, but worth
# warning about.
LEGACY_PREFIX = "AIza"

# Short enough to accept any plausible key, long enough to catch someone
# pasting a word or half a key.
MIN_KEY_LENGTH = 20

OK = "ok"
EMPTY = "empty"
MALFORMED = "malformed"
LEGACY = "legacy"


def check_key(key: str) -> str:
    """Classify a pasted API key.

    Returns one of:
      "ok"         looks like a key, send it
      "empty"      nothing was entered
      "malformed"  cannot be a key (too short, or contains whitespace)
      "legacy"     plausible, but in the pre-2026 format that Google is
                   retiring. Usable for now, so the caller should warn
                   rather than block.
    """
    if key is None:
        return EMPTY

    key = key.strip()
    if not key:
        return EMPTY

    # Whitespace inside the key means a partial copy or a pasted sentence.
    # No API key format has ever contained a space or a newline.
    if any(character.isspace() for character in key):
        return MALFORMED

    if len(key) < MIN_KEY_LENGTH:
        return MALFORMED

    if key.startswith(LEGACY_PREFIX):
        return LEGACY

    return OK
