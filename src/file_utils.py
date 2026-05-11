import re


def clean_filename(filename):

    invalid_chars = r'[\\/*?:"<>|]'

    cleaned = re.sub(
        invalid_chars,
        "_",
        filename
    )

    return cleaned.strip()