import re


def clean_filename(filename):

    invalid_chars = r'[\\/*?:"<>|]'

    cleaned = re.sub(
        invalid_chars,
        "_",
        filename
    )

    cleaned = cleaned.strip()

    return cleaned