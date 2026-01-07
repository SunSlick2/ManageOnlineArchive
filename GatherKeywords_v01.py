import win32com.client
import pandas as pd
from collections import defaultdict

# ----------------------------
# Configuration
# ----------------------------
STOPWORDS = {
    "is", "in", "the", "a", "an", "and", "or", "of", "to", "for", "on", "with", "at", "by"
}

OUTPUT_FILE = r"C:\Temp\subject_keyword_frequency.xlsx"

# ----------------------------
# Keyword generation
# ----------------------------
def generate_phrases_from_subject(subject, stopwords):
    """
    Generate all contiguous phrases from a subject line.
    - Lowercase only
    - Whitespace tokenization
    - Stopwords excluded only for single-word phrases
    - Each phrase returned once per subject
    """
    if not subject:
        return set()

    subject = subject.lower()
    words = subject.split()
    n = len(words)

    phrases = set()

    for i in range(n):
        for j in range(i + 1, n + 1):
            phrase = " ".join(words[i:j])

            # Exclude single-word stopwords
            if j - i == 1 and phrase in stopwords:
                continue

            phrases.add(phrase)

    return phrases

# ----------------------------
# Outlook access
# ----------------------------
def get_subjects_from_outlook_folder():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
    folder = inbox.Folders["ToDelete"].Folders["Manual"]

    subjects = []

    for item in folder.Items:
        # Only MailItem
        if item.Class == 43:
            subjects.append(item.Subject)

    return subjects

# ----------------------------
# Main processing
# ----------------------------
def build_keyword_frequency_table():
    subjects = get_subjects_from_outlook_folder()

    phrase_counts = defaultdict(int)

    for subject in subjects:
        phrases_in_subject = generate_phrases_from_subject(subject, STOPWORDS)

        # Count each phrase once per email
        for phrase in phrases_in_subject:
            phrase_counts[phrase] += 1

    # Convert to DataFrame
    df = pd.DataFrame(
        [(count, phrase) for phrase, count in phrase_counts.items()],
        columns=["Count", "Phrase"]
    )

    # Sort: highest count first, then longest phrase (better matching priority)
    df.sort_values(
        by=["Count", "Phrase"],
        ascending=[False, True],
        inplace=True
    )

    return df

# ----------------------------
# Export to Excel
# ----------------------------
def export_to_excel(df):
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Excel file written to: {OUTPUT_FILE}")

# ----------------------------
# Run
# ----------------------------
if __name__ == "__main__":
    df = build_keyword_frequency_table()
    export_to_excel(df)
