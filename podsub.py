import os
import re
import math
from datetime import timedelta
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document

# ==========================================
# TIME & PARSING UTILITIES
# ==========================================

def parse_time(time_str):
    """
    Converts a tolerant subtitle timestamp into milliseconds.

    Accepted examples:
        00:00:01,000
        00:00:01.000
        00:00:01:000
        00:00:01
        0:00:01,5
        01:23,500
    """
    time_str = time_str.strip()

    # Normalize the final millisecond separator to a comma.
    time_str = re.sub(r'(?<=\d)[.,:](\d{1,3})$', r',\1', time_str)

    # Hours:minutes:seconds, optionally followed by milliseconds.
    match = re.fullmatch(
        r'(\d{1,3}):(\d{1,2}):(\d{1,2})(?:,(\d{1,3}))?',
        time_str
    )

    if match:
        hours, minutes, seconds, millis = match.groups()
    else:
        # Also accept minutes:seconds.
        match = re.fullmatch(
            r'(\d{1,3}):(\d{1,2})(?:,(\d{1,3}))?',
            time_str
        )

        if not match:
            raise ValueError(f"Invalid subtitle timestamp: {time_str!r}")

        minutes, seconds, millis = match.groups()
        hours = "0"

    millis = (millis or "0").ljust(3, "0")[:3]

    return (
        int(hours) * 3_600_000
        + int(minutes) * 60_000
        + int(seconds) * 1_000
        + int(millis)
    )

def format_time(ms):
    """Converts milliseconds back to SRT time string."""
    td = timedelta(milliseconds=ms)
    total_seconds = int(td.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    millis = int(ms % 1000)
    return f"{hours:02}:{minutes:02}:{seconds:02},{millis:03}"

def parse_srt_to_dicts(
    content,
    pseudo_start_ms=1000,
    pseudo_duration_ms=100
):
    """
    Tolerant SRT/plain-text parser.

    Rules:
    - Valid timestamps are preserved.
    - Corrupted or missing timestamps between valid blocks divide the available
      interval between the previous valid end and the next valid start.
    - Corrupted blocks before the first valid timestamp receive consecutive
      pseudo timestamps ending at the first valid start where possible.
    - Corrupted blocks after the last valid timestamp receive consecutive
      pseudo timestamps after the last valid end.
    - Completely timestamp-free text receives consecutive 100 ms timestamps.
    """

    if not content or not content.strip():
        return []

    content = content.replace("\ufeff", "")
    content = content.replace("\r\n", "\n").replace("\r", "\n").strip()

    time_pattern = re.compile(
        r'(?P<start>\d{1,3}:\d{1,2}(?::\d{1,2})?(?:[.,:]\d{1,3})?)'
        r'\s*(?:-->|->|–>|—>|→)\s*'
        r'(?P<end>\d{1,3}:\d{1,2}(?::\d{1,2})?(?:[.,:]\d{1,3})?)'
    )

    raw_blocks = re.split(r'\n[ \t]*\n+', content)

    # Completely plain text without blank lines:
    # treat each non-empty line as a separate block.
    if len(raw_blocks) == 1 and not time_pattern.search(content):
        raw_blocks = [
            line.strip()
            for line in content.splitlines()
            if line.strip()
        ]

    blocks = []

    for raw_block in raw_blocks:
        lines = [
            line.strip()
            for line in raw_block.splitlines()
            if line.strip()
        ]

        if not lines:
            continue

        timestamp_match = None
        timestamp_line_index = None
        start_ms = None
        end_ms = None

        for index, line in enumerate(lines):
            match = time_pattern.search(line)

            if not match:
                continue

            try:
                candidate_start = parse_time(match.group("start"))
                candidate_end = parse_time(match.group("end"))

                if candidate_end <= candidate_start:
                    candidate_end = candidate_start + pseudo_duration_ms

                start_ms = candidate_start
                end_ms = candidate_end
                timestamp_match = match
                timestamp_line_index = index
                break

            except ValueError:
                continue

        text_lines = []

        for index, line in enumerate(lines):
            # Remove valid timestamp line.
            if timestamp_match and index == timestamp_line_index:
                continue

            # Remove the subtitle number at the beginning of a block.
            if (
                index == 0
                and len(lines) > 1
                and re.fullmatch(r'\d+', line)
            ):
                continue

            # Remove malformed timestamp-looking lines such as:
            # "bad timestamp", provided there is other text in the block.
            if (
                len(lines) > 1
                and (
                    re.search(r'-->|->|–>|—>|→', line)
                    or re.fullmatch(
                        r'(?i)(?:bad|invalid|broken|missing)\s+timestamp',
                        line
                    )
                )
            ):
                continue

            text_lines.append(line)

        text = " ".join(text_lines).strip()

        if not text:
            continue

        blocks.append({
            "start": start_ms,
            "end": end_ms,
            "text": text,
            "has_valid_time": timestamp_match is not None
        })

    if not blocks:
        return []

    # Assign timestamps to every consecutive run of invalid blocks.
    index = 0

    while index < len(blocks):
        if blocks[index]["has_valid_time"]:
            index += 1
            continue

        run_start = index

        while (
            index < len(blocks)
            and not blocks[index]["has_valid_time"]
        ):
            index += 1

        run_end = index
        run_length = run_end - run_start

        previous_valid_end = None
        next_valid_start = None

        if run_start > 0 and blocks[run_start - 1]["has_valid_time"]:
            previous_valid_end = blocks[run_start - 1]["end"]

        if run_end < len(blocks) and blocks[run_end]["has_valid_time"]:
            next_valid_start = blocks[run_end]["start"]

        # Invalid blocks between two valid blocks:
        # divide the whole available interval evenly.
        if (
            previous_valid_end is not None
            and next_valid_start is not None
            and next_valid_start > previous_valid_end
        ):
            available_duration = next_valid_start - previous_valid_end

            for offset in range(run_length):
                block_start = (
                    previous_valid_end
                    + available_duration * offset // run_length
                )
                block_end = (
                    previous_valid_end
                    + available_duration * (offset + 1) // run_length
                )

                blocks[run_start + offset]["start"] = block_start
                blocks[run_start + offset]["end"] = block_end

        # Invalid blocks before the first valid block.
        elif next_valid_start is not None:
            total_needed = run_length * pseudo_duration_ms
            first_start = max(
                0,
                next_valid_start - total_needed
            )

            for offset in range(run_length):
                block_start = first_start + offset * pseudo_duration_ms

                blocks[run_start + offset]["start"] = block_start
                blocks[run_start + offset]["end"] = (
                    block_start + pseudo_duration_ms
                )

        # Invalid blocks after the last valid block.
        elif previous_valid_end is not None:
            for offset in range(run_length):
                block_start = (
                    previous_valid_end
                    + offset * pseudo_duration_ms
                )

                blocks[run_start + offset]["start"] = block_start
                blocks[run_start + offset]["end"] = (
                    block_start + pseudo_duration_ms
                )

        # No valid timestamps anywhere in the file.
        else:
            for offset in range(run_length):
                block_start = (
                    pseudo_start_ms
                    + offset * pseudo_duration_ms
                )

                blocks[run_start + offset]["start"] = block_start
                blocks[run_start + offset]["end"] = (
                    block_start + pseudo_duration_ms
                )

    return [
        {
            "start": block["start"],
            "end": block["end"],
            "text": block["text"]
        }
        for block in blocks
    ]

def dicts_to_srt_string(subs_list):
    """Converts a list of subtitle dictionaries back into an SRT formatted string."""
    output_lines = []
    for idx, sub in enumerate(subs_list, 1):
        if sub:
            text_with_newlines = sub['text'] # Keeps [pause] intact in the SRT
            output_lines.append(f"{idx}\n{format_time(sub['start'])} --> {format_time(sub['end'])}\n{text_with_newlines}")
    return "\n\n".join(output_lines) + "\n\n"


def contains_valid_subtitle_timestamp(content):
    """
    Returns True when the input contains at least one usable subtitle
    timestamp range.

    This uses the same tolerant formats as parse_srt_to_dicts().
    """
    time_pattern = re.compile(
        r'(?P<start>\d{1,3}:\d{1,2}(?::\d{1,2})?(?:[.,:]\d{1,3})?)'
        r'\s*(?:-->|->|–>|—>|→)\s*'
        r'(?P<end>\d{1,3}:\d{1,2}(?::\d{1,2})?(?:[.,:]\d{1,3})?)'
    )

    for match in time_pattern.finditer(content):
        try:
            start_ms = parse_time(match.group("start"))
            end_ms = parse_time(match.group("end"))

            if end_ms > start_ms:
                return True

        except ValueError:
            continue

    return False
    

# ==========================================
# STAGE 1: STRETCH FUNCTION
# ==========================================

def stretch_srt_content(content, gap_ms=83, max_stretch_ms=1500):
    """
    Stretches subtitle end times up to 'gap_ms' before the next subtitle starts,
    but will not extend the subtitle by more than 'max_stretch_ms'.
    """
    parsed_subs = parse_srt_to_dicts(content)
    
    for i in range(len(parsed_subs) - 1):
        current_end = parsed_subs[i]['end']
        next_start = parsed_subs[i+1]['start']
        
        # Calculate the actual gap between this subtitle and the next
        gap = next_start - current_end
        
        # If the gap is larger than our target gap (83ms), we can stretch
        if gap > gap_ms:
            # How much time we need to add to perfectly hit the 83ms gap
            potential_stretch = gap - gap_ms
            
            # Restrict the stretch amount to our maximum limit (e.g., 1500ms)
            if max_stretch_ms is not None:
                actual_stretch = min(potential_stretch, max_stretch_ms)
            else:
                actual_stretch = potential_stretch
                
            # Apply the stretch
            parsed_subs[i]['end'] = current_end + actual_stretch
            
    return dicts_to_srt_string(parsed_subs)


# ==========================================
# STAGE 2: MERGE & OPTIMIZE FUNCTION
# ==========================================

def combine_block_group(group):
    """Combines a list of subtitle dictionaries into a single SRT block."""
    if not group: return None
    return {
        'start': group[0]['start'],
        'end': group[-1]['end'],
        'text': " ".join(s['text'] for s in group)
    }
    
def merge_arabic_prefix_blocks(parsed_subs):
    """
    Pre-processes subtitles: If a block has an Arabic character in its first 
    30 characters, it merges it with the previous block regardless of length.
    """
    if not parsed_subs: 
        return []
        
    # Start the list with the first subtitle
    merged_subs = [parsed_subs[0]]
    
    for i in range(1, len(parsed_subs)):
        current_sub = parsed_subs[i]
        
        # Check the first 30 characters for any Arabic unicode character
        if re.search(r'[\u0600-\u06FF]', current_sub['text'][:30]):
            # Merge with the previous block
            merged_subs[-1]['text'] += f" {current_sub['text']}"
            merged_subs[-1]['end'] = current_sub['end']
        else:
            # Keep it as a separate block
            merged_subs.append(current_sub)
            
    return merged_subs
    
    

def optimize_srt_content(content, threshold_ms=800, max_chars=400):
    """Merges and optimizes SRT blocks based on time gaps and character limits."""
    parsed_subs = parse_srt_to_dicts(content)
    if not parsed_subs: return ""

    parsed_subs = merge_arabic_prefix_blocks(parsed_subs)

    # Group into clusters based on time gap
    clusters = []
    current_cluster = [parsed_subs[0]]
    for i in range(1, len(parsed_subs)):
        gap = parsed_subs[i]['start'] - current_cluster[-1]['end']
        if 0 <= gap < threshold_ms:
            current_cluster.append(parsed_subs[i])
        else:
            clusters.append(current_cluster)
            current_cluster = [parsed_subs[i]]
    clusters.append(current_cluster)
    

    # Process clusters with character limits
    final_subs = []
    for cluster in clusters:
        total_text_len = sum(len(s['text']) for s in cluster) + (len(cluster) - 1)
        if total_text_len <= max_chars:
            final_subs.append(combine_block_group(cluster))
        else:
            num_chunks = math.ceil(total_text_len / max_chars)
            blocks_per_chunk = math.ceil(len(cluster) / num_chunks)
            for i in range(0, len(cluster), blocks_per_chunk):
                chunk = cluster[i : i + blocks_per_chunk]
                final_subs.append(combine_block_group(chunk))

    # Post-process small blocks (< 90 characters)
    i = 0
    while i < len(final_subs):
        if len(final_subs[i]['text']) < 90:
            if len(final_subs) == 1:
                i += 1
                continue
            
            merge_prev, merge_next = False, False
            if i == 0: merge_next = True
            elif i == len(final_subs) - 1: merge_prev = True
            else:
                gap_prev = final_subs[i]['start'] - final_subs[i-1]['end']
                gap_next = final_subs[i+1]['start'] - final_subs[i]['end']
                if gap_prev < gap_next: merge_prev = True
                else: merge_next = True
                    
            if merge_prev:
                final_subs[i-1]['text'] += f" [pause] {final_subs[i]['text']}"
                final_subs[i-1]['end'] = final_subs[i]['end']
                final_subs.pop(i)
            elif merge_next:
                final_subs[i+1]['text'] = f"{final_subs[i]['text']} [pause] {final_subs[i+1]['text']}"
                final_subs[i+1]['start'] = final_subs[i]['start']
                final_subs.pop(i)
        else:
            i += 1

    return dicts_to_srt_string(final_subs)


# ==========================================
# STAGE 3: ARABIC/LATIN PARAGRAPH FORMATTER
# ==========================================

def format_to_paragraph(srt_content):
    """
    Removes timestamps, joins subtitle lines, strips punctuation after Arabic,
    and removes the final harakah only from Arabic passages containing more
    than 3 Arabic words.

    Short Arabic expressions of 1–3 words keep their final harakah because
    they may be quoted for pronunciation, grammar, roots, or word analysis.
    """

    pattern = re.compile(
        r'\d+\n'
        r'\d{2}:\d{2}:\d{2},\d{3}'
        r'\s*-->\s*'
        r'\d{2}:\d{2}:\d{2},\d{3}\n'
    )

    blocks = pattern.split(srt_content)

    # Arabic letters, including Persian/extended Arabic letters commonly
    # encountered in copied Qur'anic text.
    arabic_letter_range = r'\u0621-\u064A\u066E-\u06D3'

    # Harakat and common Qur'anic annotation/recitation marks.
    arabic_mark_range = r'\u064B-\u065F\u0670\u06D6-\u06ED'

    # Characters allowed inside one continuous Arabic passage.
    # Includes:
    # - Arabic letters
    # - diacritics and Qur'anic marks
    # - ordinary, non-breaking, and repeated whitespace
    # - Arabic and Latin punctuation that may occur inside the quotation
    arabic_passage_pattern = re.compile(
        rf'''
        (?P<passage>
            [{arabic_letter_range}{arabic_mark_range}]
            (?:
                [{arabic_letter_range}{arabic_mark_range}]
                |
                [\s\u00A0]+
                |
                [،؛؟,:;.!?]
                |
                [﴾﴿]
            )*
            [{arabic_letter_range}{arabic_mark_range}]
        )
        ''',
        re.VERBOSE
    )

    def remove_final_harakah_from_long_passage(match):
        passage = match.group('passage')

        # Count actual Arabic words, ignoring spaces, punctuation,
        # diacritics, and Qur'anic annotation marks.
        arabic_words = re.findall(
            rf'[{arabic_letter_range}]'
            rf'[{arabic_letter_range}{arabic_mark_range}]*',
            passage
        )

        # Preserve one-word and short Arabic expressions.
        if len(arabic_words) <= 3:
            return passage

        # Remove only the mark or marks attached to the final Arabic letter.
        # This works whether Latin text follows directly, after one space,
        # after several spaces, or after a tab.
        return re.sub(
            rf'([{arabic_letter_range}])'
            rf'[{arabic_mark_range}]+$',
            r'\1',
            passage
        )

    cleaned_blocks = []

    for block in blocks:
        text = block.strip()

        if not text:
            continue

        # 1. Convert line breaks to spaces while preserving [pause].
        text = text.replace('\n', ' ')

        # Normalize accidental repeated whitespace without disturbing [pause].
        text = re.sub(r'[ \t\u00A0]+', ' ', text).strip()

        # 2. Strip standard punctuation immediately after Arabic text.
        text = re.sub(
            r'([\u0600-\u06FF])'
            r'[\u200B-\u200F\u202A-\u202E]*'
            r'[.,!?;:،؛؟]+',
            r'\1',
            text
        )

        # 3. Remove the final harakah only from Arabic passages
        # containing more than 3 Arabic words.
        text = arabic_passage_pattern.sub(
            remove_final_harakah_from_long_passage,
            text
        )

        cleaned_blocks.append(text)

    return '\n\n'.join(cleaned_blocks)

def read_input_content(input_path):
    """
    Reads supported input formats and returns plain text.

    Supported:
    - .srt
    - .txt
    - .docx
    """
    extension = os.path.splitext(input_path)[1].lower()
    raw_content = ""

    if extension in {".srt", ".txt"}:
        # utf-8-sig also removes a UTF-8 BOM when present.
        with open(input_path, "r", encoding="utf-8-sig") as file:
            raw_content = file.read()

    elif extension == ".docx":
        document = Document(input_path)
        blocks = []
        for paragraph in document.paragraphs:
            text = paragraph.text.strip()
            if text:
                blocks.append(text)
        raw_content = "\n\n".join(blocks)

    else:
        raise ValueError(
            "Unsupported input format. Please select an SRT, TXT, or DOCX file."
        )

    # GLOBAL HTML TAG STRIPPER: 
    # Automatically wipes out <i>, <b>, <🔴1>, and any other bracketed tags
    raw_content = re.sub(r'<[^>]+>', '', raw_content)

    return raw_content
    
# ==========================================
# MAIN PIPELINE RUNNER
# ==========================================

def run_pipeline():
    root = tk.Tk()
    root.withdraw() # Hide the main tk window
    
    # Ask for Input Files (allows multiple selection)
    input_paths = filedialog.askopenfilenames(
        title="Select Subtitle or Text Files",
        filetypes=[
            (
                "Supported files",
                "*.srt *.txt *.docx"
            ),
            ("SRT subtitle files", "*.srt"),
            ("Text files", "*.txt"),
            ("Word documents", "*.docx"),
            ("All files", "*.*")
        ]
    )
    
    if not input_paths:
        return # User cancelled

    # Process files in order
    sorted_input_paths = sorted(input_paths)
    saved_files = []
    errors = []

    for input_path in sorted_input_paths:
        dir_name, full_file_name = os.path.split(os.path.splitext(input_path)[0])
        
        # If the filename contains an underscore, strip everything before and including the first underscore
        if "_" in full_file_name:
            clean_file_name = re.sub(r'^[^_]*_', '', full_file_name)
        else:
            clean_file_name = full_file_name

        try:
            # Read the original file
            raw_content = read_input_content(input_path)

            if not raw_content.strip():
                raise ValueError(
                    f"The file '{os.path.basename(input_path)}' contains no readable text."
                )

            # Pass the raw text into the optimization pipeline
            merged_srt = optimize_srt_content(raw_content)
            paragraph_txt = format_to_paragraph(merged_srt)
            
            # Get the directory and the actual file name separately to keep the prefix in front of the file name
            out_txt_path = os.path.join(dir_name, f"podcast_{clean_file_name}.txt")
            
            with open(out_txt_path, 'w', encoding='utf-8') as f:
                f.write(paragraph_txt)
                
            saved_files.append(os.path.basename(out_txt_path))

        except Exception as e:
            errors.append(f"{os.path.basename(input_path)}: {e}")

    # Summary report
    if saved_files and not errors:
        success_msg = f"Pipeline completed successfully for {len(saved_files)} file(s)!\n\nFiles saved in their original folders:\n"
        for file in saved_files:
            success_msg += f" • {file}\n"
        messagebox.showinfo("Success", success_msg)
    elif saved_files and errors:
        msg = f"Pipeline partially completed ({len(saved_files)} succeeded, {len(errors)} failed).\n\nSaved:\n"
        for file in saved_files:
            msg += f" • {file}\n"
        msg += "\nErrors:\n"
        for err in errors:
            msg += f" • {err}\n"
        messagebox.showwarning("Completed with Warnings", msg)
    elif errors:
        err_msg = "An error occurred while processing files:\n\n" + "\n".join(errors)
        messagebox.showerror("Error", err_msg)

if __name__ == "__main__":
    run_pipeline()
