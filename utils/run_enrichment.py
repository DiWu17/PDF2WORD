import os
import json
import re
from thefuzz import fuzz
from tqdm import tqdm
from loguru import logger

def enrich_layout_with_font_size(layout_path, analysis_path, output_path):
    """
    Enriches the layout.json file with average font size information from the analysis.json file.
    Uses a pre-built index to significantly speed up matching for long text blocks.
    """
    import json
    import re
    from thefuzz import fuzz
    from tqdm import tqdm

    logger.info("Loading layout file...")
    with open(layout_path, 'r', encoding='utf-8') as f:
        layout_data = json.load(f)

    logger.info("Loading analysis file...")
    with open(analysis_path, 'r', encoding='utf-8') as f:
        analysis_data = json.load(f)

    logger.info("Extracting spans from analysis data...")
    analysis_spans = []
    for page in analysis_data.get('pages', []):
        for block in page.get('blocks', []):
            if 'lines' in block:
                for line in block.get('lines', []):
                    for span in line.get('spans', []):
                        analysis_spans.append({
                            'text': span.get('text', ''),
                            'size': span.get('size', 0)
                        })

    # 1. Build a quick-lookup index to find starting positions of text sequences.
    logger.info("\nBuilding a quick-lookup index for analysis spans...")
    INDEX_NGRAM_SIZE = 7  # Use a 7-character key for the index.
    span_index = {}
    
    # The index maps a short, normalized text snippet (key) to the starting span index (value).
    for i in tqdm(range(len(analysis_spans)), desc="Building Index", leave=False):
        # Create a candidate key by joining a few subsequent spans
        spans_to_join = analysis_spans[i : i + 5]
        raw_text = "".join(s['text'] for s in spans_to_join)
        normalized_text = re.sub(r'\s+', '', raw_text)
        
        if len(normalized_text) >= INDEX_NGRAM_SIZE:
            key = normalized_text[:INDEX_NGRAM_SIZE]
            if key not in span_index:  # Store only the first occurrence to keep the index small
                span_index[key] = i

    logger.info("\nIndex built. Starting block matching process...")

    # 2. Process all blocks.
    for page_info in tqdm(layout_data.get('pdf_info', []), desc="Processing Pages"):
        for block in page_info.get('para_blocks', []):
            if block.get('type') in ['text', 'title']:
                content_parts = []
                if 'lines' in block:
                    for line in block['lines']:
                        for span in line.get('spans', []):
                            content_parts.append(span.get('content', ''))
                
                target_text = "".join(content_parts)
                normalized_target = re.sub(r'\s+', '', target_text)

                if not normalized_target:
                    continue

                best_match_info = {'score': 0, 'avg_size': 0, 'text': ''}
                
                # --- OPTIMIZATION LOGIC ---
                # For long texts, we use the index. For short texts, a full scan is fast enough.
                if len(normalized_target) >= 20:
                    # FAST PATH for long texts
                    search_key = normalized_target[:INDEX_NGRAM_SIZE]
                    start_pos = span_index.get(search_key)

                    if start_pos is not None:
                        # Index hit! We only need to search from this one starting position.
                        for j in range(start_pos, min(start_pos + 50, len(analysis_spans))):
                            current_text_parts = [s['text'] for s in analysis_spans[start_pos:j+1]]
                            normalized_current = re.sub(r'\s+', '', "".join(current_text_parts))
                            
                            similarity = fuzz.partial_ratio(normalized_target, normalized_current)
                            if similarity > 90 and len(normalized_current) > len(best_match_info['text']):
                                sizes = [s['size'] for s in analysis_spans[start_pos:j+1] if s['size'] > 0]
                                if sizes:
                                    avg_size = sum(sizes) / len(sizes)
                                    best_match_info = {'score': similarity, 'avg_size': round(avg_size, 2), 'text': normalized_current}
                            
                            if len(normalized_current) > len(normalized_target) + 10:
                                break
                else:
                    # SLOW PATH for short texts (but it's fast enough)
                    for i in range(len(analysis_spans)):
                        current_text_parts = []
                        for j in range(i, min(i + 50, len(analysis_spans))):
                            current_text_parts.append(analysis_spans[j]['text'])
                            normalized_current = re.sub(r'\s+', '', "".join(current_text_parts))
                            
                            similarity = fuzz.partial_ratio(normalized_target, normalized_current)
                            if similarity > 90 and len(normalized_current) > len(best_match_info['text']):
                                sizes = [s['size'] for s in analysis_spans[i:j+1] if s['size'] > 0]
                                if sizes:
                                    avg_size = sum(sizes) / len(sizes)
                                    best_match_info = {'score': similarity, 'avg_size': round(avg_size, 2), 'text': normalized_current}
                            
                            if len(normalized_current) > len(normalized_target) + 10:
                                break
                        if best_match_info['score'] > 98: # Optimization for short texts: exit if a near-perfect match is found
                            break

                # Update block with the best match found
                if best_match_info['score'] > 90:
                    block['avg_size'] = best_match_info['avg_size']

    logger.info(f"\nSaving enriched layout to {output_path}")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(layout_data, f, ensure_ascii=False, indent=4)
    
    logger.info("Enrichment complete.")

if __name__ == "__main__":
    layout_file = os.path.join('output', 'layout.json')
    analysis_file = os.path.join('output', 'sample_2_analysis.json')
    output_file = os.path.join('output', 'layout_enriched.json')

    if not os.path.exists(layout_file):
        logger.error(f"Error: Layout file not found at {layout_file}")
    elif not os.path.exists(analysis_file):
        logger.error(f"Error: Analysis file not found at {analysis_file}")
    else:
        enrich_layout_with_font_size(layout_file, analysis_file, output_file)
        logger.info(f"Enriched layout saved to {output_file}") 