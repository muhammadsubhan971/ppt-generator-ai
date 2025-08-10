import google.generativeai as genai
import requests
import PyPDF2
import time
import re
import os
from emotion import EnhancedPPTGenerator  


genai.configure(api_key="AIzaSyC4RF-x0XD4Ccq-AAOHS9u4YxscJpZEiBc")  # Replace with your actual API key


def clean_text_lines(text):
    """
    Clean and normalize lines extracted from AI response with improved regex
    """
    cleaned_lines = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        
        # Remove bullet chars or numbering at start (fixed pattern)
        line = re.sub(r'^\s*([•\-\*\d+\.]+)\s*', '', line)
        
        # Remove common prefixes like "Topic:", "Subject:", etc.
        line = re.sub(r'^\s*(Topic|Subject|Chapter|Section|Point):\s*', '', line, flags=re.IGNORECASE)
        
        # Remove trailing punctuation like ':', '.', ';'
        line = re.sub(r'[\:\.\;,]+$', '', line)
        
        # Normalize multiple spaces to single space
        line = re.sub(r'\s+', ' ', line)
        
        # Remove empty parentheses or brackets
        line = re.sub(r'\(\s*\)|\[\s*\]', '', line)
        
        final_line = line.strip()
        if final_line and len(final_line) > 3:  # Only keep meaningful content
            cleaned_lines.append(final_line)
    
    return cleaned_lines


def send_content(message):
    """
    Enhanced content analysis with better, more general prompt
    """
    generation_config = {
        "temperature": 0.3,
        "top_k": 40,
        "top_p": 0.8,
        "max_output_tokens": 500,
        "response_mime_type": "text/plain",
    }
    
    model = genai.GenerativeModel(
        model_name='gemini-2.0-flash-exp',
        generation_config=generation_config,
    )
    
    chat_session = model.start_chat()
    
    # IMPROVED: General prompt that works for any subject
    enhanced_prompt = f"""
    Analyze the following academic content and extract the main topics/concepts covered.

    Content: {message}

    Instructions:
    - Identify the core topics, concepts, theories, or subjects mentioned
    - Make each topic clear and specific (3-10 words)
    - Focus on educational concepts that can be explained in a presentation
    - Include both main topics and important subtopics
    - Avoid administrative details, page numbers, references
    - Don't include generic words like "introduction", "overview", "summary"
    - Each topic should be suitable as a slide title
    - List one topic per line
    - No bullet points, numbers, or special formatting
    """
    
    try:
        response = chat_session.send_message(enhanced_prompt)
        cleaned_topics = clean_text_lines(response.text)
        return cleaned_topics
    except Exception as e:
        print(f"Error in send_content: {e}")
        return []


def engine(topic, model_name, subject_context=""):
    """
    FIXED: Enhanced slide generation with comprehensive content and general prompts
    """
    generation_config = {
        "temperature": 0.4,
        "top_k": 50,
        "top_p": 0.9,
        "max_output_tokens": 800,  # INCREASED for more content
        "response_mime_type": "text/plain",
    }
    
    model = genai.GenerativeModel(
        model_name=model_name,
        generation_config=generation_config,
    )
    
    chat_session = model.start_chat()
    
    # FIXED: Comprehensive, general prompt for any subject
    enhanced_prompt = f"""
    Create detailed slide content for the topic: "{topic}"
    {f"Subject context: {subject_context}" if subject_context else ""}

    Generate 5-8 comprehensive bullet points that thoroughly explain this topic.

    Requirements for each bullet point:
    - 15-35 words per point (detailed explanations, not single lines)
    - Use clear, educational language appropriate for students
    - Include specific examples, definitions, or details where relevant
    - Cover different aspects of the topic (definition, importance, examples, applications)
    - Make each point informative and educational
    - Use complete sentences that provide real value
    - Include technical terms with brief explanations when needed
    - Ensure points build understanding progressively

    Format: Write each point as a complete sentence, one per line, no bullets or numbers.

    Example for "Photosynthesis Process":
    Photosynthesis is the biological process where plants convert sunlight carbon dioxide and water into glucose and oxygen
    The process occurs primarily in chloroplasts which contain chlorophyll pigments that capture light energy effectively
    Light dependent reactions take place in thylakoids where water molecules are split to release oxygen as a byproduct
    Calvin cycle occurs in the stroma where carbon dioxide is fixed into organic molecules using ATP and NADPH
    Plants produce glucose through photosynthesis which serves as their primary energy source for growth and metabolism
    This process is essential for life on Earth as it produces the oxygen that most organisms need for respiration
    Factors affecting photosynthesis rate include light intensity temperature carbon dioxide concentration and water availability
    """
    
    try:
        response = chat_session.send_message(enhanced_prompt)
        cleaned_bullets = clean_text_lines(response.text)
        
        # IMPROVED: Better filtering for quality content
        filtered_bullets = []
        for bullet in cleaned_bullets:
            # Accept bullets with good length (more comprehensive)
            if 50 <= len(bullet) <= 200:  # Longer, more detailed bullets
                # Skip overly generic or vague content
                generic_phrases = ['very important', 'quite useful', 'extremely helpful', 'it is noted that']
                if not any(phrase in bullet.lower() for phrase in generic_phrases):
                    # Ensure bullet has substantial content
                    word_count = len(bullet.split())
                    if word_count >= 8:  # Minimum 8 words for comprehensive content
                        filtered_bullets.append(bullet)
        
        # If we don't have enough quality bullets, take the best available
        if len(filtered_bullets) < 3:
            # Fallback: take longer bullets even if not perfect
            backup_bullets = [b for b in cleaned_bullets if len(b.split()) >= 5]
            filtered_bullets = backup_bullets[:6]
        
        return filtered_bullets[:8]  # Return up to 8 comprehensive bullets
        
    except Exception as e:
        print(f"Error in engine for topic {topic}: {e}")
        return [f"Detailed information and key concepts related to {topic} will be covered in this section"]


def detect_subject_area(topics_sample):
    """
    Detect the general subject area from topics to provide better context
    """
    if not topics_sample:
        return ""
    
    # Combine first few topics to analyze
    combined_text = " ".join(topics_sample[:5]).lower()
    
    # Subject area detection patterns
    subject_patterns = {
        "Computer Science": ["network", "algorithm", "programming", "database", "software", "computer", "data structure", "coding"],
        "Biology": ["cell", "organism", "dna", "protein", "evolution", "photosynthesis", "genetics", "anatomy"],
        "Chemistry": ["molecule", "atom", "reaction", "compound", "element", "bond", "acid", "base"],
        "Physics": ["force", "energy", "wave", "particle", "quantum", "motion", "electricity", "magnetism"],
        "Mathematics": ["equation", "theorem", "calculus", "algebra", "geometry", "probability", "statistics"],
        "History": ["war", "empire", "revolution", "century", "ancient", "medieval", "dynasty", "civilization"],
        "Economics": ["market", "economy", "trade", "finance", "money", "supply", "demand", "economic"],
        "Psychology": ["behavior", "cognitive", "mental", "brain", "psychology", "social", "personality"]
    }
    
    for subject, keywords in subject_patterns.items():
        if any(keyword in combined_text for keyword in keywords):
            return subject
    
    return "General Academic"


def get_available_templates():
    """
    Get list of available template files from templates folder
    """
    templates_folder = "templates"
    
    # Create templates folder if it doesn't exist
    if not os.path.exists(templates_folder):
        os.makedirs(templates_folder)
        print(f"Created {templates_folder} folder. Please add your .pptx template files there.")
        return []
    
    # Get all .pptx files from templates folder
    template_files = []
    try:
        for file in os.listdir(templates_folder):
            if file.lower().endswith('.pptx') and not file.startswith('~'):
                template_files.append(file)
    except Exception as e:
        print(f"Error reading templates folder: {e}")
        return []
    
    return sorted(template_files)


def note(pdf_file, template_name="default"):
    """
    IMPROVED: Main processing function with template selection instead of theme
    """
    topics = []
    
    try:
        with open(pdf_file, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            total_pages = len(reader.pages)
            print(f"Processing {total_pages} pages...")
            
            for i in range(total_pages):
                try:
                    page = reader.pages[i]
                    text1 = page.extract_text()
                    
                    if text1.strip():  # Only process non-empty pages
                        page_topics = send_content(text1)
                        topics.extend(page_topics)
                        print(f"Page {i+1}/{total_pages}: Found {len(page_topics)} topics")
                    
                    time.sleep(2)  # Rate limiting
                    
                except Exception as e:
                    print(f"Error processing page {i+1}: {e}")
                    continue
                    
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return

    if not topics:
        print("No topics found in PDF")
        return

    # Enhanced deduplication with better similarity detection
    def normalize_for_comparison(text):
        text = text.lower()
        text = re.sub(r'[^\w\s]', '', text)  # Remove punctuation
        text = re.sub(r'\s+', ' ', text).strip()
        return text

    def are_similar(text1, text2, threshold=0.7):
        """Check if two topics are similar based on word overlap"""
        words1 = set(normalize_for_comparison(text1).split())
        words2 = set(normalize_for_comparison(text2).split())
        
        if not words1 or not words2:
            return False
            
        intersection = len(words1.intersection(words2))
        union = len(words1.union(words2))
        
        return (intersection / union) >= threshold

    # IMPROVED: Better deduplication
    filtered_topics = []
    for topic in topics:
        norm_topic = normalize_for_comparison(topic)
        if len(norm_topic) > 5:  # Minimum length check
            # Check for similarity with existing topics
            is_duplicate = False
            for existing in filtered_topics:
                if are_similar(topic, existing):
                    is_duplicate = True
                    break
            
            if not is_duplicate:
                filtered_topics.append(topic)

    print(f"Found {len(filtered_topics)} unique topics after deduplication")

    # ADDED: Detect subject area for better context
    subject_context = detect_subject_area(filtered_topics)
    print(f"Detected subject area: {subject_context}")

    # Model rotation for variety
    models = [
        "gemini-2.0-flash",
        "gemini-1.5-flash", 
        "gemini-1.5-pro",  # Fixed model name
        "gemini-2.0-flash"
    ]
    
    sections = []
    
    for p, topic in enumerate(filtered_topics):
        model_name = models[p % len(models)]
        print(f"Processing topic {p+1}/{len(filtered_topics)}: {topic}")
        
        try:
            # IMPROVED: Pass subject context for better content generation
            slide_bullets = engine(topic, model_name, subject_context)
            
            if slide_bullets and len(slide_bullets) >= 3:  # Ensure minimum quality content
                sections.append({
                    "title": topic,
                    "content": slide_bullets
                })
                print(f"  Generated {len(slide_bullets)} comprehensive bullet points")
            else:
                print(f"  Insufficient quality content for: {topic}")
                
        except Exception as e:
            print(f"  Error generating content for {topic}: {e}")
            continue
            
        time.sleep(3)  # Rate limiting

    if not sections:
        print("No sections generated")
        return

    # IMPROVED: Dynamic presentation setup with template selection
    presentation_title = f"{subject_context} Presentation" if subject_context != "General Academic" else "Academic Presentation"
    
    content_dict = {
        "title": presentation_title,
        "subtitle": "Generated from PDF Analysis",
        "target_slides": min(25, len(sections) + 4),  # More realistic slide count
        "sections": sections,
        "call_to_action": "Questions and Discussion"
    }

    try:
        # CHANGED: Use template instead of theme
        ppt_gen = EnhancedPPTGenerator(template_name="green")
        ppt, actual_slide_count = ppt_gen.generate_from_content(content_dict)
        
        # IMPROVED: Better filename with subject and template name
        safe_subject = re.sub(r'[^\w\s-]', '', subject_context).replace(' ', '_')
        safe_template = re.sub(r'[^\w\s-]', '', template_name.replace('.pptx', '')).replace(' ', '_')
        output_file = ppt_gen.save(f"{safe_subject.lower()}_{safe_template}_presentation.pptx")
        
        print(f"\n✅ SUCCESS!")
        print(f"📊 Presentation saved as: {output_file}")
        print(f"📄 Total slides: {actual_slide_count}")
        print(f"📚 Topics covered: {len(sections)}")
        print(f"🎯 Subject area: {subject_context}")
        print(f"🎨 Template used: {template_name}")
        
    except Exception as e:
        print(f"❌ Error creating presentation: {e}")


if __name__ == "__main__":
    pdf_path = r"E:\os mod\Information security.pdf"  # Change this to your PDF path
    
    # Get available templates
    available_templates = get_available_templates()
    
    if available_templates:
        print("Available templates:")
        for i, template in enumerate(available_templates, 1):
            print(f"  {i}. {template}")
        
        # You can select template here or pass it as parameter
        selected_template = available_templates[0] if available_templates else "default"
        print(f"Using template: {selected_template}")
    else:
        print("No templates found in templates folder. Using default template.")
        selected_template = "default"
    
    # Check if file exists
    if os.path.exists(pdf_path):
        print(f"Starting processing of: {pdf_path}")
        note(pdf_path, selected_template)
    else:
        print(f"❌ PDF file not found: {pdf_path}")
        print("Please update the pdf_path variable with the correct path to your PDF file.")