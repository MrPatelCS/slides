import io
import json
import os
import urllib.request
from google import genai
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
import streamlit as st
import soundfile as sf
import numpy as np
from kokoro_onnx import Kokoro
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Page configuration
st.set_page_config(page_title="Teacher Toolkit", layout="centered")

# --- GITHUB RELEASE MODEL LOADER (NO LOCAL FILES) ---
@st.cache_resource
def load_kokoro_files():
    """Downloads Kokoro model files exclusively from GitHub Releases into /tmp/."""
    MODEL_URL = "https://github.com/MrPatelCS/slides/releases/download/v1.0.0/kokoro-v1.0.onnx"
    VOICES_URL = "https://github.com/MrPatelCS/slides/releases/download/v1.0.0/voices-v1.0.bin"
    
    model_dir = "/tmp/models"
    os.makedirs(model_dir, exist_ok=True)
    
    model_path = os.path.join(model_dir, "kokoro-v1.0.onnx")
    voices_path = os.path.join(model_dir, "voices-v1.0.bin")
    
    if not os.path.exists(model_path):
        with st.spinner("Downloading Kokoro model weights from GitHub (approx. 310MB)... Please wait."):
            urllib.request.urlretrieve(MODEL_URL, model_path)
            
    if not os.path.exists(voices_path):
        with st.spinner("Downloading Kokoro voice files from GitHub..."):
            urllib.request.urlretrieve(VOICES_URL, voices_path)
            
    return model_path, voices_path

# Load the files into variables for use across tabs
model_path, voices_path = load_kokoro_files()

st.title("Teacher Toolkit")

# --- SIDEBAR: API KEY ---
default_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        default_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

if st.session_state.get("gemini_api_key_v5") != default_key and default_key:
    st.session_state["gemini_api_key_v5"] = default_key

api_key_input = st.sidebar.text_input(
    "Enter Gemini API Key:", type="password", key="gemini_api_key_v5"
)

# --- NAVIGATION TABS ---
tab_slides, tab_tts, tab_dialogue, tab_workflow = st.tabs(["Slide Maker", "Text-to-Speech", "Dialogue Maker", "Workflow Planner"])

# =========================================================================
# TAB 1: SLIDE GENERATOR
# =========================================================================
with tab_slides:
    st.subheader("Slide Generator")

    selected_standard = st.selectbox(
        "Select Standards to Follow:",
        [
            "New York Common Core",
            "California Common Core",
            "Cambridge",
            "UK National Curriculum"
        ],
        index=0,
        key="slides_standard",
        help="This standard will be applied to guide all generated content across tools."
    )

    def get_balanced_font_size(text):
        length = len(text)
        if length < 60:
            return Pt(28)
        elif length < 150:
            return Pt(22)
        elif length < 350:
            return Pt(18)
        else:
            return Pt(14)

    def create_dual_coding_icon(phase_name):
        img = Image.new("RGBA", (180, 180), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle([10, 10, 170, 170], radius=40, fill=(0, 128, 128, 255))
        
        teal_dark = (0, 90, 90, 255)
        white = (255, 255, 255, 255)
        accent_yellow = (255, 215, 0, 255)
        phase = phase_name.lower()

        if "engage" in phase:
            draw.ellipse([50, 50, 130, 130], outline=white, width=10)
            draw.ellipse([75, 75, 105, 105], fill=accent_yellow)
        elif "explore" in phase:
            draw.ellipse([50, 50, 110, 110], outline=white, width=10)
            draw.line([95, 95, 135, 135], fill=white, width=10)
        elif "explain" in phase:
            draw.ellipse([65, 45, 115, 95], fill=accent_yellow)
            draw.rectangle([78, 95, 102, 120], fill=white)
            draw.rectangle([75, 125, 105, 135], fill=teal_dark)
        elif "elaborate" in phase:
            draw.polygon([(90, 60), (140, 50), (140, 120), (90, 130)], fill=white)
            draw.polygon([(90, 60), (40, 50), (40, 120), (90, 130)], fill=white)
            draw.line([90, 60, 90, 130], fill=teal_dark, width=4)
        else:
            draw.rounded_rectangle([50, 45, 130, 140], radius=15, outline=white, width=8)
            draw.line([68, 85, 85, 105], fill=accent_yellow, width=8)
            draw.line([85, 105, 115, 70], fill=accent_yellow, width=8)

        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format="PNG")
        img_byte_arr.seek(0)
        return img_byte_arr

    lesson_type = "general"

    generation_mode = st.radio(
        "Choose Input Mode:",
        ["Text Only", "Files + Instructions"],
        horizontal=True,
    )

    with st.form("slide_form"):
        if generation_mode == "Text Only":
            lesson_input = st.text_area("Paste Lesson Plan & Text Here:", height=250)
            uploaded_files = []
        else:
            uploaded_files = st.file_uploader(
                "Upload Documents / Screenshots (PDF, PNG, JPG - Max 5 pages recommended):",
                type=["png", "jpg", "jpeg", "pdf"],
                accept_multiple_files=True,
            )
            lesson_input = st.text_area(
                "Enter Additional Instructions / Lesson Plan (Optional):", height=120
            )

        with st.expander("⚙️ Customize Slide Section Prompts (Advanced)", expanded=False):
            st.write("Modify the generation guidelines for each slide component below:")
            p_do_now = st.text_input("do_now", "Student-facing warm-up task, prediction question, or initial engagement prompt.")
            p_learning_focus = st.text_input("learning_focus", "Must start explicitly with 'I can [skill] by [method].'")
            p_activity_title = st.text_input("activity_title", "Student-facing title of the main activity.")
            p_new_learning = st.text_input("new_learning", "Student-facing informational text or core conceptual bullet points summarizing the reading/background.")
            p_vocabulary = st.text_input("vocabulary", "A clean list of key vocabulary words and definitions relevant to the text.")
            p_modeling = st.text_input("modeling", "A clear model example or guiding question showing students what a strong response looks like.")
            p_guided_practice = st.text_input("guided_practice", "Discussion prompts or collaborative questions for students to analyze together in pairs.")
            p_independent_practice = st.text_input("independent_practice", "The actual complete practice task, writing prompt, or exercise for students to solve independently.")
            p_exit_task = st.text_input("exit_task", "A final student-facing reflection or check-for-understanding prompt.")

        submitted = st.form_submit_button("Generate Slides")

    if submitted:
        if not api_key_input.strip():
            st.error("Please enter your API key in the sidebar first!")
        elif generation_mode == "Text Only" and not lesson_input.strip():
            st.warning("Please paste some lesson text first!")
        elif generation_mode == "Files + Instructions" and not uploaded_files:
            st.warning("Please upload at least one file (PDF or image) first!")
        else:
            with st.spinner(f"Generating slides aligned with {selected_standard}..."):
                try:
                    client = genai.Client(api_key=api_key_input.strip())
                    base_prompt = f"""
                        You are an expert curriculum designer creating a projector-ready, student-facing slide deck optimized for a **{lesson_type}** lesson format.
                        All content, competencies, and terminology must strictly align with the **{selected_standard}** framework.
                        Return your response strictly as a valid JSON object with the following exact keys:
                        - "subject": The subject name, e.g. English Language Arts or Computer Science
                        - "topic": The complete unit and lesson topic name
                        - "do_now": {p_do_now}
                        - "learning_focus": {p_learning_focus}
                        - "activity_title": {p_activity_title}
                        - "new_learning": {p_new_learning}
                        - "vocabulary": {p_vocabulary}
                        - "modeling": {p_modeling}
                        - "guided_practice": {p_guided_practice}
                        - "independent_practice": {p_independent_practice}
                        - "exit_task": {p_exit_task}
                        CRITICAL RULES:
                        1. STRICTLY STUDENT-FACING: Never write teacher meta-instructions. Write everything directly as text for students.
                        2. NO EMPTY VALUES: Fill out every field with rich, complete content tailored to a {lesson_type} pedagogical approach.
                        """

                    if generation_mode == "Text Only":
                        prompt = base_prompt + f"\nTake the following lesson plan and text and map it completely into the required structure:\n{lesson_input}"
                        api_contents = prompt
                    else:
                        prompt = base_prompt + f"\nAnalyze the provided files and incorporate the following instructions to map everything completely into the required structure:\nInstructions: {lesson_input}"
                        api_contents = [prompt]
                        for uploaded_file in uploaded_files:
                            file_extension = uploaded_file.name.split(".")[-1].lower()
                            if file_extension in ["png", "jpg", "jpeg"]:
                                api_contents.append(Image.open(uploaded_file))
                            elif file_extension == "pdf":
                                api_contents.append({
                                    "mime_type": "application/pdf",
                                    "data": uploaded_file.getvalue(),
                                })

                    response = client.models.generate_content(
                        model="gemini-3.1-flash-lite", contents=api_contents
                    )

                    if not response or not response.text:
                        st.error("Empty response. Please try again.")
                    else:
                        raw_text = response.text.strip()
                        if raw_text.startswith("```json"):
                            raw_text = raw_text[7:]
                        if raw_text.endswith("```"):
                            raw_text = raw_text[:-3]
                        raw_text = raw_text.strip()

                        deck_data = json.loads(raw_text)

                        st.subheader("Output Preview:")
                        st.json(deck_data)

                        template_path = "template.pptx"

                        if not os.path.exists(template_path):
                            st.error(f"Template not found at path: {template_path}. Please include 'template.pptx' in your repository.")
                        else:
                            prs = Presentation(template_path)

                            def format_slide_content(val):
                                if isinstance(val, list):
                                    formatted_lines = []
                                    for item in val:
                                        if isinstance(item, dict):
                                            for k, v in item.items():
                                                formatted_lines.append(f"• **{k}**: {v}")
                                        else:
                                            formatted_lines.append(f"• {item}")
                                    return "\n".join(formatted_lines)
                                elif isinstance(val, dict):
                                    formatted_lines = []
                                    for k, v in val.items():
                                        formatted_lines.append(f"• **{k}**: {v}")
                                    return "\n".join(formatted_lines)
                                return str(val)

                            slide_mapping = {
                                1: ("Engage", format_slide_content(deck_data.get("do_now", ""))),
                                2: ("Engage", format_slide_content(deck_data.get("learning_focus", ""))),
                                3: ("Explore", format_slide_content(deck_data.get("activity_title", ""))),
                                4: ("Explain", format_slide_content(deck_data.get("new_learning", ""))),
                                5: ("Explain", format_slide_content(deck_data.get("vocabulary", ""))),
                                6: ("Explain", format_slide_content(deck_data.get("modeling", ""))),
                                7: ("Elaborate", format_slide_content(deck_data.get("guided_practice", ""))),
                                8: ("Elaborate", format_slide_content(deck_data.get("independent_practice", ""))),
                                9: ("Evaluate", format_slide_content(deck_data.get("exit_task", ""))),
                            }

                            for s_idx, (phase_name, text_val) in slide_mapping.items():
                                if len(prs.slides) > s_idx:
                                    slide = prs.slides[s_idx]
                                    content_box = slide.shapes.add_textbox(Inches(1.2), Inches(1.6), Inches(11.0), Inches(5.0))
                                    target_tf = content_box.text_frame
                                    target_tf.word_wrap = True
                                    target_tf.margin_left = Inches(0.2)
                                    target_tf.margin_right = Inches(0.2)
                                    target_tf.margin_top = Inches(0.1)
                                    target_tf.margin_bottom = Inches(0.1)

                                    lines = str(text_val).strip().split("\n")
                                    font_size = get_balanced_font_size(text_val)

                                    for i, line in enumerate(lines):
                                        if i == 0:
                                            p = target_tf.paragraphs[0]
                                        else:
                                            p = target_tf.add_paragraph()
                                        p.text = line.strip()
                                        p.font.name = "Aptos"
                                        p.font.size = font_size
                                        p.font.color.rgb = RGBColor(40, 40, 40)
                                        p.space_after = Pt(4)

                                    if s_idx >= 1:
                                        icon_stream = create_dual_coding_icon(phase_name)
                                        slide.shapes.add_picture(
                                            icon_stream,
                                            left=Inches(0.8),
                                            top=Inches(6.4),
                                            width=Inches(0.9),
                                            height=Inches(0.9),
                                        )

                            output_path = "Lesson.pptx"
                            prs.save(output_path)

                            st.success(f"Presentation generated successfully under {selected_standard}")
                            with open(output_path, "rb") as f:
                                st.download_button(
                                    label="Download Presentation (.pptx)",
                                    data=f,
                                    file_name="Lesson.pptx",
                                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                )

                except Exception as e:
                    st.error(f"Error processing generation: {e}")

# =========================================================================
# TAB 2: TEXT-TO-SPEECH (Kokoro Cloud Workspace)
# =========================================================================
with tab_tts:
    st.subheader("Audio Generator")
    
    tts_text = st.text_area("Enter text for audio generation:", "Welcome to your lesson. Listen carefully and repeat after me.", key="kokoro_text_input")
    
    tts_voice_choice = st.selectbox(
        "Select Accent / Voice Profile", 
        [
            "British Female - Emma (bf_emma)", 
            "British Female - Isabella (bf_isabella)", 
            "British Male - Daniel (bm_daniel)", 
            "British Male - Fable (bm_fable)", 
            "American Female - Sarah (af_sarah)", 
            "American Male - Michael (am_michael)"
        ],
        key="kokoro_voice_select"
    )
    
    tts_voice_id = tts_voice_choice.split("(")[1].strip(")")
    
    if st.button("Generate Audio File", key="kokoro_generate_btn"):
        if not tts_text.strip():
            st.warning("Please enter text to synthesize.")
        else:
            with st.spinner("Synthesizing voice audio..."):
                try:
                    kokoro = Kokoro(model_path, voices_path)
                    samples, sample_rate = kokoro.create(
                        tts_text, 
                        voice=tts_voice_id, 
                        speed=1.0, 
                        lang="en-us"
                    )
                    
                    output_audio_path = "lesson_output.wav"
                    sf.write(output_audio_path, samples, sample_rate)
                    
                    st.success("Audio successfully synthesized!")
                    st.audio(output_audio_path)
                    
                    col_wav, col_mp3 = st.columns(2)
                    with col_wav:
                        with open(output_audio_path, "rb") as f:
                            st.download_button(
                                label="Download Audio (.wav)", 
                                data=f, 
                                file_name="lesson_audio.wav",
                                mime="audio/wav",
                                key="kokoro_download_btn_wav"
                            )
                    with col_mp3:
                        with open(output_audio_path, "rb") as f_mp3:
                            st.download_button(
                                label="Download Audio (.wav as MP3)", 
                                data=f_mp3, 
                                file_name="lesson_audio.mp3",
                                mime="audio/mpeg",
                                key="kokoro_download_btn_mp3"
                            )
                            
                except Exception as e:
                    st.error(f"TTS Generation Error: {e}")

# =========================================================================
# TAB 3: DIALOGUE MAKER & WORKSHEET CREATOR
# =========================================================================
with tab_dialogue:
    st.subheader("Dialogue Maker")
    
    selected_standard_dialogue = st.selectbox(
        "Select Standards to Follow:",
        [
            "New York Common Core",
            "California Common Core",
            "Cambridge",
            "UK National Curriculum"
        ],
        index=0,
        key="dialogue_standard",
    )
    
    st.write(f"Generate interactive multi-speaker conversation transcripts aligned with **{selected_standard_dialogue}**, custom voice pairings, and grade-aligned worksheets.")

    with st.form("dialogue_form"):
        dialogue_topic = st.text_input(
            "Dialogue Topic / Scenario:", 
            placeholder="e.g., Two experts discussing climate change solutions for a grade 9 lesson"
        )
        
        col_chars, col_len = st.columns(2)
        with col_chars:
            num_speakers = st.slider("Number of Speakers:", min_value=2, max_value=3, value=2)
            speaker_names_input = st.text_input(
                "Speaker Names (comma-separated):", 
                value="John, Alice",
                help="Enter names matching the speakers (e.g. John, Alice)"
            )
        with col_len:
            target_duration = st.selectbox(
                "Target Approximate Length:",
                [
                    "Short (~1-2 mins / 200 words)", 
                    "Medium (~3-5 mins / 500 words)",
                    "Large (~6-8 mins / 800+ words)"
                ],
                index=0
            )
            dialogue_tone = st.selectbox(
                "Dialogue Tone / Level:",
                ["Podcast / Engaging Deep-Dive", "Casual / Conversational", "Academic Discussion"]
            )

        dialogue_instructions = st.text_area(
            "Specific Vocabulary or Key Points to Include (Optional):",
            placeholder="e.g., Focus on carbon footprints, renewable energy, and active listening questions."
        )

        dialogue_submitted = st.form_submit_button("Generate Dialogue Script")

    if dialogue_submitted:
        if not api_key_input.strip():
            st.error("Please enter your Gemini API Key in the sidebar first!")
        elif not dialogue_topic.strip():
            st.warning("Please enter a dialogue topic or scenario first!")
        else:
            with st.spinner("Generating dialogue overview & script..."):
                try:
                    client = genai.Client(api_key=api_key_input.strip())
                    
                    dialogue_prompt = f"""
                    You are an expert audio curriculum producer and scriptwriter. Create a realistic multi-person podcast or conversational dialogue transcript based on the following criteria:
                    - Topic: {dialogue_topic}
                    - Standards Framework: {selected_standard_dialogue}
                    - Speakers: {speaker_names_input} ({num_speakers} speakers total)
                    - Target Length: {target_duration}
                    - Tone / Level: {dialogue_tone}
                    - Special Instructions / Key Points: {dialogue_instructions}

                    Return your response strictly as a valid JSON object with the exact following structure:
                    {{
                        "title": "A catchy podcast/dialogue title",
                        "setting": "Brief description of the audio overview context",
                        "transcript": [
                            {{
                                "speaker": "Name of Speaker",
                                "line": "What they say in the conversation"
                            }}
                        ]
                    }}
                    """

                    response = client.models.generate_content(
                        model="gemini-3.1-flash-lite", 
                        contents=dialogue_prompt
                    )

                    if not response or not response.text:
                        st.error("Empty response received. Please try again.")
                    else:
                        raw_json = response.text.strip()
                        if raw_json.startswith("```json"):
                            raw_json = raw_json[7:]
                        if raw_json.endswith("```"):
                            raw_json = raw_json[:-3]
                        raw_json = raw_json.strip()

                        dialogue_data = json.loads(raw_json)

                        st.session_state["last_dialogue_data"] = dialogue_data
                        st.session_state["last_dialogue_level"] = dialogue_tone
                        st.session_state["last_dialogue_topic"] = dialogue_topic
                        st.session_state["last_dialogue_standard"] = selected_standard_dialogue
                        
                        extracted_speakers = []
                        for turn in dialogue_data.get("transcript", []):
                            s_name = turn.get("speaker", "").strip()
                            if s_name and s_name not in extracted_speakers:
                                extracted_speakers.append(s_name)
                        st.session_state["active_speakers"] = extracted_speakers

                        st.success("Dialogue script generated successfully!")

                except Exception as e:
                    st.error(f"Error generating dialogue: {e}")

    if "last_dialogue_data" in st.session_state:
        dialogue_data = st.session_state["last_dialogue_data"]
        curr_std = st.session_state.get("last_dialogue_standard", "Standard")
        
        st.markdown(f"### 🎙️ {dialogue_data.get('title', 'Audio Overview')}")
        st.caption(f"**Context:** {dialogue_data.get('setting', 'N/A')} | **Standard:** {curr_std}")
        st.divider()

        transcript_lines = dialogue_data.get("transcript", [])
        
        full_script_text = f"Title: {dialogue_data.get('title')}\nSetting: {dialogue_data.get('setting')}\nStandard: {curr_std}\n\n"
        for turn in transcript_lines:
            speaker = turn.get("speaker", "Unknown")
            line = turn.get("line", "")
            full_script_text += f"{speaker}: {line}\n\n"
            st.markdown(f"**🗣️ {speaker}:** {line}")

        st.divider()
        st.markdown("### 🎛️ Assign Individual Speaker Voices")
        st.write("Choose any combination of British or American male/female voice profiles for each speaker in your transcript:")

        available_voice_options = {
            "British Female - Emma (bf_emma)": "bf_emma",
            "British Female - Isabella (bf_isabella)": "bf_isabella",
            "British Male - Daniel (bm_daniel)": "bm_daniel",
            "British Male - Fable (bm_fable)": "bm_fable",
            "American Female - Sarah (af_sarah)": "af_sarah",
            "American Male - Michael (am_michael)": "am_michael"
        }

        speaker_voice_mapping = {}
        active_speakers = st.session_state.get("active_speakers", [])
        
        voice_cols = st.columns(len(active_speakers) if len(active_speakers) > 0 else 1)
        for idx, spk in enumerate(active_speakers):
            with voice_cols[idx % len(voice_cols)]:
                chosen_label = st.selectbox(
                    f"Voice for **{spk}**",
                    list(available_voice_options.keys()),
                    index=idx % len(available_voice_options),
                    key=f"voice_select_{spk}"
                )
                speaker_voice_mapping[spk] = available_voice_options[chosen_label]

        if st.button("Synthesize & Play Audio Overview", key="synthesize_audio_btn"):
            with st.spinner("Synthesizing multi-speaker audio..."):
                try:
                    kokoro = Kokoro(model_path, voices_path)
                    combined_samples = []
                    sample_rate = 24000
                    pause_duration = int(sample_rate * 0.4)
                    silence = np.zeros(pause_duration, dtype=np.float32)

                    default_fallback_voice = "bf_emma"

                    for idx, turn in enumerate(transcript_lines):
                        spk = turn.get("speaker", "").strip()
                        voice_id = speaker_voice_mapping.get(spk, default_fallback_voice)
                        
                        line = turn.get("line", "").strip()
                        if not line:
                            continue
                            
                        samples, sample_rate = kokoro.create(
                            line, 
                            voice=voice_id, 
                            speed=1.0, 
                            lang="en-us"
                        )
                        combined_samples.extend(samples)
                        combined_samples.extend(silence)

                    if combined_samples:
                        audio_output_path = "dialogue_overview.wav"
                        sf.write(audio_output_path, np.array(combined_samples), sample_rate)
                        
                        st.markdown("### 🎧 Listen to Audio Overview")
                        st.audio(audio_output_path)
                        
                        with open(audio_output_path, "rb") as aud_file:
                            st.download_button(
                                label="Download Audio Overview (.wav)",
                                data=aud_file,
                                file_name="audio_overview.wav",
                                mime="audio/wav",
                                key="download_gemini_audio"
                            )
                except Exception as audio_err:
                    st.error(f"Audio generation notice: {audio_err}.")

        st.download_button(
            label="Download Dialogue Script (.txt)",
            data=full_script_text,
            file_name="classroom_dialogue_script.txt",
            mime="text/plain",
            key="download_dialogue_txt"
        )

        st.divider()

        st.subheader("📝 Worksheet Creator")
        st.write(f"Generate 5 comprehension questions and model answers aligned with **{curr_std}** based on the generated dialogue script.")

        if st.button("Generate 5-Question Worksheet", key="generate_worksheet_btn"):
            with st.spinner("Generating worksheet questions and answers..."):
                try:
                    client = genai.Client(api_key=api_key_input.strip())
                    worksheet_prompt = f"""
                    You are an expert educator adhering to the **{curr_std}** framework. Based on the following dialogue transcript, create a 5-question student worksheet including comprehension questions and a corresponding teacher answer key.
                    
                    Topic: {st.session_state.get('last_dialogue_topic', '')}
                    Level: {st.session_state.get('last_dialogue_level', '')}
                    Standards Framework: {curr_std}
                    
                    Transcript:
                    {full_script_text}
                    
                    Return your response in clean Markdown format with:
                    1. Worksheet Title & Instructions
                    2. 5 Student Questions (mix of short-answer and critical thinking matching {curr_std})
                    3. A separate Teacher Answer Key section with model responses.
                    """

                    ws_response = client.models.generate_content(
                        model="gemini-3.1-flash-lite",
                        contents=worksheet_prompt
                    )

                    if ws_response and ws_response.text:
                        st.markdown("### 📄 Generated Student Worksheet & Answer Key")
                        st.markdown(ws_response.text)
                        
                        st.download_button(
                            label="Download Worksheet (.txt)",
                            data=ws_response.text,
                            file_name="lesson_worksheet.txt",
                            mime="text/plain",
                            key="download_worksheet_txt"
                        )
                    else:
                        st.error("Failed to generate worksheet. Please try again.")
                except Exception as ws_err:
                    st.error(f"Worksheet generation error: {ws_err}")

# =========================================================================
# TAB 4: WORKFLOW PLANNER (Curriculum & Multi-Session Splitter)
# =========================================================================
with tab_workflow:
    st.subheader("Workflow & Curriculum Planner")
    st.markdown("Design multi-session units with distributed 5-E Framework pacing tailored to your exact lesson count, outputting directly into a cell-by-cell template.")

    with st.form("workflow_form"):
        st.markdown("### Unit Parameters")

        framework_option = st.selectbox(
            "Curriculum Framework",
            [
                "New York Common Core",
                "California State Standards",
                "NGSS",
                "UK National Curriculum",
                "IB PYP/MYP",
            ],
            key="wf_framework"
        )

        col_input1, col_input2 = st.columns(2)

        with col_input1:
            subject_choice = st.selectbox(
                "Subject",
                [
                    "Computer Science",
                    "English Language Arts",
                    "Mathematics",
                    "Science",
                    "History / Social Studies",
                    "Art / Technology",
                    "Other (Custom)",
                ],
            )
            if subject_choice == "Other (Custom)":
                subject_custom = st.text_input("Enter Custom Subject")
                final_subject = subject_custom if subject_custom else "General Subject"
            else:
                final_subject = subject_choice

        with col_input2:
            grade_level = st.selectbox(
                "Grade Level",
                [f"Grade {i}" for i in range(1, 13)],
                index=7,
            )

        col_input3, col_input4 = st.columns(2)
        with col_input3:
            num_sessions = st.slider(
                "Number of Sessions", min_value=1, max_value=10, value=2
            )
        with col_input4:
            session_duration = st.selectbox(
                "Session Duration", ["40 Minutes", "45 Minutes", "60 Minutes"]
            )

        topic_input = st.text_input(
            "Topic / Core Concept",
            placeholder="e.g., while loops in python, fractions, context clues",
        )
        custom_notes = st.text_area(
            "Optional Context / Custom Notes",
            placeholder="Add any specific textbook references, or extra requirements.",
        )

        wf_submitted = st.form_submit_button("Generate Multi-Session Unit Plan", type="primary")

    if wf_submitted:
        if not api_key_input.strip():
            st.error("Please enter your Gemini API Key in the sidebar first!")
        elif not topic_input.strip():
            st.warning("Please provide a topic or core concept.")
        else:
            with st.spinner("Generating structured cell-by-cell unit plan..."):
                try:
                    client = genai.Client(api_key=api_key_input.strip())

                    prompt = f"""
                    You are an expert curriculum designer and teacher. Create a multi-session unit plan for {final_subject}, {grade_level}, covering the topic: "{topic_input}".
                    The curriculum framework to follow is: {framework_option}.
                    The unit must consist of exactly {num_sessions} session(s). Each session duration is {session_duration}.
                    Additional notes: {custom_notes}.

                    For EACH session, provide the output strictly as a JSON list of objects matching this exact structure:
                    [
                      {{
                        "session_number": 1,
                        "total_sessions": {num_sessions},
                        "standard": "Relevant standard code and description according to {framework_option}",
                        "learning_focus": "Clear 'I can...' statement and learning objective",
                        "engage_time": "5 min",
                        "engage_desc": "Hook description and opening question",
                        "explore_time": "7 min",
                        "explore_desc": "Hands-on exploration or investigation activity",
                        "explain_time": "10 min",
                        "explain_desc": "Direct instruction or concept deep-dive",
                        "elaborate_time": "12 min",
                        "elaborate_desc": "Application or coding/practice challenge",
                        "evaluate_time": "6 min",
                        "evaluate_desc": "Exit ticket or formative assessment task",
                        "textbook": "Textbook chapter or reference info",
                        "resources": "Required materials, videos, or worksheets",
                        "evidence_exit_task": "Formative evidence collected / Exit ticket summary",
                        "support_extension": "Support scaffolding + Extension challenge",
                        "adjustment_next_step": "Notes for differentiation or pacing changes"
                      }}
                    ]
                    Return ONLY valid JSON with no extra markdown formatting or backticks if possible, or standard JSON inside ```json tags.
                    """

                    response = client.models.generate_content(
                        model="gemini-3.1-flash-lite",
                        contents=prompt,
                    )

                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"):
                        raw_text = raw_text[7:]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]

                    unit_data = json.loads(raw_text.strip())
                    st.session_state["unit_data"] = unit_data
                    st.session_state["wf_subject_val"] = final_subject
                    st.session_state["wf_grade_val"] = grade_level
                    st.success("Multi-session unit plan generated successfully!")

                except Exception as e:
                    st.error(f"Error generating unit plan: {e}")

    if "unit_data" in st.session_state:
        st.markdown("---")
        st.subheader("Unit Plan Preview (Cell Structure Mimic)")

        sessions = st.session_state["unit_data"]
        f_subj = st.session_state.get("wf_subject_val", "Subject")
        f_grade = st.session_state.get("wf_grade_val", "Grade")

        for idx, s in enumerate(sessions):
            with st.container():
                st.markdown(
                    f"**Week 1 | Unit: {f_subj} ({f_grade}) | Session {s.get('session_number', idx+1)} of {len(sessions)}**"
                )

                col_s1, col_s2, col_s3 = st.columns([1, 2, 3])
                with col_s1:
                    st.info(f"**Session**\n\nSession {s.get('session_number', idx+1)} of {len(sessions)}")
                with col_s2:
                    st.info(f"**Standard(s)**\n\n{s.get('standard', '')}")
                with col_s3:
                    st.info(f"**Learning Focus**\n\n{s.get('learning_focus', '')}")

                st.markdown("##### 5-E Instructional Pacing")
                e_cols = st.columns(5)
                phases = [
                    ("Engage", s.get("engage_time", ""), s.get("engage_desc", "")),
                    ("Explore", s.get("explore_time", ""), s.get("explore_desc", "")),
                    ("Explain", s.get("explain_time", ""), s.get("explain_desc", "")),
                    ("Elaborate", s.get("elaborate_time", ""), s.get("elaborate_desc", "")),
                    ("Evaluate", s.get("evaluate_time", ""), s.get("evaluate_desc", "")),
                ]
                for col, (p_name, p_time, p_desc) in zip(e_cols, phases):
                    with col:
                        st.markdown(
                            f"""
                            <div style="border: 1px solid #333; padding: 8px; border-radius: 4px; background-color: #1e1e1e; min-height: 140px;">
                                <b style="color: #4CAF50;">{p_name}</b><br>
                                <span style="font-size: 0.8em; color: #aaa;">Time: {p_time}</span><hr style="margin: 4px 0;">
                                <span style="font-size: 0.85em;">{p_desc}</span>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

                r_cols = st.columns(5)
                details = [
                    ("Textbook", s.get("textbook", "")),
                    ("Resources", s.get("resources", "")),
                    ("Evidence / Exit Task", s.get("evidence_exit_task", "")),
                    ("Support / Extension", s.get("support_extension", "")),
                    ("Adjustment / Next Step", s.get("adjustment_next_step", "")),
                ]
                for col, (d_title, d_val) in zip(r_cols, details):
                    with col:
                        st.markdown(
                            f"""
                            <div style="border: 1px solid #333; padding: 6px; border-radius: 4px; background-color: #161616; min-height: 100px;">
                                <span style="font-size: 0.8em; color: #888;"><b>{d_title}</b></span><br>
                                <span style="font-size: 0.8em;">{d_val}</span>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                st.markdown("<br>", unsafe_allow_html=True)

        wb = Workbook()
        ws = wb.active
        ws.title = "Curriculum Unit Plan"

        current_row = 1
        for s in sessions:
            ws.cell(row=current_row, column=1, value=f"Session {s.get('session_number')} of {len(sessions)}")
            ws.cell(row=current_row, column=2, value="Standard(s)")
            ws.cell(row=current_row, column=3, value="Learning Focus")
            current_row += 1

            ws.cell(row=current_row, column=1, value=f"Session {s.get('session_number')}")
            ws.cell(row=current_row, column=2, value=s.get("standard", ""))
            ws.cell(row=current_row, column=3, value=s.get("learning_focus", ""))
            current_row += 1

            phases_headers = ["Engage", "Explore", "Explain", "Elaborate", "Evaluate"]
            for col_idx, ph in enumerate(phases_headers, start=1):
                ws.cell(row=current_row, column=col_idx, value=ph)
            current_row += 1

            phase_vals = [
                f"Time: {s.get('engage_time')}\n{s.get('engage_desc')}",
                f"Time: {s.get('explore_time')}\n{s.get('explore_desc')}",
                f"Time: {s.get('explain_time')}\n{s.get('explain_desc')}",
                f"Time: {s.get('elaborate_time')}\n{s.get('elaborate_desc')}",
                f"Time: {s.get('evaluate_time')}\n{s.get('evaluate_desc')}",
            ]
            for col_idx, pval in enumerate(phase_vals, start=1):
                cell = ws.cell(row=current_row, column=col_idx, value=pval)
                cell.alignment = Alignment(wrap_text=True)
            current_row += 1

            detail_headers = [
                "Textbook",
                "Resources",
                "Evidence / Exit Task",
                "Support / Extension",
                "Adjustment / Next Step",
            ]
            for col_idx, dh in enumerate(detail_headers, start=1):
                ws.cell(row=current_row, column=col_idx, value=dh)
            current_row += 1

            detail_vals = [
                s.get("textbook", ""),
                s.get("resources", ""),
                s.get("evidence_exit_task", ""),
                s.get("support_extension", ""),
                s.get("adjustment_next_step", ""),
            ]
            for col_idx, dval in enumerate(detail_vals, start=1):
                cell = ws.cell(row=current_row, column=col_idx, value=dval)
                cell.alignment = Alignment(wrap_text=True)
            current_row += 2

        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)

        st.download_button(
            label="Download Complete Unit Plan (.xlsx)",
            data=excel_buffer,
            file_name=f"{f_subj}_{f_grade}_Unit_Plan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )