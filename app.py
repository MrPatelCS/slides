import io
import json
import os
from google import genai
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
import streamlit as st

# Page configuration
st.set_page_config(page_title="Lesson Slide Generator", layout="centered")

st.title("Slide Generator")
st.write("Generate classroom slides.")

# Sidebar for API Key input
st.sidebar.header("Configuration")

# Fetch API key from Streamlit secrets if available, otherwise allow input
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


def get_balanced_font_size(text):
  """Balanced font scaling that adjusts properly to content length

  without squashing lists or leaving excessive dead space.
  """
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
  """Programmatically draws a distinct 180x180 dual-coding icon (vector graphic)

  representing the pedagogical stage (Engage, Explore, Explain, Elaborate,
  Evaluate).
  """
  img = Image.new("RGBA", (180, 180), (255, 255, 255, 0))
  draw = ImageDraw.Draw(img)

  # Base rounded background badge (Teal theme)
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
    draw.rounded_rectangle(
        [50, 45, 130, 140], radius=15, outline=white, width=8
    )
    draw.line([68, 85, 85, 105], fill=accent_yellow, width=8)
    draw.line([85, 105, 115, 70], fill=accent_yellow, width=8)

  img_byte_arr = io.BytesIO()
  img.save(img_byte_arr, format="PNG")
  img_byte_arr.seek(0)
  return img_byte_arr


# --- Radio toggle to switch between Text Only and Files + Instructions ---
generation_mode = st.radio(
    "Choose Generation Input Mode:",
    ["Text Only", "Files + Instructions"],
    horizontal=True,
)

with st.form("slide_form"):
  if generation_mode == "Text Only":
    lesson_input = st.text_area("Paste Lesson Plan & Text Here:", height=250)
    uploaded_files = []
  else:
    uploaded_files = st.file_uploader(
        "Upload Documents / Screenshots (PDF, PNG, JPG - Max 5 pages"
        " recommended):",
        type=["png", "jpg", "jpeg", "pdf"],
        accept_multiple_files=True,
    )
    lesson_input = st.text_area(
        "Enter Additional Instructions / Lesson Plan (Optional):", height=120
    )

  # Embedded and editable prompt configuration expander
  with st.expander("⚙️ Customize Slide Section Prompts (Advanced)", expanded=False):
    st.write(
        "Modify the generation guidelines for each slide component below:"
    )
    p_do_now = st.text_input(
        "do_now",
        "Student-facing warm-up task, prediction question, or initial"
        " engagement prompt.",
    )
    p_learning_focus = st.text_input(
        "learning_focus",
        "Must start explicitly with 'I can [skill] by [method].'",
    )
    p_activity_title = st.text_input(
        "activity_title", "Student-facing title of the main activity."
    )
    p_new_learning = st.text_input(
        "new_learning",
        "Student-facing informational text or core conceptual bullet points"
        " summarizing the reading/background.",
    )
    p_vocabulary = st.text_input(
        "vocabulary",
        "A clean list of key vocabulary words and definitions relevant to the"
        " text.",
    )
    p_modeling = st.text_input(
        "modeling",
        "A clear model example or guiding question showing students what a"
        " strong response looks like.",
    )
    p_guided_practice = st.text_input(
        "guided_practice",
        "Discussion prompts or collaborative questions for students to analyze"
        " together in pairs.",
    )
    p_independent_practice = st.text_input(
        "independent_practice",
        "The actual complete practice task, writing prompt, or exercise for"
        " students to solve independently.",
    )
    p_exit_task = st.text_input(
        "exit_task",
        "A final student-facing reflection or check-for-understanding prompt.",
    )

  submitted = st.form_submit_button("Generate Slides")

if submitted:
  if not api_key_input.strip():
    st.error("Please enter your API key in the sidebar first!")
  elif generation_mode == "Text Only" and not lesson_input.strip():
    st.warning("Please paste some lesson text first!")
  elif generation_mode == "Files + Instructions" and not uploaded_files:
    st.warning("Please upload at least one file (PDF or image) first!")
  else:
    with st.spinner("Generating slides..."):
      try:
        client = genai.Client(api_key=api_key_input.strip())

        base_prompt = f"""
                You are an expert curriculum designer creating a projector-ready, student-facing slide deck. 
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
                2. NO EMPTY VALUES: Fill out every field with rich, complete content.
                """

        if generation_mode == "Text Only":
          prompt = (
              base_prompt
              + f"\nTake the following lesson plan and text and map it completely into the required structure:\n{lesson_input}"
          )
          api_contents = prompt
        else:
          prompt = (
              base_prompt
              + f"\nAnalyze the provided files and incorporate the following instructions to map everything completely into the required structure:\nInstructions: {lesson_input}"
          )
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
            st.error(
                f"Template not found at path: {template_path}. Please include"
                " 'template.pptx' in your GitHub repository."
            )
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
                2: (
                    "Engage",
                    format_slide_content(deck_data.get("learning_focus", "")),
                ),
                3: (
                    "Explore",
                    format_slide_content(deck_data.get("activity_title", "")),
                ),
                4: (
                    "Explain",
                    format_slide_content(deck_data.get("new_learning", "")),
                ),
                5: (
                    "Explain",
                    format_slide_content(deck_data.get("vocabulary", "")),
                ),
                6: (
                    "Explain",
                    format_slide_content(deck_data.get("modeling", "")),
                ),
                7: (
                    "Elaborate",
                    format_slide_content(deck_data.get("guided_practice", "")),
                ),
                8: (
                    "Elaborate",
                    format_slide_content(
                        deck_data.get("independent_practice", "")
                    ),
                ),
                9: (
                    "Evaluate",
                    format_slide_content(deck_data.get("exit_task", "")),
                ),
            }

            for s_idx, (phase_name, text_val) in slide_mapping.items():
              if len(prs.slides) > s_idx:
                slide = prs.slides[s_idx]

                content_box = slide.shapes.add_textbox(
                    Inches(1.2), Inches(1.6), Inches(11.0), Inches(5.0)
                )
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

                # Fixed check to include slide index 1 (Do Now) onwards
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

            st.success("Presentation generated successfully!")
            with open(output_path, "rb") as f:
              st.download_button(
                  label="Download Gemini Presentation (.pptx)",
                  data=f,
                  file_name="StudySync_Gemini_Lesson.pptx",
                  mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
              )

      except Exception as e:
        st.error(f"Error processing generation: {e}")