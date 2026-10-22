import streamlit as st
import win32com.client as wincom
import pythoncom

st.set_page_config(
    page_title="Text-to-Speech Converter",
    page_icon="🔊",
    layout="centered"
)

st.title("🔊 Text-to-Speech Converter")
st.write("Enter your text and let Python speak it aloud.")

text = st.text_area(
    "Enter the text you want to speak:",
    placeholder="Type something here...",
    height=150
)


def speak_text(text):
    pythoncom.CoInitialize()
    try:
        speaker = wincom.Dispatch("SAPI.SpVoice")
        speaker.Speak(text)
    finally:
        pythoncom.CoUninitialize()


def stop_speech():
    pythoncom.CoInitialize()
    try:
        speaker = wincom.Dispatch("SAPI.SpVoice")
        speaker.Speak("", 3)
    finally:
        pythoncom.CoUninitialize()


col1, col2 = st.columns(2)

with col1:
    speak_button = st.button("🔊 Speak", use_container_width=True)

with col2:
    stop_button = st.button("⏹️ Stop", use_container_width=True)

if speak_button:
    if text.strip():
        try:
            speak_text(text)
            st.success("Speech completed!")
        except Exception as e:
            st.error(f"Speech error: {e}")
    else:
        st.warning("Please enter some text first.")

if stop_button:
    try:
        stop_speech()
        st.info("Speech stopped.")
    except Exception as e:
        st.error(f"Unable to stop speech: {e}")

st.divider()
st.caption("Powered by Python, Streamlit, and Windows SAPI.")