import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
from gtts import gTTS
import pygame
import os
import re
import threading

class PDFVoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDFVoice - PDF to Audiobook Converter")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        self.root.configure(bg="#f0f0f0")

        # Initialize pygame for audio playback
        pygame.mixer.init()
        self.audio_file = None
        self.is_playing = False
        self.is_paused = False

        # Set up accessibility attributes
        self.root.option_add("*Font", "Helvetica 12")
        self.root.option_add("*Button*highlightThickness", 2)
        self.root.option_add("*Button*relief", "raised")

        # GUI Elements
        self.setup_gui()

    def setup_gui(self):
        # Main frame
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Upload PDF button
        self.upload_btn = ttk.Button(self.main_frame, text="Upload PDF", command=self.upload_pdf)
        self.upload_btn.grid(row=0, column=0, columnspan=2, pady=10, sticky="ew")
        self.upload_btn.focus_set()  # Accessibility: Set initial focus

        # Text area for displaying extracted text
        self.text_area = tk.Text(self.main_frame, height=15, width=70, wrap=tk.WORD, state='disabled')
        self.text_area.grid(row=1, column=0, columnspan=2, pady=10, padx=5)
        self.text_scroll = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.text_area.yview)
        self.text_area['yscrollcommand'] = self.text_scroll.set
        self.text_scroll.grid(row=1, column=2, sticky="ns")

        # Playback controls
        self.play_btn = ttk.Button(self.main_frame, text="Play", command=self.play_audio)
        self.play_btn.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

        self.pause_btn = ttk.Button(self.main_frame, text="Pause", command=self.pause_audio, state='disabled')
        self.pause_btn.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        self.stop_btn = ttk.Button(self.main_frame, text="Stop", command=self.stop_audio, state='disabled')
        self.stop_btn.grid(row=2, column=2, padx=5, pady=5, sticky="ew")

        # Volume and speed sliders
        self.volume_label = ttk.Label(self.main_frame, text="Volume:")
        self.volume_label.grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.volume_slider = ttk.Scale(self.main_frame, from_=0, to=1, orient=tk.HORIZONTAL, command=self.set_volume)
        self.volume_slider.set(0.7)  # Default volume
        self.volume_slider.grid(row=3, column=1, columnspan=2, sticky="ew", padx=5, pady=5)

        self.speed_label = ttk.Label(self.main_frame, text="Speed (0.5x - 2x):")
        self.speed_label.grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.speed_slider = ttk.Scale(self.main_frame, from_=0.5, to=2.0, orient=tk.HORIZONTAL)
        self.speed_slider.set(1.0)  # Default speed
        self.speed_slider.grid(row=4, column=1, columnspan=2, sticky="ew", padx=5, pady=5)

        # Export to MP3 button
        self.export_btn = ttk.Button(self.main_frame, text="Export to MP3", command=self.export_mp3, state='disabled')
        self.export_btn.grid(row=5, column=0, columnspan=3, pady=10, sticky="ew")

        # Status label for user feedback
        self.status_label = ttk.Label(self.main_frame, text="Ready to upload a PDF file.", foreground="blue")
        self.status_label.grid(row=6, column=0, columnspan=3, pady=10)

        # Accessibility: Bind keyboard shortcuts
        self.root.bind('<Control-o>', lambda event: self.upload_pdf())
        self.root.bind('<Control-p>', lambda event: self.play_audio())
        self.root.bind('<Control-s>', lambda event: self.stop_audio())
        self.root.bind('<Control-e>', lambda event: self.export_mp3())

    def upload_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not file_path:
            return

        self.status_label.config(text="Extracting text from PDF...", foreground="blue")
        self.root.update()

        try:
            # Extract text from PDF using pdfplumber
            with pdfplumber.open(file_path) as pdf:
                text = ""
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text and page_text.strip():  # Skip empty pages
                        text += page_text + "\n"

            # Clean extracted text
            text = self.clean_text(text)
            if not text:
                messagebox.showerror("Error", "No readable text found in the PDF.")
                self.status_label.config(text="Ready to upload a PDF file.", foreground="blue")
                return

            # Display extracted text
            self.text_area.config(state='normal')
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(tk.END, text)
            self.text_area.config(state='disabled')

            # Generate temporary audio file for playback
            self.generate_audio(text)
            self.export_btn.config(state='normal')
            self.status_label.config(text="PDF loaded successfully. Ready to play or export.", foreground="green")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process PDF: {str(e)}")
            self.status_label.config(text="Failed to load PDF.", foreground="red")

    def clean_text(self, text):
        # Remove excessive newlines, spaces, and unwanted characters
        text = re.sub(r'\n\s*\n', '\n', text)  # Remove multiple newlines
        text = re.sub(r'\s+', ' ', text)  # Replace multiple spaces with single space
        text = text.strip()
        return text

    def generate_audio(self, text):
        try:
            # Generate audio with specified speed
            tts = gTTS(text=text, lang='en', slow=False)
            self.audio_file = "temp_audio.mp3"
            tts.save(self.audio_file)
            pygame.mixer.music.load(self.audio_file)
            self.set_volume(None)  # Apply default volume
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate audio: {str(e)}")
            self.audio_file = None

    def play_audio(self):
        if not self.audio_file or not os.path.exists(self.audio_file):
            messagebox.showerror("Error", "No audio available. Please upload a PDF first.")
            return

        if not self.is_playing:
            pygame.mixer.music.play()
            self.is_playing = True
            self.play_btn.config(state='disabled')
            self.pause_btn.config(state='normal')
            self.stop_btn.config(state='normal')
            self.status_label.config(text="Playing audio...", foreground="blue")
            threading.Thread(target=self.monitor_playback, daemon=True).start()
        elif self.is_paused:
            pygame.mixer.music.unpause()
            self.is_paused = False
            self.play_btn.config(state='disabled')
            self.pause_btn.config(state='normal')
            self.status_label.config(text="Playing audio...", foreground="blue")

    def pause_audio(self):
        if self.is_playing and not self.is_paused:
            pygame.mixer.music.pause()
            self.is_paused = True
            self.play_btn.config(state='normal')
            self.pause_btn.config(state='disabled')
            self.status_label.config(text="Audio paused.", foreground="blue")

    def stop_audio(self):
        if self.is_playing:
            pygame.mixer.music.stop()
            self.is_playing = False
            self.is_paused = False
            self.play_btn.config(state='normal')
            self.pause_btn.config(state='disabled')
            self.stop_btn.config(state='disabled')
            self.status_label.config(text="Audio stopped.", foreground="blue")

    def set_volume(self, event):
        volume = self.volume_slider.get()
        pygame.mixer.music.set_volume(volume)

    def export_mp3(self):
        if not self.audio_file or not os.path.exists(self.audio_file):
            messagebox.showerror("Error", "No audio available to export. Please upload a PDF first.")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".mp3", filetypes=[("MP3 Files", "*.mp3")])
        if output_path:
            try:
                os.rename(self.audio_file, output_path)
                self.audio_file = None
                self.export_btn.config(state='disabled')
                self.status_label.config(text=f"Audio exported to {output_path}.", foreground="green")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export MP3: {str(e)}")

    def monitor_playback(self):
        while self.is_playing and pygame.mixer.music.get_busy():
            self.root.update()
        if self.is_playing and not self.is_paused:
            self.stop_audio()

    def on_closing(self):
        if self.audio_file and os.path.exists(self.audio_file):
            pygame.mixer.music.stop()
            os.remove(self.audio_file)
        pygame.mixer.quit()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFVoiceApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()