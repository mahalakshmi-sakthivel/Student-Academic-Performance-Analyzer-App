import os
import subprocess
import glob
from PIL import Image, ImageDraw, ImageFont

FFMPEG_PATH = r"C:\Users\MAHALAKSHMI S\AppData\Local\Microsoft\WinGet\Packages\Gyan.FFmpeg.Essentials_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-9.0.1-essentials_build\bin\ffmpeg.exe"
OUT_VIDEO = os.path.abspath("Student-Academic-Performance-Analyzer-Explainer.mp4")

SCENES = [
    {
        "tag": "SCENE 1 • 0:00 - 0:24",
        "title": "Student Academic Performance Analyzer",
        "subtitle": "An offline-first academic analytics application for tracking student performance, calculating grades, and generating actionable insights.",
        "cards": [
            ("Offline First", "Desktop & Web"),
            ("Client-Side", "Instant & Secure"),
            ("Zero Backend", "Web Storage API")
        ],
        "speech": "Welcome to the Student Academic Performance Analyzer. In educational institutions, tracking student marks, attendance, and subject arrears across various semesters often leads to scattered data and manual calculation errors. This application solves that problem by providing a fast, offline-first dashboard that analyzes student academic records and delivers actionable insights in real time."
    },
    {
        "tag": "SCENE 2 • 0:24 - 0:48",
        "title": "Calculation Engine & Assessment Model",
        "subtitle": "Evaluates 4 normalized assessment components per subject, combining subject average with attendance and calculating a 10-point CGPA.",
        "cards": [
            ("Internal", "Max 60 marks"),
            ("Assignment", "Max 10 marks"),
            ("Unit Test", "Max 15 marks"),
            ("Final Exam", "Max 100 marks")
        ],
        "speech": "At the core of the system is a transparent calculation engine. Each subject is evaluated using four normalized components: Internal marks, Assignments, Unit Tests, and Final Semester Exams. The overall student score combines the subject average at 85% with attendance at 15%. In addition, a 10-point CGPA scale is automatically computed."
    },
    {
        "tag": "SCENE 3 • 0:48 - 1:12",
        "title": "Arrear Detection & Personalized Feedback",
        "subtitle": "Independent subject arrear detection flags any subject below 50%, while the rule engine recommends specific study actions.",
        "cards": [
            ("Arrear Alert", "Subject < 50%"),
            ("10-pt CGPA", "Standard Scale"),
            ("Rule Engine", "Tailored Guidance")
        ],
        "speech": "A key innovation is independent subject arrear detection. Even if a student achieves a good overall score, any subject below 50% is highlighted with a clear red arrear alert. The system also runs a rule-based advisor that generates personalized recommendations, advising students on attendance improvement, exam preparation, and their weakest subjects."
    },
    {
        "tag": "SCENE 4 • 1:12 - 1:36",
        "title": "Class Analytics & Multi-Level Filtering",
        "subtitle": "Real-time class monitoring with KPI counters, search bar, and filters across 14 departments, sections, batches, and semesters.",
        "cards": [
            ("14 Departments", "Engineering"),
            ("2,940 Students", "Deterministic Demo"),
            ("Multi-Filter", "Dept • Sec • Batch • Sem")
        ],
        "speech": "For educators and mentors, the Class Analytics dashboard provides immediate class-level visibility. It includes interactive KPI cards and supports filtering across 14 engineering departments, sections A, B, and C, academic batches, and semesters. With the built-in demo generator, over 2,900 realistic student records can be analyzed instantly."
    },
    {
        "tag": "SCENE 5 • 1:36 - 2:00",
        "title": "PDF Reporting & Desktop Deployment",
        "subtitle": "Generate print-ready A4 reports with interactive live preview, packaged natively with Electron and NSIS for Windows.",
        "cards": [
            ("PDF Preview", "Before Print / Save"),
            ("A4 Cards", "Batch & Semester Groups"),
            ("Electron App", "Offline Windows Setup")
        ],
        "speech": "Finally, the reporting module produces structured, print-ready A4 PDF reports with a live preview before saving. Packaged as a native Windows desktop application with Electron, it requires no internet connection or external server, ensuring complete privacy, reliability, and ease of use. Thank you!"
    }
]

def make_audio(text, out_wav):
    ps_cmd = f"""
    $s = New-Object -ComObject SAPI.SpVoice
    $s.Rate = 0
    $f = New-Object -ComObject SAPI.SpFileStream
    $f.Open('{out_wav}', 3, 0)
    $s.AudioOutputStream = $f
    $s.Speak('{text}')
    $f.Close()
    """
    script_path = "temp_sapi.ps1"
    with open(script_path, "w", encoding="utf-8") as fp:
        fp.write(ps_cmd)
    subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path], check=True)
    if os.path.exists(script_path):
        os.remove(script_path)

def get_font(size, bold=False):
    font_names = ["segoeui.ttf", "arial.ttf", "tahoma.ttf"]
    if bold:
        font_names = ["segoeuib.ttf", "arialbd.ttf", "tahomabd.ttf"]
    for fn in font_names:
        fp = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts", fn)
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                pass
    return ImageFont.load_default()

def create_slide_image(scene, out_png):
    width, height = 1920, 1080
    im = Image.new("RGB", (width, height), color="#0b0f19")
    draw = ImageDraw.Draw(im)

    # Top brand bar
    draw.rectangle([(0, 0), (width, 100)], fill="#111827")
    draw.line([(0, 100), (width, 100)], fill="#1f2937", width=2)

    font_brand = get_font(32, bold=True)
    font_sub = get_font(20, bold=False)
    draw.text((80, 24), "SA", font=font_brand, fill="#60a5fa")
    draw.text((140, 24), "Student Academic Performance Analyzer", font=font_brand, fill="#ffffff")
    draw.text((140, 64), "Project Overview & Architecture Walkthrough", font=font_sub, fill="#94a3b8")

    # Scene Tag badge
    font_tag = get_font(24, bold=True)
    tag_text = scene["tag"]
    draw.rounded_rectangle([(80, 150), (460, 200)], radius=25, fill="#1e293b", outline="#3b82f6", width=2)
    draw.text((105, 160), tag_text, font=font_tag, fill="#93c5fd")

    # Main Title
    font_title = get_font(56, bold=True)
    draw.text((80, 230), scene["title"], font=font_title, fill="#ffffff")

    # Subtitle
    font_desc = get_font(30, bold=False)
    draw.text((80, 310), scene["subtitle"], font=font_desc, fill="#cbd5e1")

    # Visual Cards
    cards = scene["cards"]
    card_width = 380
    card_height = 200
    start_x = 80
    start_y = 420
    gap = 40

    font_c_title = get_font(36, bold=True)
    font_c_sub = get_font(24, bold=False)

    for i, (ct, cs) in enumerate(cards):
        cx = start_x + i * (card_width + gap)
        draw.rounded_rectangle([(cx, start_y), (cx + card_width, start_y + card_height)], radius=18, fill="#141b2d", outline="#334155", width=2)
        draw.text((cx + 30, start_y + 45), ct, font=font_c_title, fill="#60a5fa")
        draw.text((cx + 30, start_y + 115), cs, font=font_c_sub, fill="#94a3b8")

    # Bottom Narration Caption Box
    draw.rounded_rectangle([(80, 840), (width - 80, 1000)], radius=16, fill="#0f172a", outline="#1e293b", width=2)
    font_caption_tag = get_font(24, bold=True)
    font_caption = get_font(24, bold=False)
    draw.text((110, 860), "🎙 NARRATOR:", font=font_caption_tag, fill="#38bdf8")
    
    # Simple word wrap for caption
    words = scene["speech"].split()
    line1, line2 = "", ""
    for w in words:
        if len(line1) + len(w) < 115:
            line1 += w + " "
        else:
            line2 += w + " "
    draw.text((110, 900), line1.strip(), font=font_caption, fill="#e2e8f0")
    if line2:
        draw.text((110, 938), line2.strip(), font=font_caption, fill="#cbd5e1")

    im.save(out_png)

def main():
    print("Generating 2-minute video presentation...")
    temp_files = []
    video_segments = []

    for i, sc in enumerate(SCENES):
        print(f"Processing Scene {i+1} of {len(SCENES)}...")
        wav_file = f"scene_{i}.wav"
        png_file = f"scene_{i}.png"
        seg_mp4 = f"segment_{i}.mp4"

        make_audio(sc["speech"], wav_file)
        create_slide_image(sc, png_file)
        temp_files.extend([wav_file, png_file, seg_mp4])

        # Create video segment
        cmd = [
            FFMPEG_PATH, "-y",
            "-loop", "1", "-i", png_file,
            "-i", wav_file,
            "-c:v", "libx264", "-tune", "stillimage",
            "-c:a", "aac", "-b:a", "192k",
            "-pix_fmt", "yuv420p",
            "-shortest",
            seg_mp4
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        video_segments.append(seg_mp4)

    # Concat segments
    concat_txt = "concat_list.txt"
    temp_files.append(concat_txt)
    with open(concat_txt, "w", encoding="utf-8") as f:
        for seg in video_segments:
            f.write(f"file '{os.path.abspath(seg)}'\n")

    print("Stitching scenes into final MP4 video...")
    cmd_concat = [
        FFMPEG_PATH, "-y",
        "-f", "concat", "-safe", "0",
        "-i", concat_txt,
        "-c", "copy",
        OUT_VIDEO
    ]
    subprocess.run(cmd_concat, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # Cleanup temporary segment files
    for tf in temp_files:
        if os.path.exists(tf):
            try:
                os.remove(tf)
            except Exception:
                pass

    if os.path.exists("test_audio.wav"):
        os.remove("test_audio.wav")
    if os.path.exists("test_sapi.ps1"):
        os.remove("test_sapi.ps1")

    size_mb = os.path.getsize(OUT_VIDEO) / (1024 * 1024)
    print(f"SUCCESS: Created {OUT_VIDEO} ({size_mb:.2f} MB)")

if __name__ == "__main__":
    main()
