import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Professional Modern Palette (Slate Dark + Cyber Cyan + Warm Amber)
    BG_DARK = RGBColor(15, 23, 42)          # Deep Slate #0F172A
    CARD_BG = RGBColor(30, 41, 59)          # Card Slate #1E293B
    CARD_BORDER = RGBColor(51, 65, 85)      # Border Slate #334155
    ACCENT_CYAN = RGBColor(6, 182, 212)     # Cyan #06B6D4
    ACCENT_AMBER = RGBColor(245, 158, 11)   # Warm Amber #F59E0B
    ACCENT_ROSE = RGBColor(244, 63, 94)     # Rose #F43F5E
    ACCENT_PURPLE = RGBColor(168, 85, 247)  # Purple #A855F7
    TEXT_WHITE = RGBColor(248, 250, 252)    # Slate 50
    TEXT_MUTED = RGBColor(148, 163, 184)    # Slate 400

    def add_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_DARK
        bg.line.fill.background()
        return bg

    def add_header(slide, tag, title, subtitle=None):
        tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.5), Inches(0.35))
        tf_tag = tag_box.text_frame
        tf_tag.word_wrap = True
        p_tag = tf_tag.paragraphs[0]
        p_tag.text = tag.upper()
        p_tag.font.size = Pt(11)
        p_tag.font.bold = True
        p_tag.font.color.rgb = ACCENT_CYAN
        p_tag.font.name = "Segoe UI"

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.5), Inches(0.65))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = TEXT_WHITE
        p_title.font.name = "Segoe UI"

        if subtitle:
            sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(11.5), Inches(0.45))
            tf_sub = sub_box.text_frame
            tf_sub.word_wrap = True
            p_sub = tf_sub.paragraphs[0]
            p_sub.text = subtitle
            p_sub.font.size = Pt(13)
            p_sub.font.color.rgb = TEXT_MUTED
            p_sub.font.name = "Segoe UI"

    def add_card(slide, left, top, width, height, title, body_bullets, accent_color=ACCENT_CYAN):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1.5)

        tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), height - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        if title:
            p_t = tf.paragraphs[0]
            p_t.text = title
            p_t.font.size = Pt(17)
            p_t.font.bold = True
            p_t.font.color.rgb = accent_color
            p_t.font.name = "Segoe UI"
            p_t.space_after = Pt(8)

        first = True if not title else False
        for bullet in body_bullets:
            p = tf.add_paragraph() if not first else tf.paragraphs[0]
            first = False
            p.text = "• " + bullet
            p.font.size = Pt(12.5)
            p.font.color.rgb = TEXT_WHITE
            p.font.name = "Segoe UI"
            p.space_after = Pt(6)

    # ==================== SLIDE 1: TITLE ====================
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1)

    tbox = s1.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(11.0), Inches(4.5))
    tf1 = tbox.text_frame
    tf1.word_wrap = True

    p0 = tf1.paragraphs[0]
    p0.text = "🎬 CINEMATCH v2.0"
    p0.font.size = Pt(42)
    p0.font.bold = True
    p0.font.color.rgb = ACCENT_CYAN
    p0.font.name = "Segoe UI"
    p0.space_after = Pt(12)

    p1 = tf1.add_paragraph()
    p1.text = "Ending 'Scroll Paralysis' — The AI Movie Discovery Companion Built for Real Humans"
    p1.font.size = Pt(21)
    p1.font.color.rgb = TEXT_WHITE
    p1.font.name = "Segoe UI"
    p1.space_after = Pt(24)

    p2 = tf1.add_paragraph()
    p2.text = "🌐 Live Public URL: https://pauraveephuke2-a11y.github.io/cinematch/"
    p2.font.size = Pt(15)
    p2.font.bold = True
    p2.font.color.rgb = ACCENT_AMBER
    p2.font.name = "Segoe UI"
    p2.space_after = Pt(8)

    p3 = tf1.add_paragraph()
    p3.text = "A Humanized Walkthrough · What It Does, Why It Matters & How It Works"
    p3.font.size = Pt(13)
    p3.font.color.rgb = TEXT_MUTED

    # ==================== SLIDE 2: THE REAL PROBLEM ====================
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2)
    add_header(s2, "The Real-World Frustration", "The 45-Minute Friday Night Trap", "Why finding something to watch has become more exhausting than work.")
    
    add_card(s2, Inches(0.8), Inches(2.0), Inches(3.6), Inches(4.8), "🍕 The Cold Pizza Dilemma", [
        "It's 8:30 PM on Friday. You order hot food and sit down to unwind.",
        "You open Netflix, Prime Video, or Hotstar.",
        "40 minutes pass. You've scrolled through 200 thumbnails, watched 8 trailers, and argued with your partner.",
        "Your pizza is cold, the vibe is gone, and you end up rewatching an old sitcom."
    ], ACCENT_AMBER)

    add_card(s2, Inches(4.8), Inches(2.0), Inches(3.6), Inches(4.8), "🧠 The Fatigue of Infinite Choice", [
        "Modern streaming catalogues boast 15,000+ titles but zero clarity.",
        "Endless horizontal rows overload human working memory.",
        "Every movie claims to be a '98% Match' without explaining why.",
        "Result: Decision paralysis. Too many options, zero conviction."
    ], ACCENT_CYAN)

    add_card(s2, Inches(8.8), Inches(2.0), Inches(3.6), Inches(4.8), "🚫 The Empty Screen Wall", [
        "Traditional catalog search filters are rigid and binary.",
        "Try filtering: 'Sci-Fi' + 'Under 2 Hours' + 'On JioCinema'.",
        "Result: '0 Results Found — Clear your filters.'",
        "Algorithms fail us precisely when our requirements are specific."
    ], ACCENT_ROSE)

    # ==================== SLIDE 3: WHY CURRENT ALGORITHMS FAIL ====================
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3)
    add_header(s3, "Root Cause Analysis", "Why Big Streaming Algorithms Don't Understand You", "They are engineered for corporate watch-time, not your personal evening happiness.")

    add_card(s3, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8), "📉 What Netflix & Prime Do", [
        "Popularity Traps: They push trending titles or originals they spent $200M producing, regardless of your mood.",
        "Black-Box Percentages: A '97% Match' badge hides whether it matches your pacing, visual style, or plot complexity.",
        "One-Sided Logic: They try to guess what you might like, but give you NO way to exclude what you hate (gore, trauma, depressing endings).",
        "Solo-User Bias: They assume one person is watching, collapsing when two people sit on the couch together."
    ], ACCENT_ROSE)

    add_card(s3, Inches(6.8), Inches(2.0), Inches(5.6), Inches(4.8), "✨ What CineMatch Does Differently", [
        "Human Micro-Tropes: Reads films through 7,400+ emotional story tropes ('zero gravity fight', 'rain dance', 'twisty heist').",
        "100% Explainable AI (XAI): Interactive 5-Axis Spider charts show exact Complexity, Pacing, Emotion, Visuals, and Depth.",
        "Dealbreaker Shields: One-tap negative exclusion for gore, jump scares, depressing endings, and long runtimes.",
        "Consensus Solver: Minimax fairness math that solves movie arguments for couples."
    ], ACCENT_CYAN)

    # ==================== SLIDE 4: THE 5 GAME CHANGERS ====================
    s4 = prs.slides.add_slide(blank_layout)
    add_bg(s4)
    add_header(s4, "The CineMatch Philosophy", "Five Human-Centric Discoveries", "Taking you from 'What should we watch?' to watching in under 60 seconds.")

    add_card(s4, Inches(0.8), Inches(2.0), Inches(3.6), Inches(2.3), "🧪 Movie DNA Mixer", [
        "Mix two films like a cocktail recipe.",
        "Blend 70% Inception + 30% La La Land."
    ], ACCENT_CYAN)

    add_card(s4, Inches(4.8), Inches(2.0), Inches(3.6), Inches(2.3), "👥 Duo Consensus Solver", [
        "Stops Friday night arguments.",
        "Finds the Pareto-optimal sweet spot."
    ], ACCENT_AMBER)

    add_card(s4, Inches(8.8), Inches(2.0), Inches(3.6), Inches(2.3), "🚫 Dealbreaker Shield", [
        "Exclude triggers with one tap.",
        "No gore, no sad endings, strictly < 2 hrs."
    ], ACCENT_ROSE)

    add_card(s4, Inches(0.8), Inches(4.6), Inches(3.6), Inches(2.3), "⚡ Swipe to Decide", [
        "Tinder-style swipe gestures.",
        "Taste vector converges in 5 quick swipes."
    ], ACCENT_PURPLE)

    add_card(s4, Inches(4.8), Inches(4.6), Inches(3.6), Inches(2.3), "🔬 5-Axis Film DNA", [
        "Interactive radar charts on every title.",
        "See pacing & complexity before clicking play."
    ], ACCENT_CYAN)

    add_card(s4, Inches(8.8), Inches(4.6), Inches(3.6), Inches(2.3), "📺 India OTT Deep-Links", [
        "Direct watch links with free/paid badges.",
        "Netflix, Prime, Hotstar, JioCinema, SonyLIV."
    ], ACCENT_AMBER)

    # ==================== SLIDE 5: FEATURE 1 DNA MIXER ====================
    s5 = prs.slides.add_slide(blank_layout)
    add_bg(s5)
    add_header(s5, "Deep Dive · Feature 1", "The Movie DNA Mixer: Film Blending Like a Recipe", "Instead of choosing rigid categories, cook your own hybrid movie flavor.")

    add_card(s5, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8), "💡 How a Human Uses It", [
        "Imagine saying: 'I want the mind-bending mystery of Interstellar, but with the colorful musical energy of La La Land.'",
        "In CineMatch, pick Movie A and Movie B from a dropdown.",
        "Drag the ratio slider (e.g., 65% Cosmic Mystery : 35% Romance).",
        "Hit 'Synthesize Hybrid' and immediately get 8 curated films that bridge that exact artistic intersection.",
        "Takes 5 seconds. Sparks instant excitement."
    ], ACCENT_CYAN)

    add_card(s5, Inches(6.8), Inches(2.0), Inches(5.6), Inches(4.8), "🧠 The Engineering Under The Hood", [
        "Vector Arithmetic: V_hybrid = w * V_A + (1 - w) * V_B.",
        "Each film is mapped into an L2-normalized 7,431-dimensional TF-IDF space.",
        "Calculates real-time Cosine Similarity across all 100 films in memory.",
        "Zero server lag: Runs 100% inside your browser in under 2 milliseconds."
    ], ACCENT_AMBER)

    # ==================== SLIDE 6: FEATURE 2 DUO SOLVER ====================
    s6 = prs.slides.add_slide(blank_layout)
    add_bg(s6)
    add_header(s6, "Deep Dive · Feature 2", "The Duo Consensus Solver: Saving Couch Relationships", "No more 'Whatever you want' followed by two hours of quiet resentment.")

    add_card(s6, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8), "👫 The Human Scenario", [
        "Person 1 loves intense psychological thrillers and crime mysteries (The Dark Knight, Memento).",
        "Person 2 had a grueling workday and wants vibrant, feel-good comedy (3 Idiots, Amélie).",
        "Normal compromise means someone is bored to tears.",
        "CineMatch asks both viewers to pick one movie they love and computes a mathematically fair middle ground."
    ], ACCENT_AMBER)

    add_card(s6, Inches(6.8), Inches(2.0), Inches(5.6), Inches(4.8), "⚖️ The Minimax Fairness Formula", [
        "Consensus Score = (Sim_1 + Sim_2) / 2 - |Sim_1 - Sim_2| * 0.35.",
        "The penalty term prevents 'tyranny of the majority'—it penalizes movies that heavily favor one person while alienating the other.",
        "Finds gems that bridge both moods (e.g., Knives Out, Parasite, Vikram Vedha).",
        "Displays a 'Fairness Match %' so both parties feel respected."
    ], ACCENT_CYAN)

    # ==================== SLIDE 7: FEATURE 3 DEALBREAKER SHIELD ====================
    s7 = prs.slides.add_slide(blank_layout)
    add_bg(s7)
    add_header(s7, "Deep Dive · Feature 3", "The Dealbreaker Shield: Protecting Your Headspace", "Knowing what NOT to show you is just as vital as knowing what you like.")

    add_card(s7, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8), "🛡️ Human-First Negative Filters", [
        "You had a stressful week. You want a gripping film, but you CANNOT handle depressing hospital grief or trauma.",
        "Or you're watching with family and need to guarantee ZERO jump scares or gore.",
        "Or you have an early flight tomorrow and CANNOT watch a 3-hour epic.",
        "CineMatch gives you one-tap Dealbreaker Pills right below the search bar."
    ], ACCENT_ROSE)

    add_card(s7, Inches(6.8), Inches(2.0), Inches(5.6), Inches(4.8), "⚡ Instant Exclusion & Soft Fallback", [
        "🩸 No Gore: Drops slasher horror and graphic violence.",
        "👻 No Jump Scares: Keeps psychological tension without cheap frights.",
        "💔 No Depressing Endings: Guarantees emotional satisfaction.",
        "⏱️ Under 2 Hours: Enforces runtime <= 120 minutes.",
        "Zero-Result Safety: If tight filters eliminate 95% of films, the engine scores closest alternatives so you NEVER see an empty screen."
    ], ACCENT_CYAN)

    # ==================== SLIDE 8: FEATURE 4 SWIPE TO DECIDE ====================
    s8 = prs.slides.add_slide(blank_layout)
    add_bg(s8)
    add_header(s8, "Deep Dive · Feature 4", "Swipe to Decide: Gamified Taste Discovery", "When typing in a search box feels like too much work.")

    add_card(s8, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8), "📱 The Experience", [
        "Click '⚡ Swipe to Decide' on your phone or laptop.",
        "A poster card appears with film mood, tropes, and rating.",
        "👉 Swipe Right (or tap ❤️ Like) if you like the vibe.",
        "👈 Swipe Left (or tap ✕ Pass) to skip.",
        "In 5-8 quick swipes, CineMatch reveals 'Tonight's Winner' with 100% confidence.",
        "Playful, fast, and completely eliminates decision fatigue."
    ], ACCENT_PURPLE)

    add_card(s8, Inches(6.8), Inches(2.0), Inches(5.6), Inches(4.8), "📈 Dynamic Preference Gradient Engine", [
        "Taste Vector Convergence: Starts with a neutral vector V_taste.",
        "Swipe Right: V_taste = V_taste + 0.35 * V_movie.",
        "Swipe Left: V_taste = V_taste - 0.20 * V_movie (mild negative repulsion).",
        "Online Learning: Your taste profile actively shifts with every gesture.",
        "Supports touch swipes, keyboard arrow keys, and touch buttons."
    ], ACCENT_CYAN)

    # ==================== SLIDE 9: FEATURE 5 XAI RADAR ====================
    s9 = prs.slides.add_slide(blank_layout)
    add_bg(s9)
    add_header(s9, "Deep Dive · Feature 5", "Explainable AI (XAI): 5-Axis Film DNA Radar Charts", "Demystifying the algorithm. See the personality of every film at a glance.")

    add_card(s9, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8), "📊 The 5 Human Dimensions", [
        "🧩 Complexity: How hard will your brain work? (High: Inception, Low: 3 Idiots).",
        "⚡ Pacing: Is it an adrenaline thrill-ride or a slow, reflective burn?",
        "❤️ Emotion: Will it make you tear up, feel warm, or laugh out loud?",
        "🎨 Visuals: Stunning cinematography, CGI spectacles, and color palettes.",
        "🌌 Depth: Philosophical, existential layers vs. pure popcorn fun."
    ], ACCENT_CYAN)

    add_card(s9, Inches(6.8), Inches(2.0), Inches(5.6), Inches(4.8), "🔬 Why This Builds Trust", [
        "Traditional platforms hide their algorithms behind corporate black boxes.",
        "CineMatch renders an interactive SVG polygon radar on every movie card.",
        "Users glance at the chart in 1 second and immediately know if a movie fits their current energy level.",
        "Demystifies Artificial Intelligence: Transparent, educational, and fun."
    ], ACCENT_AMBER)

    # ==================== SLIDE 10: INDIAN OTT ECOSYSTEM ====================
    s10 = prs.slides.add_slide(blank_layout)
    add_bg(s10)
    add_header(s10, "Practical Everyday Utility", "Where to Watch in India: One-Tap Streaming", "Finding a movie is useless if you don't know where to actually click play.")

    add_card(s10, Inches(0.8), Inches(2.0), Inches(3.6), Inches(4.8), "🇮🇳 Native Indian Coverage", [
        "Curated database of 100 films across Bollywood, South Indian Cinema (Malayalam, Tamil, Telugu), and Hollywood.",
        "Deep-linked to India's major streaming services: Netflix, Prime Video, Disney+ Hotstar, JioCinema, SonyLIV, Zee5."
    ], ACCENT_CYAN)

    add_card(s10, Inches(4.8), Inches(2.0), Inches(3.6), Inches(4.8), "🔗 Direct Deep-Links", [
        "No more copying titles and searching manually inside 5 different streaming apps.",
        "Every card has a direct 'Watch on [Platform]' button with official platform colors.",
        "Opens the exact movie page on mobile apps and smart TVs."
    ], ACCENT_AMBER)

    add_card(s10, Inches(8.8), Inches(2.0), Inches(3.6), Inches(4.8), "🎟️ Binge Ticket Generator", [
        "Save movies to your private offline Watchlist.",
        "Click 'Download Binge Ticket' to generate a vintage cinematic movie ticket on HTML5 Canvas.",
        "Export as high-res PNG to share with friends on WhatsApp or Instagram stories."
    ], ACCENT_PURPLE)

    # ==================== SLIDE 11: TECHNICAL SIMPLICITY ====================
    s11 = prs.slides.add_slide(blank_layout)
    add_bg(s11)
    add_header(s11, "Engineering Highlights", "Fast, Private, and Zero-Maintenance Architecture", "High mathematical rigor packaged into a zero-latency web application.")

    add_card(s11, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8), "⚡ 100% Client-Side Speed", [
        "Zero Cloud Server Cost: Precomputed TF-IDF Cosine Similarity matrix embedded directly into static JSON payload.",
        "Sub-Millisecond Query Time: All vector blending, minimax consensus, and filtering execute in browser memory in < 5ms.",
        "100% Privacy-Preserving: User watchlists, swipe history, and ratings stay in browser LocalStorage—no tracking, no data leaks."
    ], ACCENT_CYAN)

    add_card(s11, Inches(6.8), Inches(2.0), Inches(5.6), Inches(4.8), "🎯 Zero-Result Fallback Algorithm", [
        "Solves the notorious 'Cold Start & Sparse Constraint' problem.",
        "If exact boolean filters match < 3 films, CineMatch automatically applies a relaxed soft-relevance scoring function (+50 text, +30 genre, +30 OTT, +20 mood, -40 dealbreaker penalty).",
        "Displays closest alternatives with a clear 'Showing closest partial matches' banner.",
        "Result: 100% reliability under every edge case."
    ], ACCENT_AMBER)

    # ==================== SLIDE 12: CONCLUSION ====================
    s12 = prs.slides.add_slide(blank_layout)
    add_bg(s12)
    add_header(s12, "Summary & Live Demo", "Experience CineMatch Live Today", "Putting the joy and curiosity back into movie night.")

    add_card(s12, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8), "🚀 What We Shipped", [
        "Transformed a generic catalog into an intelligent discovery companion.",
        "Solved decision fatigue with 5 distinct human-first interaction paradigms.",
        "100 Curated global & regional masterpieces ready to stream.",
        "Completely responsive across phones, tablets, and laptops.",
        "Fully deployed and accessible worldwide via GitHub Pages."
    ], ACCENT_CYAN)

    add_card(s12, Inches(6.8), Inches(2.0), Inches(5.6), Inches(4.8), "🌐 Live Links & Project Info", [
        "Live Website: https://pauraveephuke2-a11y.github.io/cinematch/",
        "GitHub Repository: https://github.com/pauraveephuke2-a11y/cinematch",
        "Creator: pauraveephuke2-a11y (phukearnav0@gmail.com)",
        "Technology Stack: Python, Scikit-Learn TF-IDF, HTML5 Canvas, SVG Radar, Vanilla JS."
    ], ACCENT_AMBER)

    out_path = r"C:\Users\Nucom\Projects\MovieRecommender\CineMatch_Humanized_Presentation.pptx"
    prs.save(out_path)
    print(f"Presentation saved successfully to {out_path}")

if __name__ == "__main__":
    create_deck()
