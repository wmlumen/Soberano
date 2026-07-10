import os
import re

panel_dir = r"c:\Users\HP 250 G10\Documents\GITHUT\RItoMemphisMisraim\Web\panel"
dashboards = [
    "dashboard.html",
    "dashboard_logias.html",
    "dashboard_decretos.html",
    "dashboard_solicitudes.html",
    "dashboard_actas.html",
    "dashboard_tesoreria.html",
    "dashboard_usuarios.html"
]

# Read the base dashboard
with open(os.path.join(panel_dir, "dashboard.html"), "r", encoding="utf-8") as f:
    base_html = f.read()

# Find the content area
content_regex = re.compile(r'(<div class="container-fluid mt-4">.*?)</div>\s*</div>\s*</div>\s*<!-- Scripts -->', re.DOTALL)
base_match = content_regex.search(base_html)

if not base_match:
    print("Could not find content area in dashboard.html")
    exit(1)

# We will replace the single container-fluid with multiple divs (one for each section)
combined_content = ""
combined_scripts = ""

for dash in dashboards:
    if not os.path.exists(os.path.join(panel_dir, dash)):
        continue
        
    with open(os.path.join(panel_dir, dash), "r", encoding="utf-8") as f:
        html = f.read()
        
    # Extract content
    match = content_regex.search(html)
    if match:
        section_id = "section-" + dash.replace("dashboard_", "").replace(".html", "").replace("dashboard", "obreros")
        display = 'style="display: block;"' if dash == "dashboard.html" else 'style="display: none;"'
        
        # wrap in section div
        content = f'<div id="{section_id}" class="dashboard-section" {display}>\n' + match.group(1) + '\n</div>\n'
        combined_content += content
        
    # Extract scripts inside <script type="module">
    script_regex = re.compile(r'<script type="module">(.*?)</script>', re.DOTALL)
    script_match = script_regex.search(html)
    if script_match:
        # Avoid duplicate imports
        script_code = script_match.group(1)
        combined_scripts += f"\n// Scripts from {dash}\n" + script_code

# Clean up duplicate imports in combined scripts
imports = set()
cleaned_scripts = ""
for line in combined_scripts.split('\n'):
    if line.strip().startswith('import '):
        if line.strip() not in imports:
            imports.add(line.strip())
            cleaned_scripts += line + "\n"
    elif 'const firebaseConfig' in line or 'initializeApp' in line or 'getAuth' in line or 'getFirestore' in line:
        # We only want one initialization
        pass
    else:
        cleaned_scripts += line + "\n"

# Add the one true initialization at the top of scripts
firebase_init = """
        import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
        import { getAuth, onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
        import { getFirestore, collection, getDocs, query, doc, getDoc, setDoc, addDoc, orderBy } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

        const firebaseConfig = {
            apiKey: "AIzaSyABX07keytsClNlr-93pxsljEVZAWo3Tpg",
            authDomain: "mmpy-146a6.firebaseapp.com",
            projectId: "mmpy-146a6",
            storageBucket: "mmpy-146a6.firebasestorage.app",
            messagingSenderId: "215936118140",
            appId: "1:215936118140:web:4e9635a8b925bf5d5a88f9",
            measurementId: "G-YDWYBKMMQR"
        };

        const app = initializeApp(firebaseConfig);
        const auth = getAuth(app);
        const db = getFirestore(app);
"""

final_script = firebase_init + cleaned_scripts

# Replace content
new_html = base_html[:base_match.start()] + combined_content + "\n</div>\n</div>\n<!-- Scripts -->" + base_html[base_match.end():]

# Replace module script
new_html = re.sub(r'<script type="module">.*?</script>', f'<script type="module">{final_script}</script>', new_html, flags=re.DOTALL)

# Add router logic
router_logic = """
<script>
function showSection(sectionId) {
    document.querySelectorAll('.dashboard-section').forEach(el => el.style.display = 'none');
    document.querySelectorAll('.side-nav-link').forEach(el => el.classList.remove('active'));
    
    document.getElementById('section-' + sectionId).style.display = 'block';
    event.currentTarget.classList.add('active');
}
</script>
"""
new_html = new_html.replace('<!-- Scripts -->', router_logic + '<!-- Scripts -->')

# Update sidebar links to use onclick="showSection('...')"
new_html = new_html.replace('href="dashboard.html"', 'href="#" onclick="showSection(\'obreros\')"')
new_html = new_html.replace('href="dashboard_logias.html"', 'href="#" onclick="showSection(\'logias\')"')
new_html = new_html.replace('href="dashboard_decretos.html"', 'href="#" onclick="showSection(\'decretos\')"')
new_html = new_html.replace('href="dashboard_solicitudes.html"', 'href="#" onclick="showSection(\'solicitudes\')"')
new_html = new_html.replace('href="dashboard_actas.html"', 'href="#" onclick="showSection(\'actas\')"')
new_html = new_html.replace('href="dashboard_tesoreria.html"', 'href="#" onclick="showSection(\'tesoreria\')"')
new_html = new_html.replace('href="dashboard_usuarios.html"', 'href="#" onclick="showSection(\'usuarios\')"')

with open(os.path.join(panel_dir, "dashboard_unified.html"), "w", encoding="utf-8") as f:
    f.write(new_html)

print("Created dashboard_unified.html")
