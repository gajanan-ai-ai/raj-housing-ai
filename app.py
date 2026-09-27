
import gradio as gr
import re
import urllib.parse
import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import os

file_path = "Raj_Housing_AI_FINAL_Verified_Property_Database.xlsx"
df = pd.read_excel(file_path)

exact_locations = {
    "Raj Durga Tower": "Opposite Our Lady of Mount Carmel Chapel, Sadar, Ponda, Goa 403401",
    "Raj Elite": "Near Goa Urban Bank, Sadar, Ponda, Goa 403401",
    "Raj City Plaza": "Sadar, Ponda, Goa 403401",
    "Raj Enclave": "Karaswada, Mapusa, North Goa 403526",
    "Raj Ryle Residency": "Acoi Village, Karaswada, Mapusa, North Goa 403526",
    "Raj Durga Heritage": "Dag, Ponda, Goa 403401",
    "Raj Harmony": "Upper Bazar, Ponda, Goa 403401",
    "Ganaraj Arcade": "Upper Bazaar, Near Kamat Nursing Home, Ponda, Goa 403401",
    "Raj Aguiar Enclave": "Dhawali, Ponda, Goa 403401",
    "Raj Madhurangan": "Mardol-Mangeshi Area, Ponda, Goa 403404",
    "Raj Exellency": "Ribandar, North Goa 403006",
    "Rajdeep Galleria": "Opposite Municipality, Sadar, Ponda, Goa 403401",
    "Sairaj Residency": "Porta Vaddo, Siolim, Bardez, North Goa 403517",
    "Raj Braganca Residency": "Ekta Nagar, Mapusa, North Goa 403507",
    "Raj Krishna Towers": "Opposite Kamat Nursing Home, Shanti Nagar, Ponda, Goa 403401",
    "Raj Mudra": "Ponda, South Goa 403401",
    "Raj Vatika": "Near Airport Road, Dabolim, Vasco, South Goa 403801",
    "Sairaj Park": "26, Shanti Nagar, Ponda, Goa 403401"
}

for index, row in df.iterrows():
    proj = str(row['Project Name']).strip()
    if proj in exact_locations:
        df.at[index, 'Location'] = exact_locations[proj]

PROXIMITY_MAP = {
    "sanquelim": "mapusa", "sankhali": "mapusa", "sankhalim": "mapusa", 
    "bicholim": "mapusa", "porvorim": "mapusa", "calangute": "mapusa", 
    "baga": "mapusa", "candolim": "mapusa", "anjuna": "mapusa", 
    "vagator": "mapusa", "saligao": "mapusa", "assagao": "mapusa", 
    "pernem": "mapusa", "mopa": "mapusa", "arambol": "mapusa", 
    "morjim": "mapusa", "mandrem": "mapusa", "colvale": "mapusa", 
    "aldona": "mapusa", "thivim": "mapusa", "tivim": "mapusa",
    "valpoi": "ponda", "sattari": "ponda", "honda": "ponda", 
    "usgao": "ponda", "khandepar": "ponda", "marcel": "ponda", 
    "banastarim": "ponda", "priol": "ponda", "farmagudi": "ponda", 
    "shiroda": "ponda", "borim": "ponda", "sanguem": "ponda", 
    "quepem": "ponda", "curchorem": "ponda", "sanvordem": "ponda",
    "margao": "ponda", "fatorda": "ponda", "colva": "ponda", 
    "benaulim": "ponda", "navelim": "ponda", "canacona": "ponda", 
    "palolem": "ponda", "agonda": "ponda", "corlim": "ponda",
    "cavelossim": "ponda",
    "panjim": "ribandar", "panaji": "ribandar", "miramar": "ribandar", 
    "dona paula": "ribandar", "bambolim": "ribandar", "taleigao": "ribandar", 
    "old goa": "ribandar", "karmali": "ribandar",
    "verna": "dabolim", "majorda": "dabolim", "chicalim": "vasco", 
    "sancoale": "vasco", "bogmalo": "vasco", "zuari": "vasco", "cortalim": "dabolim"
}

RERA_GUIDE = """
---
### 📘 What Does a RERA Number Mean?
1. **🏛️ Government Verified:** The builder cannot sell units without clear land titles.
2. **💰 Buyer Funds Protected:** 70% of buyer payments go into an escrow account.
3. **⏱️ Delivery Guarantee:** Builders face penalties for delayed possession.
4. **📌 Pre-RERA / Older Projects:** Projects completed before 2017 were built under municipal permissions before RERA existed.
---
"""

def search_engine(query_text, drop_type, drop_loc):
    q = query_text.lower() if query_text else ""
    if drop_type and drop_type != "Any": q += f" {drop_type.lower()}"
    if drop_loc and drop_loc != "Any Location": q += f" {drop_loc.lower()}"
    
    if not q.strip():
        return "⚠️ **Please enter your search requirement or select options from the dropdowns above.**"
    
    is_villa = any(w in q for w in ["villa", "bungalow", "bunglow", "row house", "rowhouse"])
    is_commercial = any(w in q for w in ["commercial", "shop", "office"])
    is_flat = any(w in q for w in ["flat", "apartment", "residential", "living", "bhk"])
    
    bhk_match = re.search(r'\b([1-5])\s*bhk\b', q)
    req_bhk = f"{bhk_match.group(1)} BHK" if bhk_match else None
    
    loc_list = ["mapusa", "ponda", "sadar", "siolim", "vasco", "dabolim", "ribandar", "panjim", "karaswada", "goa"]
    req_loc = next((loc for loc in loc_list if loc in q), None)
    if req_loc == "goa": req_loc = None
        
    original_search_area = None
    mapped_hub = None
    
    if not req_loc:
        for town, hub in PROXIMITY_MAP.items():
            if town in q:
                original_search_area = town.title()
                mapped_hub = hub.title()
                req_loc = hub 
                break
    
    has_villa_in_loc = False
    if is_villa and req_loc:
        for _, row in df.iterrows():
            c = str(row['Property Category']).lower()
            t = str(row['Typology']).lower()
            l = str(row['Location']).lower()
            if any(w in c or w in t for w in ["villa", "bungalow", "row house"]) and (req_loc in l):
                has_villa_in_loc = True
                break
                
    results = []
    for _, row in df.iterrows():
        score = 0
        reasons = []
        diffs = []
        
        is_ongoing = str(row['Status']).strip().lower() == 'ongoing'
        c = str(row['Property Category']).lower()
        t = str(row['Typology']).lower()
        l = str(row['Location']).lower()
        
        has_v = any(w in c or w in t for w in ["villa", "bungalow", "row house"])
        has_comm = any(w in c or w in t for w in ["commercial", "shop", "office"])
        has_fl = any(w in c or w in t for w in ["residential", "apartment", "flat", "bhk"])
        
        loc_ok = False
        if req_loc:
            loc_ok = (req_loc in l) or (req_loc == "mapusa" and "karaswada" in l)
            
        if is_ongoing:
            score += 100
            reasons.append("🟢 **PRIORITY:** Active Builder Project (Direct Sales Available)")
        else:
            score += 20
            diffs.append("🟡 **100% Delivered Project:** Direct builder bookings closed. Contact sales desk to inquire about resale availability.")

        if is_villa:
            if has_v:
                score += 50 if (req_loc is None or loc_ok) else 20
                reasons.append("🏠 Matches Villa / Bungalow property type")
                if req_loc and loc_ok and original_search_area:
                    reasons.append(f"📍 Geographically closest hub to {original_search_area}")
                elif req_loc and not loc_ok:
                    diffs.append(f"Located in {row['Location']}, not in requested location.")
            else:
                if req_loc and loc_ok and has_fl:
                    score += 10
                    reasons.append(f"📍 Geographically closest hub to {original_search_area}" if original_search_area else "📍 Located in preferred area")
                    diffs.append("Apartment / Flat option, not an independent Villa")
        else:
            if req_loc and loc_ok:
                score += 40
                reasons.append(f"📍 Geographically closest hub to {original_search_area}" if original_search_area else "📍 Location matches")
            if req_bhk and (req_bhk.lower() in t or req_bhk.lower() in str(row['BHK / Unit Type']).lower()):
                score += 40
                reasons.append(f"🛏️ Matches {req_bhk}")
            if is_commercial and has_comm:
                score += 30
                reasons.append("🏢 Matches Commercial requirement")
            elif (is_flat or "living" in q) and has_fl:
                score += 30
                reasons.append("🏠 Matches Residential requirement")
                
        matched_intent = False
        if is_villa and has_v: matched_intent = True
        elif is_commercial and has_comm: matched_intent = True
        elif is_flat and has_fl: matched_intent = True
        elif req_loc and loc_ok: matched_intent = True
        
        if matched_intent:
            item = row.to_dict()
            item['Score'] = score
            item['Why'] = " | ".join(reasons) if reasons else "Relevant option"
            item['Diff'] = " | ".join(diffs) if diffs else "None"
            item['Is_Ongoing'] = is_ongoing
            results.append(item)
            
    sorted_res = sorted(results, key=lambda x: x['Score'], reverse=True)
    
    out = ""
    if original_search_area and mapped_hub:
        out += f"> ### 📍 Area Proximity Notice\n> **Raj Housing currently does not have properties directly in {original_search_area}.**\n> Showing best verified options in nearby **{mapped_hub}** based on geographic distance.\n---\n\n"
        
    if is_villa and req_loc and not has_villa_in_loc:
        target_name = original_search_area if original_search_area else req_loc.title()
        out += f"> ### ℹ️ Notice: No Active Villas in {target_name}\n> **Raj Housing currently does not have an active/ongoing Villa development in this specific area.**\n> * **Note on Completed Villas:** Raj Housing's past villa projects are 100% delivered. Contact sales desk to check for resale.\n> * **Nearby Options:** Below are alternate residential options in the area.\n---\n\n"
    
    if not sorted_res:
        return "### 🔎 No Matching Property Found"
        
    out += "## 🏠 Verified Property Results\n\n"
    
    for r in sorted_res[:5]:
        if r['Is_Ongoing'] and r['Score'] >= 130:
            badge = "⭐ **BEST MATCH (ACTIVE BOOKING)**"
        elif r['Is_Ongoing']:
            badge = "👍 **GOOD OPTION (ACTIVE BOOKING)**"
        else:
            badge = "📌 **COMPLETED (INQUIRE FOR RESALE)**"
            
        raw_rera = str(r['RERA Number'])
        rera_disp = f"`{raw_rera}` *(Govt Registered)*" if "PRGO" in raw_rera else "Pre-RERA / Older Project"
        diff_line = f"\n- ⚠️ **Things to Consider:** {r['Diff']}" if r['Diff'] != "None" else ""
        
        map_query = f"{r['Project Name']}, {r['Location']}"
        encoded_query = urllib.parse.quote_plus(map_query)
        maps_link = f"https://www.google.com/maps/search/?api=1&query={encoded_query}"
        
        all_photos_str = str(r.get('All Project Photos', ''))
        # FIXED BUG HERE: Corrected split parenthesis
        photo_urls = [p.strip() for p in all_photos_str.split('|') if p.strip().startswith('http')]
        
        gallery_html = ""
        if photo_urls:
            gallery_html = "<div style='display:flex; overflow-x:auto; gap:10px; padding-bottom:10px; margin-bottom:12px;'>"
            for p_url in photo_urls:
                gallery_html += f"<img src='{p_url}' style='min-width:260px; height:200px; object-fit:cover; border-radius:8px; box-shadow:0 3px 8px rgba(0,0,0,0.15);'>"
            gallery_html += "</div>"
        
        out += f"""---
### 🏢 {r['Project Name']} — {badge}

{gallery_html}

- 📍 **Exact Location:** {r['Location']}
- 🏠 **Category / Typology:** {r['Property Category']} ({r['Typology']})
- 🏗️ **Project Status:** {r['Status']}
- ✅ **Why it matches:** {r['Why']}{diff_line}
- 📝 **RERA Registration:** {rera_disp}
- 📞 **Sales Contact:** {r['Sales Contact']}
- 🔗 **Official Webpage:** [View Verified Details]({r['Official Source']})

### [🧭 TAP HERE TO OPEN IN GOOGLE MAPS]({maps_link})

"""
    out += RERA_GUIDE
    return out

with gr.Blocks(theme=gr.themes.Soft(), title="Raj Housing AI Property Finder") as app:
    gr.Markdown("# 🏠 Raj Housing AI — Smart Property Finder")
    
    with gr.Row():
        txt_in = gr.Textbox(label="🔍 Search Properties", placeholder="e.g. '2 BHK in Sanquelim' or 'Villa in Porvorim'", lines=1, scale=2)
    
    with gr.Row():
        drop_type = gr.Dropdown(choices=["Any", "1 BHK", "2 BHK", "3 BHK", "Villa", "Commercial Shop"], label="🏡 Property Type", value="Any")
        drop_loc = gr.Dropdown(choices=["Any Location", "Mapusa", "Ponda", "Ribandar", "Siolim", "Vasco", "Dabolim"], label="📍 Select Hub", value="Any Location")
        
    btn = gr.Button("Find Matching Properties", variant="primary")
    
    gr.Examples(
        examples=[["2 bhk flat in Sanquelim", "Any", "Any Location"], 
                  ["", "Villa", "Any Location"], 
                  ["Commercial shop in Ponda", "Any", "Any Location"]],
        inputs=[txt_in, drop_type, drop_loc]
    )
    
    box = gr.Markdown()
    
    btn.click(fn=search_engine, inputs=[txt_in, drop_type, drop_loc], outputs=box)
    txt_in.submit(fn=search_engine, inputs=[txt_in, drop_type, drop_loc], outputs=box)
    
    with gr.Accordion("ℹ️ What does a RERA number mean? (Click to view guide)", open=False):
        gr.Markdown(RERA_GUIDE)

# Configured for Render cloud deployment
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7860))
    app.launch(server_name="0.0.0.0", server_port=port, share=False)
