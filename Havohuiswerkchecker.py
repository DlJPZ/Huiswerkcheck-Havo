import streamlit as st
from google import genai
import datetime
import os
import docx
import pandas as pd
import re

# 1. API instellen (Let op dat de API key in je st.secrets staat)
client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

# CSS voor de Mexicaanse vlag achtergrond
achtergrond_css = """
<style>
.stApp {
    background-image: linear-gradient(rgba(255, 255, 255, 0.85), rgba(255, 255, 255, 0.85)), url("https://upload.wikimedia.org/wikipedia/commons/f/fc/Flag_of_Mexico.svg");
    background-size: cover;
    background-position: center;
    background-attachment: fixed;
}
</style>
"""
st.markdown(achtergrond_css, unsafe_allow_html=True)

st.title("🗺️ Formatieve toets Aardrijkskunde")
st.markdown("Welkom! Vul je gegevens hieronder in om te beginnen met je overhoring.")

# Functie om tekst uit een Word document te halen
def lees_docx(file_path):
    doc = docx.Document(file_path)
    volledige_tekst = [para.text for para in doc.paragraphs]
    return "\n".join(volledige_tekst)

# 2. Zoek alle Word-documenten in de map 'lesmateriaal Havo'
les_map = "lesmateriaal Havo"
if not os.path.exists(les_map):
    os.makedirs(les_map)

beschikbare_bestanden = [f for f in os.listdir(les_map) if f.endswith('.docx')]

# 3. Check of er bestanden zijn
if not beschikbare_bestanden:
    st.warning("Er is op dit moment geen lesmateriaal beschikbaar. Stuur even een mailtje naar je docent.")
else:
    # 4. Keuzemenu voor de leerling
    st.header("📝 Jouw Gegevens")
    gekozen_les = st.selectbox("Kies de les die je wilt oefenen:", beschikbare_bestanden)
    
    cluster_opties = ["4Hak1", "4Hak2", "4Hak3", "4Hak4", "5Hak1", "5Hak2", "5Hak3"]
    cluster = st.selectbox("Kies je cluster:", cluster_opties)
    
    voornaam = st.text_input("Vul je voornaam in om te beginnen:")
    
    st.divider()

    if voornaam and cluster and gekozen_les:
        
        # 5. Start of reset de les
        if ("huidige_les" not in st.session_state or 
            st.session_state.huidige_les != gekozen_les or 
            st.session_state.get("actieve_voornaam") != voornaam or
            st.session_state.get("actief_cluster") != cluster):
            
            st.session_state.huidige_les = gekozen_les
            st.session_state.actieve_voornaam = voornaam
            st.session_state.actief_cluster = cluster
            st.session_state.berichten = [] 
            st.session_state.chat = None
            
            les_pad = os.path.join(les_map, gekozen_les)
            les_tekst = ""

            with st.spinner("De docent neemt de theorie door... Een moment geduld aub."):
                try:
                    les_tekst = lees_docx(les_pad)
                except Exception as e:
                    st.error(f"Er ging iets mis met het lezen van het bestand: {e}")

            if les_tekst:
                eerste_input = f"""
Je bent docent aardrijkskunde (bovenbouw). Toon: professioneel, zakelijk, maar wel aanmoedigend.
Baseer de ONDERWERPEN op de theorie. Geef NOOIT zelf direct het antwoord (behalve als een leerling een vraag definitief fout heeft).

--- START THEORIE ---
{les_tekst}
--- EINDE THEORIE ---

Volg EXACT deze chronologische structuur:

**Fase 1: Intro**
1. Zakelijke groet.
2. Geef een duidelijke waarschuwing: "Let op: let goed op je spelling, want spelfouten leiden tot puntaftrek!"
3. Vraag daarna of het boek dicht is door de leerling deze 3 opties te geven in een lijstje:
   [A] Ik heb de stof bestudeerd en ik ga het helemaal zelf doen.
   [B] Ik heb de stof niet bestudeerd, maar ik ga het gewoon proberen.
   [C] Nee, ik wil stoppen.
   Vraag de leerling expliciet om 'A', 'B' of 'C' te typen. Wacht op het antwoord voor je verdergaat.

**Fase 2: Overhoring (EXACT 5 vragen: 2 reproductie, 3 inzicht)**
- STOPPEN: Kiest de leerling optie C of typt hij "stop"? Breek alles dan af! Zeg UITSLUITEND: "Ga de stof nogmaals bestuderen en probeer het dan nog eens! [EINDE_OVERHORING]"
- CIJFER BIJHOUDEN: Start op 10. Helemaal fout = -2. Spelfout = -0.5 (max -2 aftrek voor spelling in totaal). Cijfer mag negatief zijn.
- COULANT NAKIJKEN (HUISWERKCONTROLE): Dit is een huiswerkcontrole, geen formele toets. Reken een antwoord GOED (geen aftrek) als de leerling laat zien dat hij/zij het snapt, zelfs als het antwoord niet helemáál volledig is. Geef in dat geval wél direct als feedback wat het volledige antwoord had moeten zijn, en ga daarna door naar de volgende vraag.
- ZINSBOUW: Een antwoord is qua formulering akkoord zolang het minimaal een onderwerp en één of meerdere werkwoorden bevat. Als dit ontbreekt (bijv. de leerling typt slechts één los woord), keur je het nog niet direct fout, maar zeg je: "Inhoudelijk zit je in de goede richting, maar formuleer je antwoord even in een zin met minimaal een onderwerp en een werkwoord."
- Reproductie: Vraag ALTIJD "Wat betekent [begrip]?".
- 1 vraag tegelijk. Wacht op antwoord.
- SPELLING: Corrigeer spelfouten direct, benoem ze kort, en tel de aftrek mee.
- FOUT: Is het antwoord echt onjuist? De leerling krijgt 1 herkansing per vraag. Wéér fout? Reken fout (-2), geef het goede antwoord en ga door naar de VOLGENDE vraag. Altijd 5 vragen behandelen.

**Fase 3: Afronding**
1. Zodra alle 5 vragen zijn geweest, vraag je EERST hoe de leerling de toets gemaakt heeft met deze 2 opties:
   [A] Ik heb het helemaal op eigen kracht gedaan.
   [B] Ik heb helaas vals moeten spelen om het te halen.
   Vraag de leerling om 'A' of 'B' te typen en wacht op antwoord.
2. Na hun antwoord: Geef gerichte feedback en laat de score-berekening zien.
3. Noteer het cijfer EXACT zo: [CIJFER: X]. Sluit daarna je bericht af met EXACT: [EINDE_OVERHORING].
"""
                # Start een chatsessie op de correcte manier
                chat = client.chats.create(model="gemini-1.5-flash")
                st.session_state.chat = chat
                
                response = chat.send_message(eerste_input)
                st.session_state.berichten.append(("assistant", response.text))

        # 7. Weergave van alle berichten in de app
        if "berichten" in st.session_state:
            for role, text in st.session_state.berichten:
                avatar_icoon = "🧑‍🏫" if role == "assistant" else "🎓"
                
                # Verberg de API-sturings elementen voor de leerling
                weergave_tekst = re.sub(r'\[CIJFER:\s*([\-\d\,\.]+)\]', '', text)
                weergave_tekst = weergave_tekst.replace("[EINDE_OVERHORING]", "")
                
                with st.chat_message(role, avatar=avatar_icoon):
                    st.markdown(weergave_tekst.strip())

        # 8. HET VASTE INVOERVELD
        prompt = st.chat_input("Typ hier je antwoord of keuze...")

        # 9. Invoer verwerken en API aanroepen
        if prompt and "chat" in st.session_state and st.session_state.chat is not None:
            
            st.session_state.berichten.append(("user", prompt))
            with st.chat_message("user", avatar="🎓"):
                st.markdown(prompt)
            
            with st.spinner("De docent schrijft een reactie..."):
                vervolg_response = st.session_state.chat.send_message(prompt)
                output_tekst = vervolg_response.text
            
            st.session_state.berichten.append(("assistant", output_tekst))
            
            weergave_tekst_bot = re.sub(r'\[CIJFER:\s*([\-\d\,\.]+)\]', '', output_tekst)
            weergave_tekst_bot = weergave_tekst_bot.replace("[EINDE_OVERHORING]", "")
            
            with st.chat_message("assistant", avatar="🧑‍🏫"):
                st.markdown(weergave_tekst_bot.strip())
                
            # Als de overhoring klaar is, of afgebroken, sla het resultaat op
            if "[EINDE_OVERHORING]" in output_tekst:
                cijfer_match = re.search(r'\[CIJFER:\s*([\-\d\,\.]+)\]', output_tekst)
                cijfer = ""
                cijfer_waarde = 0.0
                
                if cijfer_match:
                    cijfer_str = cijfer_match.group(1).replace(',', '.') 
                    try:
                        cijfer = float(cijfer_str)
                        cijfer_waarde = cijfer
                    except ValueError:
                        cijfer = cijfer_str
                
                schone_beoordeling = weergave_tekst_bot.strip()
                excel_bestand = "leerling_resultaten_Havo.xlsx"
                
                nieuw_resultaat = pd.DataFrame([{
                    "Tijdstip": datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
                    "Voornaam": voornaam,
                    "Cluster": cluster,
                    "Les": gekozen_les,
                    "Cijfer": cijfer,
                    "Beoordeling (Feedback AI)": schone_beoordeling
                }])
                
                if os.path.exists(excel_bestand):
                    try:
                        df_bestaand = pd.read_excel(excel_bestand)
                        df_compleet = pd.concat([df_bestaand, nieuw_resultaat], ignore_index=True)
                    except Exception:
                        df_compleet = nieuw_resultaat
                else:
                    df_compleet = nieuw_resultaat
                    
                df_compleet.to_excel(excel_bestand, index=False)
                
                if "Ga de stof nogmaals bestuderen" in output_tekst:
                    st.info("De overhoring is afgebroken. Succes met studeren en tot de volgende keer!")
                else:
                    if cijfer_waarde >= 6.0:
                        st.balloons()
                        st.success("🎉 Goed gewerkt, je hebt een voldoende! Je resultaten zijn opgeslagen. Je kunt dit venster nu sluiten.")
                    else:
                        st.success("✅ Je resultaten zijn opgeslagen. Blijf goed oefenen, volgende keer gaat het vast beter! Je kunt dit venster nu sluiten.")

# --- DOCENTENPANEEL (VOLLEDIG BEVEILIGD) ---
st.sidebar.divider()
st.sidebar.header("👨‍🏫 Docentenpaneel")

wachtwoord = st.sidebar.text_input("Wachtwoord docent:", type="password")

# Alles hieronder is onzichtbaar totdat het juiste wachtwoord is ingevoerd
if wachtwoord == "M@@rt3n": 
    
    excel_bestand = "leerling_resultaten_Havo.xlsx"

    # Statistieken & Downloads
    if os.path.exists(excel_bestand):
        df = pd.read_excel(excel_bestand)
        df['Cijfer'] = pd.to_numeric(df['Cijfer'], errors='coerce')
        
        st.sidebar.subheader("📊 Live Statistieken")
        gemiddelde = df['Cijfer'].mean()
        st.sidebar.metric(label="Gemiddeld Cijfer (Alle clusters)", value=f"{gemiddelde:.1f}")
        
        st.sidebar.write("**Aantal deelnames per cluster:**")
        st.sidebar.dataframe(df['Cluster'].value_counts(), use_container_width=True)
        
        st.sidebar.divider()
        
        st.sidebar.subheader("📥 Resultaten Exporteren")
        with open(excel_bestand, "rb") as file:
            st.sidebar.download_button(
                label="Download Resultaten (Excel)",
                data=file,
                file_name=f"Resultaten_Aardrijkskunde_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.sidebar.info("Er zijn nog geen resultaten opgeslagen.")

    st.sidebar.divider()

    # Bestand Upload (Nu ook veilig achter het wachtwoord)
    st.sidebar.subheader("📄 Nieuwe les uploaden")
    st.sidebar.write("Voeg direct een nieuw Word-document toe aan het keuzemenu.")
    
    uploaded_file = st.sidebar.file_uploader("Kies een .docx bestand", type=["docx"])
    
    if uploaded_file is not None:
        file_path = os.path.join(les_map, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        st.sidebar.success(f"✅ '{uploaded_file.name}' is geüpload!")
        
        if st.sidebar.button("Vernieuw app om les te tonen"):
            st.rerun()

elif wachtwoord != "":
    st.sidebar.error("Onjuist wachtwoord.")
