import streamlit as st
from google import genai
import datetime
import os
import docx
import pandas as pd
import re

# 1. API instellen
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

# Nieuwe titel met een icoontje
st.title("🗺️ Formatieve toets Aardrijkskunde")
st.markdown("Welkom! Vul je gegevens hieronder in om te beginnen met je overhoring.")

# Functie om tekst uit een Word document te halen
def lees_docx(file_path):
    doc = docx.Document(file_path)
    volledige_tekst = []
    for para in doc.paragraphs:
        volledige_tekst.append(para.text)
    return "\n".join(volledige_tekst)

# 2. Zoek alle Word-documenten in de map 'lesmateriaal Havo'
les_map = "lesmateriaal Havo"
if not os.path.exists(les_map):
    os.makedirs(les_map)

beschikbare_bestanden = [f for f in os.listdir(les_map) if f.endswith('.docx')]

# 3. Check of er bestanden zijn, anders instructie om te mailen
if not beschikbare_bestanden:
    st.warning("Er is op dit moment geen lesmateriaal beschikbaar. Stuur even een mailtje naar je docent.")
else:
    # 4. Keuzemenu voor de leerling (nu weer bovenaan de hoofdpagina)
    st.header("📝 Jouw Gegevens")
    gekozen_les = st.selectbox("Kies de les die je wilt oefenen:", beschikbare_bestanden)
    
    cluster_opties = ["4Hak1", "4Hak2", "4Hak3", "4Hak4", "5Hak1", "5Hak2", "5Hak3"]
    cluster = st.selectbox("Kies je cluster:", cluster_opties)
    
    voornaam = st.text_input("Vul je voornaam in om te beginnen:")
    
    st.divider() # Trekt een mooi lijntje tussen de gegevens en de chat

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
            st.session_state.interaction_id = None
            
            les_pad = os.path.join(les_map, gekozen_les)
            les_tekst = ""

            with st.spinner("De docent neemt de theorie door... Een moment geduld aub."):
                try:
                    les_tekst = lees_docx(les_pad)
                except Exception as e:
                    st.error(f"Er ging iets mis met het lezen van het bestand: {e}")

            if les_tekst:
                # 6. Start gesprek met instructies
                eerste_input = f"""
Je bent docent aardrijkskunde (bovenbouw). Toon: professioneel, zakelijk.
Baseer de ONDERWERPEN op de theorie. Geef NOOIT zelf direct het antwoord (behalve als een leerling een vraag definitief fout heeft).

--- START THEORIE ---
{les_tekst}
--- EINDE THEORIE ---

Volg EXACT deze chronologische structuur:

**Fase 1: Intro**
1. Zakelijke groet.
2. Deel een kort, verrassend wist-je-datje over **slim leren, de werking van het brein bij leren of effectieve studiemethodes** (dus NIET over aardrijkskunde).
3. Vraag daarna of het boek dicht is door de leerling deze 3 opties te geven in een lijstje:
   [A] Ik heb de stof bestudeerd en ik ga het helemaal zelf doen.
   [B] Ik heb de stof niet bestudeerd, maar ik ga het gewoon proberen.
   [C] Nee, ik wil stoppen.
   Vraag de leerling expliciet om 'A', 'B' of 'C' te typen. Wacht op het antwoord voor je verdergaat.

**Fase 2: Overhoring (EXACT 5 vragen: 2 reproductie, 3 inzicht)**
- STOPPEN: Kiest de leerling optie C of typt hij "stop"? Breek alles dan af! Zeg UITSLUITEND: "Ga de stof nogmaals bestuderen en probeer het dan nog eens! [EINDE_OVERHORING]"
- CIJFER BIJHOUDEN: Start op 10. Helemaal fout = -2. Gedeeltelijk = -1. Spelfout = -0.5 (max -2 in totaal). Cijfer mag negatief zijn.
- COULANT NAKIJKEN: De leerlingen leren uit een ánder boek. Antwoorden hoeven dus NIET exact overeen te komen met jouw theorie-tekst. Reken het GOED als de leerling in eigen woorden laat zien dat hij/zij het snapt.
- Reproductie: Vraag ALTIJD "Wat betekent [begrip]?".
- 1 vraag tegelijk. Wacht op antwoord.
- VOLLEDIGE ZINNEN: Bij een los woord keur je het nog niet goed. Zeg: "Inhoudelijk juist, maar formuleer het als een volledige zin."
- SPELLING: Corrigeer spelfouten direct en tel de aftrek mee.
- DEELS GOED / IN EIGEN WOORDEN: Reken begrip goed. Indien onvolledig: stel 1 hulpvraag.
- FOUT: De leerling krijgt 1 herkansing per vraag. Wéér fout? Reken fout (-2), geef het goede antwoord en ga door naar de VOLGENDE vraag. Altijd 5 vragen behandelen.

**Fase 3: Afronding**
1. Zodra alle 5 vragen zijn geweest, vraag je EERST hoe de leerling de toets gemaakt heeft met deze 2 opties:
   [A] Ik heb het helemaal op eigen kracht gedaan.
   [B] Ik heb helaas vals moeten spelen om het te halen.
   Vraag de leerling om 'A' of 'B' te typen en wacht op antwoord.
2. Na hun antwoord: Geef gerichte feedback en laat de score-berekening zien.
3. Noteer het cijfer EXACT zo: [CIJFER: X]. Sluit daarna je bericht af met EXACT: [EINDE_OVERHORING].
"""
                interaction = client.interactions.create(
                    model="gemini-3.1-flash-lite",
                    input=eerste_input
                )
                
                st.session_state.interaction_id = interaction.id
                st.session_state.berichten.append(("assistant", interaction.output_text))

        # 7. Weergave van alle berichten in de app MET AVATARS
        if "berichten" in st.session_state:
            for role, text in st.session_state.berichten:
                avatar_icoon = "🧑‍🏫" if role == "assistant" else "🎓"
                
                # Verberg de eind-tags
                weergave_tekst = re.sub(r'\[CIJFER:\s*([\-\d\,\.]+)\]', '', text)
                weergave_tekst = weergave_tekst.replace("[EINDE_OVERHORING]", "")
                
                with st.chat_message(role, avatar=avatar_icoon):
                    st.markdown(weergave_tekst.strip())

        # 8. HET VASTE INVOERVELD
        prompt = st.chat_input("Typ hier je antwoord of keuze...")

        # 9. Invoer verwerken en API aanroepen
        if prompt and "interaction_id" in st.session_state and st.session_state.interaction_id is not None:
            
            st.session_state.berichten.append(("user", prompt))
            with st.chat_message("user", avatar="🎓"):
                st.markdown(prompt)
            
            with st.spinner("De docent schrijft een reactie..."):
                vervolg_interaction = client.interactions.create(
                    model="gemini-3.1-flash-lite",
                    previous_interaction_id=st.session_state.interaction_id,
                    input=prompt
                )
            
            st.session_state.interaction_id = vervolg_interaction.id
            output_tekst = vervolg_interaction.output_text
            
            st.session_state.berichten.append(("assistant", output_tekst))
            
            weergave_tekst_bot = re.sub(r'\[CIJFER:\s*([\-\d\,\.]+)\]', '', output_tekst)
            weergave_tekst_bot = weergave_tekst_bot.replace("[EINDE_OVERHORING]", "")
            
            with st.chat_message("assistant", avatar="🧑‍🏫"):
                st.markdown(weergave_tekst_bot.strip())
                
            # Als de overhoring klaar is, of afgebroken
            if "[EINDE_OVERHORING]" in output_tekst:
                
                cijfer_match = re.search(r'\[CIJFER:\s*([\-\d\,\.]+)\]', output_tekst)
                cijfer = ""
                cijfer_waarde = 0.0 # Standaardwaarde om mee te rekenen
                
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
                
                # Check of er is afgebroken, of dat er een cijfer gehaald is
                if "Ga de stof nogmaals bestuderen" in output_tekst:
                    st.info("De overhoring is afgebroken. Succes met studeren en tot de volgende keer!")
                else:
                    # Ballonnen alléén bij een 6.0 of hoger
                    if cijfer_waarde >= 6.0:
                        st.balloons()
                        st.success("🎉 Goed gewerkt, je hebt een voldoende! Je resultaten zijn opgeslagen. Je kunt dit venster nu sluiten.")
                    else:
                        st.success("✅ Je resultaten zijn opgeslagen. Blijf goed oefenen, volgende keer gaat het vast beter! Je kunt dit venster nu sluiten.")
