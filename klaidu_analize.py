import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import re
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as ExcelImage
import openai

# Naujas OpenAI klientas
client = openai.OpenAI(api_key=st.secrets["openai_api_key"])

st.set_page_config(page_title="Klaidų analizė", layout="centered")
st.title("\U0001F4CA Klaidų analizė pagal mėnesius")

st.write("Įkelkite Excel failą su stulpeliais **Klientas**, **Užsakovas**, **Sąskaitos faktūros Nr.**, **Klaidos**, **Siuntėjas**")

uploaded_file = st.file_uploader("\U0001F4CE Pasirinkite Excel failą", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    def extract_month(text):
        if isinstance(text, str):
            match = re.search(r'\b(KOVAS|VASARIS|SAUSIS|BALANDIS|GEGU\u017d\u0116|BIR\u017dELIS|LIEPA|RUGPJ\u016aTIS|RUGS\u0116JIS|SPALIS|LAPKRTIS|GRUODIS)\b', text.upper())
            if match:
                return match.group(1).capitalize()
        return "Nežinoma"

    df["Mėnuo"] = df["Klientas"].apply(extract_month)
    df["Yra klaida"] = df["Klaidos"].notna()

    menesiu_tvarka = ["Sausis", "Vasaris", "Kovas", "Balandis", "Gegužė", "Birželis",
                      "Liepa", "Rugpjūtis", "Rugsėjis", "Spalis", "Lapkritis", "Gruodis"]
    visi_menesiai = sorted(df["Mėnuo"].dropna().unique(), key=lambda x: menesiu_tvarka.index(x) if x in menesiu_tvarka else 99)

    pasirinkti_menesiai = st.multiselect("\U0001F4C6 Pasirinkite mėnesius analizei", visi_menesiai, default=visi_menesiai)
    df_filtered = df[df["Mėnuo"].isin(pasirinkti_menesiai)]

    summary = df_filtered.groupby("Mėnuo").agg(
        Sąskaitų_skaičius=("Sąskaitos faktūros Nr.", "nunique"),
        Su_klaidomis=("Yra klaida", "sum")
    ).reset_index()

    summary["Klaidų_procentas"] = (summary["Su_klaidomis"] / summary["Sąskaitų_skaičius"] * 100).round(2)
    summary["Mėnesio_nr"] = summary["Mėnuo"].apply(lambda x: menesiu_tvarka.index(x) if x in menesiu_tvarka else -1)
    summary = summary.sort_values("Mėnesio_nr").drop(columns="Mėnesio_nr")

    max_skaicius = summary["Sąskaitų_skaičius"].max()
    summary["Sąskaitų_procentas"] = (summary["Sąskaitų_skaičius"] / max_skaicius * 100).round(2)

    st.subheader("\U0001F4CB Suvestinė")
    st.dataframe(summary, use_container_width=True)

    def generate_insight(row):
        klaidos = row["Su_klaidomis"]
        procentas = row["Klaidų_procentas"]
        saskaitu = row["Sąskaitų_skaičius"]
        menuo = row["Mėnuo"]

        if klaidos == 0:
            return f"✅ {menuo}: jokių klaidų – puikus rezultatas!"
        elif saskaitu < 15 and procentas >= 15:
            return f"⚠️ {menuo}: nors klaidų tik {klaidos}, jos sudaro {procentas:.2f}% – mažas kiekis padidina procentinę įtaką."
        elif procentas >= 20:
            return f"🔴 {menuo}: didelis klaidų procentas ({procentas:.2f}%) – būtina peržiūrėti procesus."
        elif procentas >= 15:
            return f"🟠 {menuo}: padidėjęs klaidų procentas ({procentas:.2f}%) – verta išsiaiškinti priežastis."
        else:
            return f"🟢 {menuo}: klaidų lygis ({procentas:.2f}%) kontroliuojamas."

    summary["Įžvalga"] = summary.apply(generate_insight, axis=1)

    st.subheader("🔎 Įžvalgos pagal mėnesius")
    st.dataframe(summary[["Mėnuo", "Klaidų_procentas", "Sąskaitų_skaičius", "Sąskaitų_procentas", "Su_klaidomis", "Įžvalga"]],
                 use_container_width=True)

    st.subheader("\U0001F4CA Normalizuotas palyginimas (% nuo maksimumo)")
    fig_export, ax_export = plt.subplots(figsize=(10, 6))
    ax_export.plot(summary["Mėnuo"], summary["Sąskaitų_procentas"], label="Sąskaitų kiekis (%)", color="blue", marker="o")
    ax_export.plot(summary["Mėnuo"], summary["Klaidų_procentas"], label="Klaidų procentas (%)", color="red", marker="o")
    ax_export.set_ylabel("Procentai (%)")
    ax_export.set_xlabel("Mėnuo")
    ax_export.set_ylim(0, 100)
    ax_export.legend()
    ax_export.grid(True)
    plt.title("Sąskaitų kiekis ir klaidų procentas (procentinė išraiška)")
    st.pyplot(fig_export)

    st.subheader("\U0001F4DD Klaidų sąrašas")
    klaidos = df_filtered[df_filtered["Yra klaida"] == True][
        ["Mėnuo", "Užsakovas", "Sąskaitos faktūros Nr.", "Klaidos", "Siuntėjas"]
    ]
    st.dataframe(klaidos.reset_index(drop=True), use_container_width=True)

    # Nustatyti dažniausiai klystantį užsakovą ir siuntėją
    st.subheader("\U0001F50E Dažniausiai klystantys užsakovai ir siuntėjai")
    uzsakovai_stats = klaidos["Užsakovas"].value_counts().reset_index()
    uzsakovai_stats.columns = ["Užsakovas", "Klaidų skaičius"]

    siuntejai_stats = klaidos["Siuntėjas"].value_counts().reset_index()
    siuntejai_stats.columns = ["Siuntėjas", "Klaidų skaičius"]

    st.write("**Dažniausi klientai su klaidinga informacija aktuose:**")
    st.dataframe(uzsakovai_stats, use_container_width=True)

    st.write("**Dažniausi šių aktų siuntėjai:**")
    st.dataframe(siuntejai_stats, use_container_width=True)

    img_buffer = io.BytesIO()
    fig_export.savefig(img_buffer, format="png")
    img_buffer.seek(0)

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Suvestinė"
    for r in dataframe_to_rows(summary, index=False, header=True):
        ws1.append(r)

    ws2 = wb.create_sheet(title="Klaidų sąrašas")
    for r in dataframe_to_rows(klaidos, index=False, header=True):
        ws2.append(r)

    img_path = "dvigubas_grafikas.png"
    with open(img_path, "wb") as f:
        f.write(img_buffer.read())

    img = ExcelImage(img_path)
    img.anchor = "E2"
    ws1.add_image(img)

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    st.download_button(
        label="\U0001F4E5 Atsisiųsti Excel ataskaitą su grafiku",
        data=excel_buffer,
        file_name="Klaidu_Ataskaita.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("\U0001F916 Dirbtinio intelekto analizė")
    try:
        markdown_table = summary.to_markdown(index=False)
        analysis_prompt = (
            "Analizuok pateiktus duomenis apie sąskaitų skaičių ir klaidų procentą pagal mėnesius. "
            "Pateik įžvalgas apie tendencijas, galimas klaidų priežastis ir pateik pasiūlymus, kaip jas sumažinti ateityje.\n\n"
            + markdown_table
        )

        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "Tu esi patyręs verslo analitikas."},
                {"role": "user", "content": analysis_prompt}
            ],
            temperature=0.4
        )

        st.markdown(response.choices[0].message.content)

    except Exception as e:
        st.warning("Nepavyko gauti AI analizės. Patikrink API raktą Streamlit `secrets` nustatymuose.")
        st.error(str(e))
