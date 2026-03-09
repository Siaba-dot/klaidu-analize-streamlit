import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import re
import io
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as ExcelImage
import openai

# OpenAI klientas
client = openai.OpenAI(api_key=st.secrets["openai_api_key"])

st.set_page_config(page_title="Klaidų analizė", layout="wide")
st.title("📊 Klaidų analizė pagal mėnesius")
st.write(
    "Įkelkite Excel failą. "
    "Analizėje naudojama:\n"
    "- **Klaidos priežastis** iš **O** stulpelio\n"
    "- **Klaidos** iš **P** stulpelio"
)

uploaded_file = st.file_uploader("📎 Pasirinkite Excel failą", type=["xlsx"])


if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # ----------------------------
    # Pagalbinės funkcijos
    # ----------------------------
    def extract_month(text):
        if isinstance(text, str):
            match = re.search(
                r"\b(KOVAS|VASARIS|SAUSIS|BALANDIS|GEGUŽĖ|BIRŽELIS|LIEPA|RUGPJŪTIS|RUGSĖJIS|SPALIS|LAPKRITIS|GRUODIS)\b",
                text.upper()
            )
            if match:
                return match.group(1).capitalize()
        return "Nežinoma"

    def clean_text(value):
        if pd.isna(value):
            return None
        text = str(value).strip()
        return text if text else None

    def generate_insight(row):
        klaidos = row["Su_klaidomis"]
        procentas = row["Klaidų_procentas"]
        saskaitu = row["Sąskaitų_skaičius"]
        menuo = row["Mėnuo"]

        if klaidos == 0:
            return f"✅ {menuo}: jokių klaidų – puikus rezultatas."
        elif saskaitu < 15 and procentas >= 15:
            return (
                f"⚠️ {menuo}: nors klaidų tik {klaidos}, jos sudaro {procentas:.2f}% – "
                "mažas kiekis padidina procentinę įtaką."
            )
        elif procentas >= 20:
            return f"🔴 {menuo}: labai aukštas klaidų procentas ({procentas:.2f}%) – būtina peržiūrėti procesą."
        elif procentas >= 15:
            return f"🟠 {menuo}: padidėjęs klaidų procentas ({procentas:.2f}%) – verta ieškoti priežasčių."
        else:
            return f"🟢 {menuo}: klaidų lygis ({procentas:.2f}%) yra kontroliuojamas."

    def safe_add_image(ws, image_path, anchor):
        if os.path.exists(image_path):
            img = ExcelImage(image_path)
            img.anchor = anchor
            ws.add_image(img)

    # ----------------------------
    # Privalomų stulpelių patikra
    # ----------------------------
    required_named_columns = ["Klientas", "Užsakovas", "Sąskaitos faktūros Nr.", "Siuntėjas"]
    missing_named = [col for col in required_named_columns if col not in df.columns]

    if missing_named:
        st.error(f"Faile trūksta šių stulpelių: {', '.join(missing_named)}")
        st.stop()

    if df.shape[1] < 16:
        st.error("Faile nepakanka stulpelių. Reikia bent 16 stulpelių, kad būtų galima paimti O ir P.")
        st.stop()

    # ----------------------------
    # Nauji stulpeliai iš O ir P
    # O = 15-as indeksas? Ne. Kadangi indeksai nuo 0:
    # O -> 14, P -> 15
    # ----------------------------
    df["Klaidos_priežastis"] = df.iloc[:, 14].apply(clean_text)  # O
    df["Klaidos"] = df.iloc[:, 15].apply(clean_text)             # P

    # Mėnuo iš Kliento stulpelio
    df["Mėnuo"] = df["Klientas"].apply(extract_month)

    # Yra klaida, jei P stulpelyje kažkas yra
    df["Yra klaida"] = df["Klaidos"].notna()

    menesiu_tvarka = [
        "Sausis", "Vasaris", "Kovas", "Balandis", "Gegužė", "Birželis",
        "Liepa", "Rugpjūtis", "Rugsėjis", "Spalis", "Lapkritis", "Gruodis"
    ]

    visi_menesiai = sorted(
        df["Mėnuo"].dropna().unique(),
        key=lambda x: menesiu_tvarka.index(x) if x in menesiu_tvarka else 99
    )

    pasirinkti_menesiai = st.multiselect(
        "📆 Pasirinkite mėnesius analizei",
        visi_menesiai,
        default=visi_menesiai
    )

    df_filtered = df[df["Mėnuo"].isin(pasirinkti_menesiai)].copy()

    # ----------------------------
    # Mėnesinė suvestinė
    # ----------------------------
    summary = df_filtered.groupby("Mėnuo").agg(
        Sąskaitų_skaičius=("Sąskaitos faktūros Nr.", "nunique"),
        Su_klaidomis=("Yra klaida", "sum")
    ).reset_index()

    summary["Klaidų_procentas"] = (
        summary["Su_klaidomis"] / summary["Sąskaitų_skaičius"] * 100
    ).round(2)

    summary["Mėnesio_nr"] = summary["Mėnuo"].apply(
        lambda x: menesiu_tvarka.index(x) if x in menesiu_tvarka else -1
    )
    summary = summary.sort_values("Mėnesio_nr").drop(columns="Mėnesio_nr")

    max_skaicius = summary["Sąskaitų_skaičius"].max() if not summary.empty else 1
    summary["Sąskaitų_procentas"] = (
        summary["Sąskaitų_skaičius"] / max_skaicius * 100
    ).round(2)

    summary["Įžvalga"] = summary.apply(generate_insight, axis=1)

    # ----------------------------
    # Klaidų sąrašas
    # ----------------------------
    klaidos = df_filtered[df_filtered["Yra klaida"] == True][[
        "Mėnuo",
        "Užsakovas",
        "Sąskaitos faktūros Nr.",
        "Klaidos_priežastis",
        "Klaidos",
        "Siuntėjas"
    ]].copy()

    # ----------------------------
    # KPI kortelės
    # ----------------------------
    viso_dokumentu = int(df_filtered["Sąskaitos faktūros Nr."].nunique())
    viso_klaidu = int(df_filtered["Yra klaida"].sum())
    klaidu_proc = round((viso_klaidu / viso_dokumentu * 100), 2) if viso_dokumentu else 0.0
    be_klaidu = viso_dokumentu - viso_klaidu

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Viso sąskaitų", f"{viso_dokumentu}")
    col2.metric("Su klaidomis", f"{viso_klaidu}")
    col3.metric("Klaidų %", f"{klaidu_proc}%")
    col4.metric("Be klaidų", f"{be_klaidu}")

    # ----------------------------
    # Suvestinė
    # ----------------------------
    st.subheader("📋 Suvestinė pagal mėnesius")
    st.dataframe(summary, use_container_width=True)

    st.subheader("🔎 Įžvalgos pagal mėnesius")
    st.dataframe(
        summary[[
            "Mėnuo",
            "Klaidų_procentas",
            "Sąskaitų_skaičius",
            "Sąskaitų_procentas",
            "Su_klaidomis",
            "Įžvalga"
        ]],
        use_container_width=True
    )

    # ----------------------------
    # 1 grafikas: normalizuotas palyginimas
    # ----------------------------
    st.subheader("📈 Normalizuotas palyginimas (% nuo maksimumo)")

    fig_export, ax_export = plt.subplots(figsize=(10, 6))
    ax_export.plot(
        summary["Mėnuo"],
        summary["Sąskaitų_procentas"],
        label="Sąskaitų kiekis (%)",
        marker="o"
    )
    ax_export.plot(
        summary["Mėnuo"],
        summary["Klaidų_procentas"],
        label="Klaidų procentas (%)",
        marker="o"
    )
    ax_export.set_ylabel("Procentai (%)")
    ax_export.set_xlabel("Mėnuo")
    ax_export.set_ylim(0, 100)
    ax_export.legend()
    ax_export.grid(True)
    plt.title("Sąskaitų kiekis ir klaidų procentas")
    st.pyplot(fig_export)

    # ----------------------------
    # Klaidų sąrašas
    # ----------------------------
    st.subheader("📝 Klaidų sąrašas")
    st.dataframe(klaidos.reset_index(drop=True), use_container_width=True)

    # ----------------------------
    # Analizė pagal klaidos priežastį
    # ----------------------------
    st.subheader("📌 Klaidos pagal priežastį")

    if not klaidos.empty:
        priezastys = klaidos["Klaidos_priežastis"].fillna("Nenurodyta").value_counts().reset_index()
        priezastys.columns = ["Klaidos priežastis", "Klaidų skaičius"]

        c1, c2 = st.columns([1, 1])

        with c1:
            st.dataframe(priezastys, use_container_width=True)

        with c2:
            fig_reason, ax_reason = plt.subplots(figsize=(8, 5))
            ax_reason.barh(priezastys["Klaidos priežastis"], priezastys["Klaidų skaičius"])
            ax_reason.set_title("Klaidos pagal priežastį")
            ax_reason.set_xlabel("Klaidų skaičius")
            ax_reason.invert_yaxis()
            ax_reason.grid(axis="x")
            st.pyplot(fig_reason)
    else:
        priezastys = pd.DataFrame(columns=["Klaidos priežastis", "Klaidų skaičius"])
        st.info("Pasirinktuose mėnesiuose klaidų nėra.")

    # ----------------------------
    # Analizė pagal užsakovą ir siuntėją
    # ----------------------------
    st.subheader("👥 Dažniausi klaidų šaltiniai")

    if not klaidos.empty:
        uzsakovai_stats = klaidos["Užsakovas"].fillna("Nenurodyta").value_counts().reset_index()
        uzsakovai_stats.columns = ["Užsakovas", "Klaidų skaičius"]

        siuntejai_stats = klaidos["Siuntėjas"].fillna("Nenurodyta").value_counts().reset_index()
        siuntejai_stats.columns = ["Siuntėjas", "Klaidų skaičius"]

        c1, c2 = st.columns(2)

        with c1:
            st.write("**Dažniausi klientai / užsakovai su klaidomis:**")
            st.dataframe(uzsakovai_stats, use_container_width=True)

        with c2:
            st.write("**Dažniausi šių aktų siuntėjai:**")
            st.dataframe(siuntejai_stats, use_container_width=True)

        # Siuntėjų grafikas
        st.subheader("📨 Klaidos pagal siuntėją")
        top_siuntejai = siuntejai_stats.head(10)

        fig_sender, ax_sender = plt.subplots(figsize=(9, 5))
        ax_sender.barh(top_siuntejai["Siuntėjas"], top_siuntejai["Klaidų skaičius"])
        ax_sender.set_title("TOP siuntėjai pagal klaidų kiekį")
        ax_sender.set_xlabel("Klaidų skaičius")
        ax_sender.invert_yaxis()
        ax_sender.grid(axis="x")
        st.pyplot(fig_sender)

        # Pareto analizė
        st.subheader("📊 Pareto analizė – kur koncentruojasi daugiausia klaidų")

        pareto = siuntejai_stats.copy()
        pareto["Kumuliacinis %"] = (
            pareto["Klaidų skaičius"].cumsum() / pareto["Klaidų skaičius"].sum() * 100
        ).round(2)

        fig_pareto, ax1 = plt.subplots(figsize=(10, 6))
        ax1.bar(pareto["Siuntėjas"], pareto["Klaidų skaičius"])
        ax1.set_ylabel("Klaidų skaičius")
        ax1.set_xlabel("Siuntėjas")
        ax1.grid(axis="y")

        ax2 = ax1.twinx()
        ax2.plot(pareto["Siuntėjas"], pareto["Kumuliacinis %"], color="red", marker="o")
        ax2.set_ylabel("Kumuliacinis %")
        ax2.set_ylim(0, 110)

        plt.title("Pareto analizė – klaidos pagal siuntėją")
        plt.xticks(rotation=45, ha="right")
        st.pyplot(fig_pareto)
    else:
        uzsakovai_stats = pd.DataFrame(columns=["Užsakovas", "Klaidų skaičius"])
        siuntejai_stats = pd.DataFrame(columns=["Siuntėjas", "Klaidų skaičius"])
        pareto = pd.DataFrame(columns=["Siuntėjas", "Klaidų skaičius", "Kumuliacinis %"])

    # ----------------------------
    # Excel eksporto paruošimas
    # ----------------------------
    # Saugome grafikus į laikinas bylas
    img_summary_path = "grafikas_menesiai.png"
    img_reason_path = "grafikas_priezastys.png"
    img_sender_path = "grafikas_siuntejai.png"
    img_pareto_path = "grafikas_pareto.png"

    fig_export.savefig(img_summary_path, bbox_inches="tight")

    if not klaidos.empty:
        fig_reason.savefig(img_reason_path, bbox_inches="tight")
        fig_sender.savefig(img_sender_path, bbox_inches="tight")
        fig_pareto.savefig(img_pareto_path, bbox_inches="tight")

    wb = Workbook()

    # 1 lapas: Suvestinė
    ws1 = wb.active
    ws1.title = "Suvestinė"
    for r in dataframe_to_rows(summary, index=False, header=True):
        ws1.append(r)
    safe_add_image(ws1, img_summary_path, "I2")

    # 2 lapas: Klaidų sąrašas
    ws2 = wb.create_sheet(title="Klaidų sąrašas")
    for r in dataframe_to_rows(klaidos, index=False, header=True):
        ws2.append(r)

    # 3 lapas: Priežastys
    ws3 = wb.create_sheet(title="Priežastys")
    for r in dataframe_to_rows(priezastys, index=False, header=True):
        ws3.append(r)
    safe_add_image(ws3, img_reason_path, "D2")

    # 4 lapas: Siuntėjai
    ws4 = wb.create_sheet(title="Siuntėjai")
    for r in dataframe_to_rows(siuntejai_stats, index=False, header=True):
        ws4.append(r)
    safe_add_image(ws4, img_sender_path, "D2")

    # 5 lapas: Pareto
    ws5 = wb.create_sheet(title="Pareto")
    for r in dataframe_to_rows(pareto, index=False, header=True):
        ws5.append(r)
    safe_add_image(ws5, img_pareto_path, "E2")

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    st.download_button(
        label="📥 Atsisiųsti Excel ataskaitą su grafikais",
        data=excel_buffer,
        file_name="Klaidu_Ataskaita.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ----------------------------
    # Dirbtinio intelekto analizė
    # ----------------------------
    st.subheader("🤖 Dirbtinio intelekto analizė")

    try:
        summary_md = summary.to_markdown(index=False)

        if not priezastys.empty:
            priezastys_md = priezastys.to_markdown(index=False)
        else:
            priezastys_md = "Nėra klaidų pagal pasirinktus mėnesius."

        if not siuntejai_stats.empty:
            siuntejai_md = siuntejai_stats.to_markdown(index=False)
        else:
            siuntejai_md = "Nėra siuntėjų su klaidomis pagal pasirinktus mėnesius."

        analysis_prompt = f"""
Analizuok pateiktus duomenis apie sąskaitų skaičių, klaidų procentą, klaidų priežastis ir siuntėjus.

Prašau:
1. Įvertinti mėnesines tendencijas
2. Nustatyti, kur koncentruojasi daugiausia klaidų
3. Įvertinti, ar matomas Pareto principas
4. Išskirti didžiausią problemą procese
5. Pateikti aiškias rekomendacijas, kaip sumažinti klaidas

Mėnesių suvestinė:
{summary_md}

Klaidos pagal priežastį:
{priezastys_md}

Klaidos pagal siuntėją:
{siuntejai_md}
"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Tu esi patyręs verslo ir procesų analitikas."},
                {"role": "user", "content": analysis_prompt}
            ],
            temperature=0.4
        )

        st.markdown(response.choices[0].message.content)

    except Exception as e:
        st.warning("Nepavyko gauti AI analizės. Patikrink API raktą Streamlit `secrets` nustatymuose.")
        st.error(str(e))
