# klaidu-analize-streamlit

Šis projektas yra interaktyvus Streamlit skydelis, skirtas analizuoti sąskaitų išrašymo klaidas, nustatyti jų priežastis ir atskleisti procesų neefektyvumą. Įrankis buvo sukurtas tam, kad rankiniu būdu tvarkomas ir chaotiškas klaidų sekimo procesas būtų paverstas struktūrizuota analitine sistema, leidžiančia suprasti, kur atsiranda klaidos, kodėl jos atsiranda ir kaip jas galima sumažinti.

Daugelyje organizacijų sąskaitų klaidos yra tvarkomos reaktyviai. Klaidos taisomos rankiniu būdu, vėliau tenka ieškoti susirašinėjimo istorijos, klaidų priežastys pamirštamos, o tos pačios problemos kartojasi. Šis dashboardas sprendžia šią problemą, paversdamas klaidų registrą aiškiomis analitinėmis įžvalgomis ir vizualizacijomis.

Programėlė leidžia įkelti Excel failą su sąskaitų duomenimis ir automatiškai sugeneruoja pilną proceso analizę. Ji padeda greitai suprasti sąskaitų apimtis, klaidų dažnį, klaidų priežastis bei pagrindinius klaidų šaltinius.

Skydelio viršuje pateikiami pagrindiniai rodikliai (KPI), kurie leidžia iš karto įvertinti proceso būklę. Čia matomas bendras apdorotų sąskaitų skaičius, sąskaitų su klaidomis skaičius, klaidų procentas ir teisingai apdorotų sąskaitų kiekis. Tai leidžia greitai įvertinti bendrą proceso kokybę.

Sistema taip pat atlieka mėnesinę analizę, automatiškai ištraukdama mėnesius iš sąskaitų įrašų. Rodoma, kiek sąskaitų buvo apdorota kiekvieną mėnesį, kiek jų turėjo klaidų ir koks buvo klaidų procentas. Normalizuotas grafikas leidžia palyginti darbo krūvį su klaidų lygiu ir padeda pastebėti tendencijas arba neįprastus klaidų šuolius.

Svarbi analizės dalis yra klaidų priežasčių vertinimas. Klaidos grupuojamos pagal jų registruotą priežastį, todėl galima aiškiai matyti, kokie klaidų tipai pasitaiko dažniausiai. Tai padeda identifikuoti sistemines proceso problemas ir nuspręsti, kurias jų reikėtų spręsti pirmiausia.

Sąskaitos taip pat analizuojamos pagal siuntėją. Skydelis apskaičiuoja, kiek dokumentų pateikė kiekvienas siuntėjas, kiek klaidų buvo susiję su jo dokumentais ir koks yra klaidų procentas. Tai svarbu todėl, kad leidžia atskirti du skirtingus aspektus: klaidų kiekį ir klaidų kokybę. Kai kurie siuntėjai gali turėti daugiau klaidų vien dėl to, kad siunčia daugiau dokumentų. Klaidų procento analizė padeda nustatyti, kur iš tikrųjų yra kokybės problema.

Klaidos taip pat analizuojamos pagal užsakovą. Tai leidžia nustatyti, ar tam tikri klientai ar projektai dažniau sukelia problemas. Tokia analizė gali parodyti neaiškius dokumentacijos reikalavimus arba pasikartojančias problemas bendradarbiaujant su tam tikrais partneriais.

Skydelyje taip pat naudojama Pareto analizė, padedanti nustatyti, kur atsiranda didžioji dalis problemų. Pareto grafikas parodo, kurie siuntėjai sukuria didžiausią klaidų dalį ir kaip klaidos pasiskirsto tarp skirtingų šaltinių. Tai leidžia prioritetizuoti procesų gerinimo veiksmus ir pirmiausia spręsti tas problemas, kurios turi didžiausią poveikį.

Sistema taip pat pateikia detalų klaidų sąrašą. Jame matomas mėnuo, užsakovas, sąskaitos numeris, klaidos aprašymas, klaidos priežastis ir siuntėjas. Tai leidžia greitai išanalizuoti konkrečius atvejus ir suprasti, iš kur atsirado problema.

Skydelis gali sugeneruoti pilną Excel ataskaitą su visa analize. Eksportuojamoje ataskaitoje yra mėnesinė suvestinė, klaidų sąrašas, klaidų priežasčių statistika, siuntėjų analizė, siuntėjų kokybės analizė, užsakovų analizė ir Pareto analizė. Visi pagrindiniai grafikai automatiškai įterpiami į Excel dokumentą.

Be vizualinės analizės, sistema gali sugeneruoti ir dirbtinio intelekto pateikiamą duomenų interpretaciją. Dirbtinis intelektas įvertina mėnesines tendencijas, nustato pagrindines klaidų priežastis, analizuoja klaidų koncentraciją, vertina siuntėjų kokybę ir pateikia galimas procesų tobulinimo rekomendacijas.

Analitinė šio dashboardo logika sąmoningai atskiria klaidų kiekį nuo klaidų kokybės. Klaidų kiekis parodo, kur atsiranda daugiausia klaidų, o klaidų procentas parodo, kur yra didžiausia klaidų tikimybė. Toks požiūris padeda išvengti klaidingų išvadų, kai kai kurie siuntėjai tiesiog apdoroja daugiau dokumentų nei kiti.

Norint naudoti šią programą, Excel faile turi būti bent šie stulpeliai: Klientas, Užsakovas, Sąskaitos faktūros Nr. ir Siuntėjas. Papildomai naudojami du stulpeliai: O stulpelis naudojamas kaip klaidos priežastis, o P stulpelis naudojamas kaip klaidos aprašymas. Sistema šiuos laukus automatiškai identifikuoja ir panaudoja analizėje.

Programėlė sukurta naudojant Python ir šias technologijas: Streamlit interaktyviai sąsajai, Pandas duomenų apdorojimui, Matplotlib vizualizacijoms, OpenPyXL Excel ataskaitų generavimui ir OpenAI API dirbtinio intelekto analitinėms įžvalgoms.

Norint paleisti programą lokaliai, reikia įdiegti reikalingas bibliotekas ir paleisti Streamlit serverį:

pip install streamlit pandas matplotlib openpyxl openai
streamlit run app.py

Šis skydelis ypač naudingas apskaitos skyriams, sąskaitų administravimo komandoms, finansų operacijų specialistams, procesų tobulinimo projektams ir vidaus auditui. Jo tikslas nėra tik suskaičiuoti klaidas, bet padėti suprasti, kur reikėtų įsikišti pirmiausia, kad klaidų skaičius sumažėtų ir procesai taptų efektyvesni.
