import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Excel Vergleichstool", layout="wide")
st.title("🔍 Vertrags-/Asset-Datenvergleich (Test vs. Prod)")

# ===== Mode selection (valid for both languages) =====
mode = st.radio(
    "Bitte Bereich wählen / Choose section:",
    ["1️⃣ Compare closings", "2️⃣ Compare contract list"],
    horizontal=True,
)

# Tabs für Sprache
tab_de, tab_en = st.tabs(["🇩🇪 Deutsch", "🇬🇧 English"])

# 💡 Gemeinsame Bereinigungsfunktion für Closings
@st.cache_data
def clean_and_prepare(uploaded_file, id_col, asset_col):
    df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    header_1 = df_raw.iloc[1]
    header_2 = df_raw.iloc[2]
    header_3 = df_raw.iloc[3]

    header_1 = header_1.fillna(method="ffill")
    header_2 = header_2.fillna(method="ffill")

    df_data = df_raw.iloc[4:].copy()
    df_data.reset_index(drop=True, inplace=True)

    columns_combined = []
    for i in range(len(header_1)):
        if i < 9:
            columns_combined.append(header_1[i])
        else:
            beschreibung = re.sub(r'\s+', ' ', str(header_1[i]).strip())
            konto_nr = str(header_2[i]).strip()
            soll_haben = str(header_3[i]).strip()
            name = f"{beschreibung} - {konto_nr}_IFRS16 - {soll_haben}"
            columns_combined.append(name)

    df_data.columns = columns_combined

    df_data[id_col] = df_data[id_col].astype(str)
    df_data[asset_col] = df_data[asset_col].astype(str)
    df_data["Key"] = df_data[id_col] + "_" + df_data[asset_col]

    return df_data.set_index("Key")

# 💡 Vorbereitung für Vertragsliste (normale Header-Struktur)
@st.cache_data
def prepare_contract_list(uploaded_file, system_id_col, asset_col):
    df = pd.read_excel(uploaded_file)  # Header in Zeile 1
    df[system_id_col] = df[system_id_col].astype(str)
    df[asset_col] = df[asset_col].astype(str)
    df["Key"] = df[system_id_col] + "_" + df[asset_col]
    return df.set_index("Key")

# ===== Nur Logik: Numerische Abweichungen < 1 ignorieren =====
TOL = 1.0  # fester Schwellwert; Frontend bleibt unverändert

def _try_parse_number(val):
    """Versucht, val als Zahl zu interpretieren (DE/EN-Formate, Währungs-/%-Zeichen)."""
    if pd.isna(val):
        return False, None
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return True, float(val)
    s = str(val).strip()
    if s == "":
        return False, None

    s_clean = (
        s.replace("\xa0", "")
         .replace("€", "")
         .replace("%", "")
         .replace(" ", "")
         .replace("’", "")
         .replace("'", "")
    )
    # DE: 1.234,56
    try:
        s_de = s_clean.replace(".", "").replace(",", ".")
        return True, float(s_de)
    except Exception:
        pass
    # EN: 1,234.56
    try:
        s_en = s_clean.replace(",", "")
        return True, float(s_en)
    except Exception:
        pass
    return False, None

def nearly_equal(a, b, tol=TOL) -> bool:
    """True, wenn a und b numerisch sind und |a-b| < tol."""
    ok_a, fa = _try_parse_number(a)
    ok_b, fb = _try_parse_number(b)
    if ok_a and ok_b:
        return abs(fa - fb) < tol
    return False
# =============================================================

# 🇩🇪 Deutsch
with tab_de:
    if mode.startswith("1️⃣"):
        # ====== SECTION 1: Closings vergleichen (DE) ======
        st.subheader("📂 Dateien für Closings hochladen")
        file_test = st.file_uploader("Test-Datei (Closings) hochladen", type=["xlsx"], key="test_de_closings")
        file_prod = st.file_uploader("Prod-Datei (Closings) hochladen", type=["xlsx"], key="prod_de_closings")

        id_col = "Vertrags-ID"
        asset_col = "Asset-ID"

        if file_test and file_prod:
            df_test = clean_and_prepare(file_test, id_col, asset_col)
            df_prod = clean_and_prepare(file_prod, id_col, asset_col)

            columns_test = set(df_test.columns) - {id_col, asset_col, "Key"}
            columns_prod = set(df_prod.columns) - {id_col, asset_col, "Key"}

            only_in_test = sorted(columns_test - columns_prod)
            only_in_prod = sorted(columns_prod - columns_test)

            if only_in_test:
                st.warning("⚠️ Spalten nur in Test:")
                st.code("\n".join(only_in_test))
            if only_in_prod:
                st.warning("⚠️ Spalten nur in Prod:")
                st.code("\n".join(only_in_prod))
            if not only_in_test and not only_in_prod:
                st.success("✅ Alle Spalten stimmen überein.")

            col1, col2 = st.columns(2)
            with col1:
                out_test = io.BytesIO()
                with pd.ExcelWriter(out_test, engine="xlsxwriter") as writer:
                    df_test.reset_index().to_excel(writer, index=False, sheet_name="Bereinigt_Test")
                st.download_button(
                    "⬇️ Bereinigte Test-Datei",
                    data=out_test.getvalue(),
                    file_name="bereinigt_test.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with col2:
                out_prod = io.BytesIO()
                with pd.ExcelWriter(out_prod, engine="xlsxwriter") as writer:
                    df_prod.reset_index().to_excel(writer, index=False, sheet_name="Bereinigt_Prod")
                st.download_button(
                    "⬇️ Bereinigte Prod-Datei",
                    data=out_prod.getvalue(),
                    file_name="bereinigt_prod.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            all_keys = sorted(set(df_test.index).union(set(df_prod.index)))
            common_cols = df_test.columns.intersection(df_prod.columns).difference([id_col, asset_col])

            results = []
            for key in all_keys:
                row = {
                    id_col: key.split("_")[0],
                    asset_col: "_".join(key.split("_")[1:])
                }

                if key not in df_test.index:
                    row["Unterschiede"] = "Nur in Prod"
                elif key not in df_prod.index:
                    row["Unterschiede"] = "Nur in Test"
                else:
                    diffs = []
                    for col in common_cols:
                        val_test = df_test.loc[key, col]
                        val_prod = df_prod.loc[key, col]
                        if isinstance(val_test, pd.Series):
                            val_test = val_test.iloc[0]
                        if isinstance(val_prod, pd.Series):
                            val_prod = val_prod.iloc[0]

                        if pd.isna(val_test) and pd.isna(val_prod):
                            continue
                        # numerische Abweichungen < 1 ignorieren
                        if nearly_equal(val_test, val_prod, TOL):
                            continue

                        if pd.isna(val_test) or pd.isna(val_prod) or val_test != val_prod:
                            diffs.append(f"{col}: Test={val_test} / Prod={val_prod}")
                    row["Unterschiede"] = "; ".join(diffs) if diffs else "Keine"
                results.append(row)

            df_diff = pd.DataFrame(results)
            df_diff = df_diff[df_diff["Unterschiede"] != "Keine"]

            st.success(f"✅ Vergleich abgeschlossen. {len(df_diff)} Zeilen analysiert.")
            st.dataframe(df_diff, use_container_width=True)

            out_result = io.BytesIO()
            with pd.ExcelWriter(out_result, engine="xlsxwriter") as writer:
                df_diff.to_excel(writer, index=False, sheet_name="Vergleich")
            st.download_button(
                "📥 Vergleichsergebnis herunterladen",
                data=out_result.getvalue(),
                file_name="vergleichsergebnis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    else:
        # ====== SECTION 2: Vertragsliste vergleichen (DE) ======
        st.subheader("📂 Vertragslisten hochladen")
        file_test = st.file_uploader("Test-Vertragsliste hochladen", type=["xlsx"], key="test_de_contracts")
        file_prod = st.file_uploader("Prod-Vertragsliste hochladen", type=["xlsx"], key="prod_de_contracts")

        system_id_col = "Vertrags-ID"
        asset_col = "Asset-ID"

        if file_test and file_prod:
            df_test = prepare_contract_list(file_test, system_id_col, asset_col)
            df_prod = prepare_contract_list(file_prod, system_id_col, asset_col)

            columns_test = set(df_test.columns) - {system_id_col, asset_col, "Key"}
            columns_prod = set(df_prod.columns) - {system_id_col, asset_col, "Key"}

            only_in_test = sorted(columns_test - columns_prod)
            only_in_prod = sorted(columns_prod - columns_test)

            if only_in_test:
                st.warning("⚠️ Spalten nur in Test:")
                st.code("\n".join(only_in_test))
            if only_in_prod:
                st.warning("⚠️ Spalten nur in Prod:")
                st.code("\n".join(only_in_prod))
            if not only_in_test and not only_in_prod:
                st.success("✅ Alle Spalten stimmen überein.")

            col1, col2 = st.columns(2)
            with col1:
                out_test = io.BytesIO()
                with pd.ExcelWriter(out_test, engine="xlsxwriter") as writer:
                    df_test.reset_index().to_excel(writer, index=False, sheet_name="Vertragsliste_Test")
                st.download_button(
                    "⬇️ Bereinigte Test-Vertragsliste",
                    data=out_test.getvalue(),
                    file_name="vertragsliste_test.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with col2:
                out_prod = io.BytesIO()
                with pd.ExcelWriter(out_prod, engine="xlsxwriter") as writer:
                    df_prod.reset_index().to_excel(writer, index=False, sheet_name="Vertragsliste_Prod")
                st.download_button(
                    "⬇️ Bereinigte Prod-Vertragsliste",
                    data=out_prod.getvalue(),
                    file_name="vertragsliste_prod.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            all_keys = sorted(set(df_test.index).union(set(df_prod.index)))
            common_cols = df_test.columns.intersection(df_prod.columns).difference([system_id_col, asset_col])

            results = []
            for key in all_keys:
                row = {
                    system_id_col: key.split("_")[0],
                    asset_col: "_".join(key.split("_")[1:])
                }

                if key not in df_test.index:
                    row["Unterschiede"] = "Nur in Prod"
                elif key not in df_prod.index:
                    row["Unterschiede"] = "Nur in Test"
                else:
                    diffs = []
                    for col in common_cols:
                        val_test = df_test.loc[key, col]
                        val_prod = df_prod.loc[key, col]
                        if isinstance(val_test, pd.Series):
                            val_test = val_test.iloc[0]
                        if isinstance(val_prod, pd.Series):
                            val_prod = val_prod.iloc[0]

                        if pd.isna(val_test) and pd.isna(val_prod):
                            continue
                        if nearly_equal(val_test, val_prod, TOL):
                            continue

                        if pd.isna(val_test) or pd.isna(val_prod) or val_test != val_prod:
                            diffs.append(f"{col}: Test={val_test} / Prod={val_prod}")
                    row["Unterschiede"] = "; ".join(diffs) if diffs else "Keine"
                results.append(row)

            df_diff = pd.DataFrame(results)
            df_diff = df_diff[df_diff["Unterschiede"] != "Keine"]

            st.success(f"✅ Vergleich abgeschlossen. {len(df_diff)} Zeilen analysiert.")
            st.dataframe(df_diff, use_container_width=True)

            out_result = io.BytesIO()
            with pd.ExcelWriter(out_result, engine="xlsxwriter") as writer:
                df_diff.to_excel(writer, index=False, sheet_name="Vertragslisten-Vergleich")
            st.download_button(
                "📥 Vergleichsergebnis herunterladen",
                data=out_result.getvalue(),
                file_name="vertragslisten_vergleich.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# 🇬🇧 English
with tab_en:
    if mode.startswith("1️⃣"):
        # ====== SECTION 1: Compare closings (EN) ======
        st.subheader("📂 Upload files for closings")
        file_test = st.file_uploader("Upload Test closing file", type=["xlsx"], key="test_en_closings")
        file_prod = st.file_uploader("Upload Prod closing file", type=["xlsx"], key="prod_en_closings")

        id_col = "Contract ID"
        asset_col = "Asset ID"

        if file_test and file_prod:
            df_test = clean_and_prepare(file_test, id_col, asset_col)
            df_prod = clean_and_prepare(file_prod, id_col, asset_col)

            columns_test = set(df_test.columns) - {id_col, asset_col, "Key"}
            columns_prod = set(df_prod.columns) - {id_col, asset_col, "Key"}

            only_in_test = sorted(columns_test - columns_prod)
            only_in_prod = sorted(columns_prod - columns_test)

            if only_in_test:
                st.warning("⚠️ Columns only in Test:")
                st.code("\n".join(only_in_test))
            if only_in_prod:
                st.warning("⚠️ Columns only in Prod:")
                st.code("\n".join(only_in_prod))
            if not only_in_test and not only_in_prod:
                st.success("✅ All columns match.")

            col1, col2 = st.columns(2)
            with col1:
                out_test = io.BytesIO()
                with pd.ExcelWriter(out_test, engine="xlsxwriter") as writer:
                    df_test.reset_index().to_excel(writer, index=False, sheet_name="Cleaned_Test")
                st.download_button(
                    "⬇️ Download cleaned Test file",
                    data=out_test.getvalue(),
                    file_name="cleaned_test.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with col2:
                out_prod = io.BytesIO()
                with pd.ExcelWriter(out_prod, engine="xlsxwriter") as writer:
                    df_prod.reset_index().to_excel(writer, index=False, sheet_name="Cleaned_Prod")
                st.download_button(
                    "⬇️ Download cleaned Prod file",
                    data=out_prod.getvalue(),
                    file_name="cleaned_prod.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            all_keys = sorted(set(df_test.index).union(set(df_prod.index)))
            common_cols = df_test.columns.intersection(df_prod.columns).difference([id_col, asset_col])

            results = []
            for key in all_keys:
                row = {
                    id_col: key.split("_")[0],
                    asset_col: "_".join(key.split("_")[1:])
                }

                if key not in df_test.index:
                    row["Differences"] = "Only in Prod"
                elif key not in df_prod.index:
                    row["Differences"] = "Only in Test"
                else:
                    diffs = []
                    for col in common_cols:
                        val_test = df_test.loc[key, col]
                        val_prod = df_prod.loc[key, col]
                        if isinstance(val_test, pd.Series):
                            val_test = val_test.iloc[0]
                        if isinstance(val_prod, pd.Series):
                            val_prod = val_prod.iloc[0]

                        if pd.isna(val_test) and pd.isna(val_prod):
                            continue
                        # Logic only: ignore numeric deltas < 1
                        if nearly_equal(val_test, val_prod, TOL):
                            continue

                        if pd.isna(val_test) or pd.isna(val_prod) or val_test != val_prod:
                            diffs.append(f"{col}: Test={val_test} / Prod={val_prod}")
                    row["Differences"] = "; ".join(diffs) if diffs else "None"
                results.append(row)

            df_diff = pd.DataFrame(results)
            df_diff = df_diff[df_diff["Differences"] != "None"]
            st.success(f"✅ Comparison complete. {len(df_diff)} rows analyzed.")
            st.dataframe(df_diff, use_container_width=True)

            out_result = io.BytesIO()
            with pd.ExcelWriter(out_result, engine="xlsxwriter") as writer:
                df_diff.to_excel(writer, index=False, sheet_name="Comparison")
            st.download_button(
                "📥 Download comparison result",
                data=out_result.getvalue(),
                file_name="comparison_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    else:
        # ====== SECTION 2: Compare contract list (EN) ======
        st.subheader("📂 Upload contract lists")
        file_test = st.file_uploader("Upload Test contract list", type=["xlsx"], key="test_en_contracts")
        file_prod = st.file_uploader("Upload Prod contract list", type=["xlsx"], key="prod_en_contracts")

        system_id_col = "Contract ID"
        asset_col = "Asset ID"

        if file_test and file_prod:
            df_test = prepare_contract_list(file_test, system_id_col, asset_col)
            df_prod = prepare_contract_list(file_prod, system_id_col, asset_col)

            columns_test = set(df_test.columns) - {system_id_col, asset_col, "Key"}
            columns_prod = set(df_prod.columns) - {system_id_col, asset_col, "Key"}

            only_in_test = sorted(columns_test - columns_prod)
            only_in_prod = sorted(columns_prod - columns_test)

            if only_in_test:
                st.warning("⚠️ Columns only in Test:")
                st.code("\n".join(only_in_test))
            if only_in_prod:
                st.warning("⚠️ Columns only in Prod:")
                st.code("\n".join(only_in_prod))
            if not only_in_test and not only_in_prod:
                st.success("✅ All columns match.")

            col1, col2 = st.columns(2)
            with col1:
                out_test = io.BytesIO()
                with pd.ExcelWriter(out_test, engine="xlsxwriter") as writer:
                    df_test.reset_index().to_excel(writer, index=False, sheet_name="ContractList_Test")
                st.download_button(
                    "⬇️ Download cleaned Test contract list",
                    data=out_test.getvalue(),
                    file_name="contract_list_test.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with col2:
                out_prod = io.BytesIO()
                with pd.ExcelWriter(out_prod, engine="xlsxwriter") as writer:
                    df_prod.reset_index().to_excel(writer, index=False, sheet_name="ContractList_Prod")
                st.download_button(
                    "⬇️ Download cleaned Prod contract list",
                    data=out_prod.getvalue(),
                    file_name="contract_list_prod.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            all_keys = sorted(set(df_test.index).union(set(df_prod.index)))
            common_cols = df_test.columns.intersection(df_prod.columns).difference([system_id_col, asset_col])

            results = []
            for key in all_keys:
                row = {
                    system_id_col: key.split("_")[0],
                    asset_col: "_".join(key.split("_")[1:])
                }

                if key not in df_test.index:
                    row["Differences"] = "Only in Prod"
                elif key not in df_prod.index:
                    row["Differences"] = "Only in Test"
                else:
                    diffs = []
                    for col in common_cols:
                        val_test = df_test.loc[key, col]
                        val_prod = df_prod.loc[key, col]
                        if isinstance(val_test, pd.Series):
                            val_test = val_test.iloc[0]
                        if isinstance(val_prod, pd.Series):
                            val_prod = val_prod.iloc[0]

                        if pd.isna(val_test) and pd.isna(val_prod):
                            continue
                        if nearly_equal(val_test, val_prod, TOL):
                            continue

                        if pd.isna(val_test) or pd.isna(val_prod) or val_test != val_prod:
                            diffs.append(f"{col}: Test={val_test} / Prod={val_prod}")
                    row["Differences"] = "; ".join(diffs) if diffs else "None"
                results.append(row)

            df_diff = pd.DataFrame(results)
            df_diff = df_diff[df_diff["Differences"] != "None"]
            st.success(f"✅ Comparison complete. {len(df_diff)} rows analyzed.")
            st.dataframe(df_diff, use_container_width=True)

            out_result = io.BytesIO()
            with pd.ExcelWriter(out_result, engine="xlsxwriter") as writer:
                df_diff.to_excel(writer, index=False, sheet_name="ContractList_Comparison")
            st.download_button(
                "📥 Download contract list comparison result",
                data=out_result.getvalue(),
                file_name="contract_list_comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
