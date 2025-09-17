# app.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import io, zipfile
import streamlit as st
import pandas as pd

from builder_backend import (
    DEFAULT_HOD_SPLIT_CANDIDATES,
    DEFAULT_L2_COL,
    DEFAULT_MAPPING_PAIRS,
    DEFAULT_NUMERIC_TARGETS,
    ALL_MANAGERS_FIXED_COLUMNS,     # fixed subset used in All_Managers
    propose_mapping,
    read_headcount,
    load_template_context,
    preview_dataframe,
    build_per_hod_workbooks,
    build_master_file,
)

st.set_page_config(page_title="HOD + Master Builder", page_icon="🧩", layout="wide")
st.title("🧩 HOD Workbooks + MasterFile")

# ---------------- Session State ----------------
def init_state():
    ss = st.session_state
    ss.setdefault("step", 1)

    ss.setdefault("df", None)          # base headcount (string-typed)
    ss.setdefault("df_aug", None)      # augmented df after Step 5 (if any)

    ss.setdefault("headcount_file", None)
    ss.setdefault("template_file", None)
    ss.setdefault("master_template_file", None)

    ss.setdefault("perhod_headers_orig", [])
    ss.setdefault("master_headers_orig", [])

    # Single, unified rename editor (original header -> unified name) across both templates
    ss.setdefault("unified_rename_map", {})   # {original_header -> unified_header}

    # These are derived from unified_rename_map, kept for build steps
    ss.setdefault("perhod_rename_map", {})    # {orig -> unified} (only keys present in per-HOD template)
    ss.setdefault("master_rename_map", {})    # {orig -> unified} (only keys present in Master template)

    # Final post-rename header lists (kept for preview/mapping)
    ss.setdefault("perhod_headers", [])
    ss.setdefault("master_headers", [])

    ss.setdefault("mapping_perhod", {})
    ss.setdefault("mapping_master", {})

    ss.setdefault("hod_col", None)
    ss.setdefault("l2_col", None)
    ss.setdefault("numeric_targets", sorted(list(DEFAULT_NUMERIC_TARGETS)))

    # formatting + filenames
    ss.setdefault("normalize_headers", True)
    ss.setdefault("master_filename", "MasterFile")
    ss.setdefault("hod_name_pattern", "{HOD}")

    # mapping editor mode
    ss.setdefault("link_mappings", True)   # edit once → applies to both

    # ---- Step 5 (augment) state ----
    ss.setdefault("augment_files_meta", [])     # list of dicts describing uploaded files and mappings
    ss.setdefault("augment_overwrite", False)   # whether to overwrite non-empty headcount values

init_state()

# Convenience: set of target headers that we will show in Step 3 (those in DEFAULT_MAPPING_PAIRS)
ALLOWED_MAPPING_TARGETS = {tgt for (_src, tgt) in DEFAULT_MAPPING_PAIRS}

# Small helpers
def _norm_id_series(s):
    """Normalize an Employee ID-like series to comparable strings."""
    return s.astype(str).str.strip()

def _get_working_df():
    """Return augmented df if exists, otherwise base df."""
    return st.session_state.df_aug if st.session_state.df_aug is not None else st.session_state.df

# ---------------- Step 1: Upload ----------------
if st.session_state.step == 1:
    st.subheader("Step 1 — Upload files")

    c1, c2, c3 = st.columns(3)
    with c1:
        headcount_file = st.file_uploader("Headcount.xlsx", type=["xlsx", "xlsm"])
    with c2:
        template_file = st.file_uploader("Template (per-HOD)", type=["xlsx", "xlsm"])
    with c3:
        master_template_file = st.file_uploader("MasterTemplate (for MasterFile)", type=["xlsx", "xlsm"])

    if st.button("Next ➡️", type="primary"):
        if not headcount_file or not template_file or not master_template_file:
            st.error("Please upload all three files.")
            st.stop()

        # read headcount
        try:
            df = read_headcount(headcount_file)
        except Exception as e:
            st.error(f"Error reading Headcount: {e}")
            st.stop()

        # template contexts (you fixed sheet order already)
        try:
            perhod_ctx  = load_template_context(template_file)
            master_ctx  = load_template_context(master_template_file)
        except Exception as e:
            st.error(f"Error reading templates: {e}")
            st.stop()

        st.session_state.df = df
        st.session_state.df_aug = None   # reset any previous augmentation
        st.session_state.headcount_file = headcount_file
        st.session_state.template_file = template_file
        st.session_state.master_template_file = master_template_file

        st.session_state.perhod_headers_orig  = perhod_ctx["target_headers"]
        st.session_state.master_headers_orig  = master_ctx["target_headers"]

        # Build a single, unified rename map over the union of headers from both templates
        combined = []
        seen = set()
        # Keep order: first per-HOD headers, then add Master-only
        for h in st.session_state.perhod_headers_orig + st.session_state.master_headers_orig:
            if h not in seen:
                combined.append(h)
                seen.add(h)
        st.session_state.unified_rename_map = {h: st.session_state.unified_rename_map.get(h, h) for h in combined}

        st.session_state.step = 2
        st.rerun()

# ---------------- Step 2: Harmonize column names & formatting (single editor) ----------------
elif st.session_state.step == 2:
    st.subheader("Step 2 — Confirm Collumn Names")

    st.markdown("**Set final unified column names (applied to both templates).**")
    combined = list(st.session_state.unified_rename_map.keys())
    for h in combined:
        st.session_state.unified_rename_map[h] = st.text_input(
            f"Rename “{h}” →", value=st.session_state.unified_rename_map.get(h, h), key=f"rn_unified_{h}"
        )

    st.markdown("---")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.session_state.normalize_headers = st.checkbox(
            "Normalize header size/format",
            value=st.session_state.normalize_headers,
            help="Sets header font ~11pt, bold, centered, normal row height; applies to sub-header too if present.",
        )
    with c2:
        st.session_state.master_filename = st.text_input("Master file name (no extension)", value=st.session_state.master_filename)
    with c3:
        st.session_state.hod_name_pattern = st.text_input(
            "HOD filename pattern", value=st.session_state.hod_name_pattern, help="Use {HOD} (e.g., HOD_{HOD}_2025)"
        )

    # compute post-rename header lists for each template based on unified rename map
    perhod_headers_final = [st.session_state.unified_rename_map.get(h, h) for h in st.session_state.perhod_headers_orig]
    master_headers_final = [st.session_state.unified_rename_map.get(h, h) for h in st.session_state.master_headers_orig]

    st.session_state.perhod_headers = perhod_headers_final
    st.session_state.master_headers = master_headers_final

    # Build per-template rename maps (only include keys present in each template)
    st.session_state.perhod_rename_map = {h: st.session_state.unified_rename_map.get(h, h) for h in st.session_state.perhod_headers_orig}
    st.session_state.master_rename_map = {h: st.session_state.unified_rename_map.get(h, h) for h in st.session_state.master_headers_orig}

    # Validate equality of sets (they should match if unified correctly)
    same_names = set(perhod_headers_final) == set(master_headers_final)
    if not same_names:
        st.warning("Both data sheets (Master & HOD) should share the **same set of column names**. Please harmonize.")

    prev, nxt = st.columns([1, 1])
    if prev.button("⬅️ Back"):
        st.session_state.step = 1
        st.rerun()
    if nxt.button("Next: Map columns ➡️", type="primary"):
        if not same_names:
            st.error("Please harmonize column names first.")
        else:
            df = st.session_state.df
            # initial proposals (full sets; the UI will show only allowed targets)
            st.session_state.mapping_perhod = propose_mapping(st.session_state.perhod_headers, df.columns)
            st.session_state.mapping_master = propose_mapping(st.session_state.master_headers, df.columns)

            # defaults for split columns, numeric targets
            hod_default = next((c for c in DEFAULT_HOD_SPLIT_CANDIDATES if c in df.columns), None)
            st.session_state.hod_col = hod_default or st.session_state.hod_col or df.columns[0]
            st.session_state.l2_col  = DEFAULT_L2_COL if DEFAULT_L2_COL in df.columns else (st.session_state.l2_col or df.columns[0])

            st.session_state.step = 3
            st.rerun()

# ---------------- Step 3: Map sources (only targets from DEFAULT_MAPPING_PAIRS) ----------------
elif st.session_state.step == 3:
    st.subheader("Step 3 — connect Headcount to our files")
    st.caption("Showing only the unified columns that are linked to Headcount via `DEFAULT_MAPPING_PAIRS`.")

    df = _get_working_df()
    src_cols = df.columns.tolist()
    select_options = ["<None>"] + src_cols

    st.session_state.link_mappings = st.checkbox(
        "Edit Master & HOD mappings together",
        value=st.session_state.link_mappings,
        help="When enabled, you map columns once and it’s applied to both Master and HOD. Uncheck to edit separately.",
    )

    # display headers = only those unified headers that are targets in DEFAULT_MAPPING_PAIRS
    display_headers = [h for h in st.session_state.master_headers if h in ALLOWED_MAPPING_TARGETS]

    def mapping_editor(unified_headers, current_mapping, key_prefix):
        # Build a tiny table: Unified → Headcount source (only for allowed targets)
        map_df = pd.DataFrame({
            "Unified column": unified_headers,
            "Headcount source": [current_mapping.get(h) if current_mapping.get(h) else "<None>" for h in unified_headers],
        })
        edited = st.data_editor(
            map_df,
            key=f"de_{key_prefix}",
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "Unified column": st.column_config.Column(disabled=True, width="medium"),
                "Headcount source": st.column_config.SelectboxColumn(options=select_options, required=False, width="large"),
            },
        )
        # Convert back to dict (turn "<None>" into None) — only for unified_headers shown
        new_map = {}
        for h, src in zip(edited["Unified column"].tolist(), edited["Headcount source"].tolist()):
            new_map[h] = None if (src is None or src == "<None>") else src
        return new_map

    # Quick actions row
    c1, c2, c3, c4 = st.columns(4)
    if c1.button("✨ Auto-fill exact matches"):
        # try to match unified headers directly to headcount columns (only apply to allowed/display headers)
        for tgt in display_headers:
            if tgt in src_cols:
                st.session_state.mapping_master[tgt] = tgt
                st.session_state.mapping_perhod[tgt] = tgt
        if st.session_state.link_mappings:
            for tgt in display_headers:
                st.session_state.mapping_perhod[tgt] = st.session_state.mapping_master.get(tgt)
        st.toast("Exact matches filled where possible (only for linked targets).")

    if c2.button("🧹 Clear shown"):
        for tgt in display_headers:
            st.session_state.mapping_master[tgt] = None
            st.session_state.mapping_perhod[tgt] = None
        st.toast("Shown mappings cleared.")

    # Editors
    if st.session_state.link_mappings:
        st.markdown("#### Unified mapping")
        current = {h: st.session_state.mapping_master.get(h) for h in display_headers}
        new_map = mapping_editor(display_headers, current, key_prefix="linked")
        for k, v in new_map.items():
            st.session_state.mapping_master[k] = v
            st.session_state.mapping_perhod[k] = v
    else:
        tabs = st.tabs(["Master mapping", "HOD mapping"])
        with tabs[0]:
            current_m = {h: st.session_state.mapping_master.get(h) for h in display_headers}
            new_map_m = mapping_editor(display_headers, current_m, key_prefix="master")
            for k, v in new_map_m.items():
                st.session_state.mapping_master[k] = v
        with tabs[1]:
            current_h = {h: st.session_state.mapping_perhod.get(h) for h in display_headers}
            new_map_h = mapping_editor(display_headers, current_h, key_prefix="hod")
            for k, v in new_map_h.items():
                st.session_state.mapping_perhod[k] = v

    # Mapped/Unmapped summary for displayed keys only
    def _stats(mapping, keys):
        total = len(keys)
        mapped = sum(1 for k in keys if mapping.get(k))
        return mapped, total - mapped, total

    m_mapped, m_unmapped, m_total = _stats(st.session_state.mapping_master, display_headers)
    h_mapped, h_unmapped, h_total = _stats(st.session_state.mapping_perhod, display_headers)

    s1, s2, s3 = st.columns(3)
    with s1:
        st.metric("Master mapped (shown)", f"{m_mapped}/{m_total}", delta=f"-{m_unmapped} unmapped" if m_unmapped else "+0")
    with s2:
        st.metric("HOD mapped (shown)", f"{h_mapped}/{h_total}", delta=f"-{h_unmapped} unmapped" if h_unmapped else "+0")
    with s3:
        st.markdown("### Split & numeric")
        src_cols = _get_working_df().columns.tolist()
        st.session_state.hod_col = st.selectbox("HOD split column", options=src_cols,
                                                index=(src_cols.index(st.session_state.hod_col) if st.session_state.hod_col in src_cols else 0))
        st.session_state.l2_col  = st.selectbox("L+2 column", options=src_cols,
                                                index=(src_cols.index(st.session_state.l2_col) if st.session_state.l2_col in src_cols else 0))
        all_targets = sorted(set(st.session_state.perhod_headers) | set(st.session_state.master_headers))
        pre = [h for h in st.session_state.numeric_targets if h in all_targets]
        st.session_state.numeric_targets = st.multiselect("Treat as numeric", options=all_targets, default=pre)

    prev, nxt = st.columns([1, 1])
    if prev.button("⬅️ Back"):
        st.session_state.step = 2
        st.rerun()
    if nxt.button("Preview ➡️", type="primary"):
        st.session_state.step = 4
        st.rerun()

# ---------------- Step 4: Preview ----------------
elif st.session_state.step == 4:
    st.subheader("Step 4 — Preview & adjust")

    df = _get_working_df()
    preview_rows = st.slider("Number of preview rows", min_value=5, max_value=50, value=10, step=5)

    # Master previews
    st.markdown("### Master previews")
    master_preview_full = preview_dataframe(
        df.head(preview_rows),
        st.session_state.master_headers,
        st.session_state.mapping_master,
        st.session_state.numeric_targets
    )
    m1, m2 = st.columns(2)
    with m1:
        st.markdown("**Master data sheet (using unified names + Master mapping)**")
        st.data_editor(master_preview_full[st.session_state.master_headers], use_container_width=True, hide_index=True, disabled=True)
    with m2:
        st.markdown("**All_Managers table (fixed subset)**")
        required = ALL_MANAGERS_FIXED_COLUMNS
        missing = [c for c in required if c not in master_preview_full.columns]
        if missing:
            st.warning(f"These required All_Managers columns are missing in your Master headers: {missing}")
        present = [c for c in required if c in master_preview_full.columns]
        if present:
            st.data_editor(master_preview_full[present], use_container_width=True, hide_index=True, disabled=True)
        else:
            st.info("No required All_Managers columns are present yet. Adjust your mappings/renames.")

    st.markdown("---")
    # HOD preview
    st.markdown("### HOD preview (pick a HOD)")
    if st.session_state.hod_col not in df.columns:
        st.error(f"HOD split column '{st.session_state.hod_col}' not found.")
    else:
        values = sorted([x for x in df[st.session_state.hod_col].dropna().unique() if str(x).strip() != ""], key=lambda s: str(s).lower())
        if values:
            hod_sel = st.selectbox("Choose HOD", options=values)
            sub = df[df[st.session_state.hod_col] == hod_sel].head(preview_rows)
            perhod_preview = preview_dataframe(sub, st.session_state.perhod_headers, st.session_state.mapping_perhod, st.session_state.numeric_targets)
            st.data_editor(perhod_preview, use_container_width=True, hide_index=True, disabled=True)
        else:
            st.info("No HOD values in the chosen split column.")

    c1, c2, c3 = st.columns(3)
    if c1.button("⬅️ Back to mapping"):
        st.session_state.step = 3
        st.rerun()
    # New path to Step 5 (augment)
    if c2.button("➕ Augment from extra files"):
        st.session_state.step = 5
        st.rerun()
    if c3.button("Build files ✅", type="primary"):
        st.session_state.step = 6
        st.rerun()

# ---------------- Step 5: Augment data from extra files ----------------
elif st.session_state.step == 5:
    st.subheader("Step 5 — Augment data from extra Excel files (match by Employee ID)")
    st.caption("Upload one or more Excel files, pick their ID column, and map any file columns to your unified columns. "
               "We’ll merge by Employee ID and fill missing values (or overwrite if you choose).")

    # Upload multiple excel files
    uploaded = st.file_uploader("Upload one or more Excel files", type=["xlsx", "xlsm"], accept_multiple_files=True)

    st.session_state.augment_overwrite = st.checkbox(
        "Overwrite non-empty values in Headcount",
        value=st.session_state.augment_overwrite,
        help="If unchecked, we only fill where the Headcount value is empty."
    )

    # Build UI blocks for each file: choose ID column + mapping pairs
    files_meta = []  # rebuilt each run
    if uploaded:
        for idx, f in enumerate(uploaded, start=1):
            st.markdown(f"#### File {idx}: **{f.name}**")
            try:
                fdf = pd.read_excel(f, dtype=str)
                fdf.columns = [str(c) for c in fdf.columns]
            except Exception as e:
                st.error(f"Could not read {f.name}: {e}")
                continue

            # Guess ID column
            cols = fdf.columns.tolist()
            guess_id = None
            for cand in ["Employee ID", "Employee Id", "Emp ID", "EmpId", "ID", "Id"]:
                if cand in cols:
                    guess_id = cand
                    break
            id_col = st.selectbox(f"Employee ID column in {f.name}", options=cols, index=(cols.index(guess_id) if guess_id in cols else 0), key=f"id_{f.name}")

            st.write("**Column mappings (file ➜ unified column in template)**")
            # Prepare a dynamic editor with two columns: File column, Unified column
            # We initialize with an empty row to let users start selecting.
            map_df_init = pd.DataFrame({"File column": [None], "Unified column": [None]})

            edited = st.data_editor(
                map_df_init,
                key=f"map_{f.name}",
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "File column": st.column_config.SelectboxColumn(options=[c for c in cols if c != id_col], required=False, width="large"),
                    "Unified column": st.column_config.SelectboxColumn(
                        options=sorted(set(st.session_state.perhod_headers) | set(st.session_state.master_headers)),
                        required=False,
                        width="large"
                    ),
                },
            )

            # Clean up mapping rows
            pairs = []
            for _, row in edited.iterrows():
                file_col = row.get("File column")
                uni_col = row.get("Unified column")
                if file_col and uni_col:
                    pairs.append((file_col, uni_col))

            files_meta.append({
                "name": f.name,
                "df": fdf.copy(),
                "id_col": id_col,
                "pairs": pairs,
            })

    # Action buttons
    c1, c2, c3 = st.columns(3)
    if c1.button("⬅️ Back to Preview"):
        st.session_state.step = 4
        st.rerun()

    def _apply_augment(base_df, files_meta, overwrite=False):
        """Create df_aug by merging in columns from files_meta by Employee ID."""
        if base_df is None or base_df.empty or not files_meta:
            return base_df

        # Ensure Headcount Employee ID exists
        if "Employee ID" not in base_df.columns:
            st.error("Headcount must contain 'Employee ID' column to augment.")
            return base_df

        out = base_df.copy()
        out["Employee ID"] = _norm_id_series(out["Employee ID"])

        for meta in files_meta:
            fdf = meta["df"].copy()
            id_col = meta["id_col"]
            pairs = meta["pairs"]

            if not pairs:
                continue

            if id_col not in fdf.columns:
                st.warning(f"Skipping {meta['name']}: selected ID column '{id_col}' not found.")
                continue

            # Normalize the ID col in the file
            fdf[id_col] = _norm_id_series(fdf[id_col])

            # Keep only needed columns
            keep_cols = [id_col] + [p[0] for p in pairs]
            fdf_narrow = fdf[keep_cols].drop_duplicates(subset=[id_col])

            # Build a temporary merge for each pair (to avoid name collisions)
            # but we can also merge once with all file columns and then assign
            merged = out.merge(fdf_narrow, how="left", left_on="Employee ID", right_on=id_col, suffixes=("", "_extra"))

            # For each pair: file_col -> unified_target
            for file_col, tgt in pairs:
                src_series = merged[file_col]
                if tgt not in merged.columns:
                    # ensure the target column exists in working df, we will create it
                    merged[tgt] = ""

                if overwrite:
                    merged[tgt] = src_series.where(src_series.notna(), merged[tgt])
                else:
                    # fill only where the current value is empty/na
                    is_empty = merged[tgt].isna() | (merged[tgt].astype(str).str.strip() == "")
                    merged.loc[is_empty, tgt] = src_series[is_empty]

                # Also make sure we point mapping to this new source if mapping is empty:
                # strategy: since df_to_rows uses mapping[target] = source-column-name,
                # create a true source column with the SAME name as the template target
                # then set mapping_*[tgt] = tgt (self-source)
                st.session_state.mapping_master[tgt] = st.session_state.mapping_master.get(tgt) or tgt
                st.session_state.mapping_perhod[tgt] = st.session_state.mapping_perhod.get(tgt) or tgt

            # Drop the right_on key col if it’s different from Employee ID
            if id_col != "Employee ID":
                merged = merged.drop(columns=[id_col])

            out = merged

        return out

    if c3.button("Apply augmentation ✅", type="primary"):
        base = _get_working_df()  # could already be augmented; apply on top to allow iterative adds
        st.session_state.df_aug = _apply_augment(base, files_meta, overwrite=st.session_state.augment_overwrite)
        # Update split column choices if new columns arrived
        if st.session_state.hod_col not in _get_working_df().columns:
            st.session_state.hod_col = "Employee ID" if "Employee ID" in _get_working_df().columns else _get_working_df().columns[0]
        if st.session_state.l2_col not in _get_working_df().columns:
            st.session_state.l2_col = "Employee ID" if "Employee ID" in _get_working_df().columns else _get_working_df().columns[0]
        st.success("Augmentation applied. Your previews & build will use the augmented data.")
        st.session_state.step = 4
        st.rerun()

# ---------------- Step 6: Build & Download ----------------
elif st.session_state.step == 6:
    st.subheader("Step 6 — Build & download")

    df = _get_working_df()

    try:
        with st.spinner("Building per-HOD workbooks…"):
            hod_files = build_per_hod_workbooks(
                df=df,
                template_file=st.session_state.template_file,
                mapping_perhod=st.session_state.mapping_perhod,
                hod_col=st.session_state.hod_col,
                l2_col=st.session_state.l2_col,
                numeric_targets=st.session_state.numeric_targets,
                perhod_rename_map=st.session_state.perhod_rename_map,   # enforce unified names
                normalize_headers=st.session_state.normalize_headers,
                output_name_pattern=st.session_state.hod_name_pattern,
            )

        with st.spinner("Building MasterFile…"):
            master_name, master_bytes = build_master_file(
                df=df,
                master_template_file=st.session_state.master_template_file,
                mapping_master=st.session_state.mapping_master,
                numeric_targets=st.session_state.numeric_targets,
                all_mgrs_columns=None,               # backend uses fixed subset
                all_mgrs_renames=None,
                master_rename_map=st.session_state.master_rename_map,   # enforce unified names
                normalize_headers=st.session_state.normalize_headers,
                master_filename=st.session_state.master_filename,
            )

        # Zip everything
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for fname, fbytes in hod_files:
                zf.writestr(f"HODs/{fname}", fbytes)
            zf.writestr(master_name, master_bytes)
        zip_buf.seek(0)

        st.success("Done! Download your files below.")
        st.download_button("📦 Download All (ZIP)", data=zip_buf, file_name="Headcount_Output.zip", mime="application/zip")

        st.markdown("**Individual downloads**")
        st.download_button(
            f"⬇️ {master_name}",
            data=master_bytes,
            file_name=master_name,
            mime="application/vnd.ms-excel.sheet.macroEnabled.12" if master_name.endswith(".xlsm") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        if hod_files:
            for fname, fbytes in hod_files:
                st.download_button(
                    f"⬇️ {fname}",
                    data=fbytes,
                    file_name=fname,
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12" if fname.endswith(".xlsm") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=fname,
                )
        else:
            st.info("No HOD workbooks were produced (no HODs detected).")

    except Exception as e:
        st.error(f"Build error: {e}")

    st.markdown("---")
    if st.button("⬅️ Start over"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()
