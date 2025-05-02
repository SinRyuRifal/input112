import streamlit as st
import pandas as pd
import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import io


def setup_driver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless=new")  # Gunakan mode headless Chrome
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    return webdriver.Chrome(options=chrome_options)


def process_excel(file):
    log = []
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"NIK_Tidak_Ditemukan_{timestamp}.xlsx"
    log_filename = f"log_input_{timestamp}.txt"
    not_found_niks = []

    try:
        df = pd.read_excel(file)
        df["Tanggal Event"] = pd.to_datetime(
            df["Tanggal Event"], errors="coerce"
        ).dt.strftime("%d-%m-%Y")

        driver = setup_driver()
        driver.get("http://112.140.160.228/pum/login")
        time.sleep(2)

        driver.find_element(By.NAME, "USERNAME").send_keys("magangmsib")
        driver.find_element(By.NAME, "PASSWORD").send_keys("msib")
        driver.find_element(By.CLASS_NAME, "login100-form-btn").click()
        time.sleep(3)

        driver.find_element(By.LINK_TEXT, "Data UMKM").click()
        time.sleep(3)

        for index, row in df.iterrows():
            nik = str(row["NIK"])
            try:
                search_box = driver.find_element(
                    By.CSS_SELECTOR, "input[type='search']"
                )
                search_box.clear()
                search_box.send_keys(nik)
                search_box.send_keys(Keys.RETURN)
                time.sleep(3)

                data_rows = driver.find_elements(
                    By.XPATH, f"//td[contains(text(), '{nik}')]"
                )
                if data_rows:
                    log.append(f"Data ditemukan: {nik}")
                    driver.find_element(
                        By.XPATH,
                        "//a[contains(@class, 'btn-warning') and contains(text(), 'Edit')]",
                    ).click()
                    time.sleep(3)
                    driver.find_element(By.LINK_TEXT, "Intervensi").click()
                    time.sleep(3)

                    driver.find_element(By.NAME, "BULAN_TAHUN").send_keys(
                        row["Tanggal Event"]
                    )
                    driver.find_element(By.NAME, "intervensi").send_keys(
                        row["Jenis Intervensi"]
                    )
                    driver.find_element(By.NAME, "nama_kegiatan").send_keys(
                        row["Nama Event"]
                    )
                    driver.find_element(By.NAME, "OMSET").send_keys(str(row["Omset"]))
                    driver.find_element(
                        By.XPATH, "//button[contains(text(), 'Tambah Data Intervensi')]"
                    ).click()
                    time.sleep(2)

                    log.append(f"Intervensi ditambahkan untuk: {nik}")
                    driver.find_element(By.LINK_TEXT, "Data UMKM").click()
                    time.sleep(2)
                else:
                    log.append(f"Tidak ditemukan: {nik}")
                    not_found_niks.append(nik)
            except Exception as e:
                log.append(f"Error saat proses NIK {nik}: {str(e)}")
                not_found_niks.append(nik)
                continue

    except Exception as e:
        log.append(f"ERROR GLOBAL: {str(e)}")

    # Simpan log ke file teks
    os.makedirs("logs", exist_ok=True)
    with open(os.path.join("logs", log_filename), "w") as log_file:
        log_file.write("\n".join(log))

    # Membuat file Excel hasil NIK tidak ditemukan langsung di memori
    if not_found_niks:
        result_df = pd.DataFrame({"NIK Tidak Ditemukan": not_found_niks})
        output = io.BytesIO()
        result_df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)  # Kembali ke awal stream untuk di-download
        return log, output, output_filename
    else:
        return log, None, None


def cek_nik_excel(file):
    log = []
    not_found = []
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"NIK_Tidak_Ditemukan_Cek_{timestamp}.xlsx"

    try:
        df = pd.read_excel(file)
        driver = setup_driver()
        driver.get("http://112.140.160.228/pum/login")
        time.sleep(2)

        driver.find_element(By.NAME, "USERNAME").send_keys("magangmsib")
        driver.find_element(By.NAME, "PASSWORD").send_keys("msib")
        driver.find_element(By.CLASS_NAME, "login100-form-btn").click()
        time.sleep(3)

        driver.find_element(By.LINK_TEXT, "Data UMKM").click()
        time.sleep(3)

        for _, row in df.iterrows():
            nik = str(row["NIK"])
            try:
                search_box = driver.find_element(
                    By.CSS_SELECTOR, "input[type='search']"
                )
                search_box.clear()
                search_box.send_keys(nik)
                search_box.send_keys(Keys.RETURN)
                time.sleep(3)

                results = driver.find_elements(
                    By.XPATH, f"//td[contains(text(), '{nik}')]"
                )
                if results:
                    log.append(f"NIK ditemukan: {nik}")
                else:
                    log.append(f"NIK tidak ditemukan: {nik}")
                    not_found.append(nik)
            except Exception as e:
                log.append(f"Error saat cek NIK {nik}: {str(e)}")
                not_found.append(nik)

    except Exception as e:
        log.append(f"ERROR GLOBAL: {str(e)}")

    # Membuat file Excel hasil NIK tidak ditemukan langsung di memori
    if not_found:
        result_df = pd.DataFrame({"NIK Tidak Ditemukan": not_found})
        output = io.BytesIO()
        result_df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)  # Kembali ke awal stream untuk di-download
        return log, output, output_filename
    else:
        return log, None, None


# ---------------------- STREAMLIT APP -----------------------
st.set_page_config(page_title="Input Otomatis Intervensi", layout="centered")
st.title("Input Otomatis Intervensi UMKM")
st.markdown("Credit : Ripal")
st.markdown("Follow Instagram @rifqiinfll sebagai ucapan terima kasih")

# Download Template
with open("Dataset112.xlsx", "rb") as f:
    st.download_button("📥 Download Template Excel", f, file_name="Dataset112.xlsx")

# ---------------------- FITUR INPUT OTOMATIS -----------------------
st.subheader("⚙️ Input Otomatis Intervensi")
uploaded_file = st.file_uploader("🔽 Upload Excel Intervensi", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 Mulai Proses Input"):
        with st.spinner("Sedang memproses..."):
            logs, output, output_filename = process_excel(uploaded_file)
        st.success("Proses selesai.")
        if output:
            st.download_button(
                "📥 Download NIK Tidak Ditemukan (Input)",
                output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.text_area("Log Proses", "\n".join(logs), height=300)

# ---------------------- FITUR CEK NIK -----------------------
st.subheader("🔍 Cek Ketersediaan NIK")
cek_file = st.file_uploader("🔽 Upload Excel untuk Cek NIK", type=["xlsx"], key="cek")

if cek_file:
    if st.button("🔎 Mulai Cek NIK"):
        with st.spinner("Sedang mengecek NIK..."):
            logs, output, output_filename = cek_nik_excel(cek_file)
        st.success("Cek NIK selesai.")
        if output:
            st.download_button(
                "📥 Download NIK Tidak Ditemukan (Cek)",
                output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.text_area("Log Cek NIK", "\n".join(logs), height=300)

# ---------------------- HISTORY LOG -----------------------
# Mencari file log dengan ekstensi .txt di folder logs
log_files = sorted([f for f in os.listdir("logs") if f.endswith(".txt")], reverse=True)

if log_files:
    # Menggunakan session_state untuk menyimpan pilihan log yang dipilih
    if (
        "selected_log" not in st.session_state
        or st.session_state.selected_log not in log_files
    ):
        st.session_state.selected_log = log_files[
            0
        ]  # Set default log jika tidak ada atau sudah dihapus

    # Membuat pilihan untuk memilih log
    selected_log = st.selectbox(
        "Pilih Log untuk Dilihat:",
        log_files,
        index=log_files.index(
            st.session_state.selected_log
        ),  # Menyimpan log yang terakhir dipilih
    )

    # Menyimpan pilihan log ke session_state agar tetap diingat
    st.session_state.selected_log = selected_log

    if selected_log:
        # Menampilkan detail log berdasarkan file yang dipilih
        with open(os.path.join("logs", selected_log), "r") as f:
            log_content = f.read()

        st.text_area(f"Log untuk {selected_log}", log_content, height=400)

        # Hapus log
        if st.button("🗑️ Hapus Log"):
            if selected_log:
                os.remove(os.path.join("logs", selected_log))
                st.success(f"Log {selected_log} telah dihapus.")
                log_files.remove(selected_log)  # Remove the deleted file from the list
                # Reset selected log to avoid errors
                if log_files:
                    st.session_state.selected_log = log_files[0]
                else:
                    st.session_state.selected_log = None
            else:
                st.error("Pilih log terlebih dahulu untuk dihapus.")
else:
    st.warning("Tidak ada file log tersedia.")
