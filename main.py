import streamlit as st
import pandas as pd
import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


def setup_driver():
    options = webdriver.ChromeOptions()
    return webdriver.Chrome(options=options)


def process_excel(file):
    log = []
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"NIK_Tidak_Ditemukan_{timestamp}.xlsx"
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
                    log.append(f"✔️ Data ditemukan: {nik}")
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

                    log.append(f"✅ Intervensi ditambahkan untuk: {nik}")
                    driver.find_element(By.LINK_TEXT, "Data UMKM").click()
                    time.sleep(2)
                else:
                    log.append(f"❌ Tidak ditemukan: {nik}")
                    not_found_niks.append(nik)
            except Exception as e:
                log.append(f"⚠️ Error saat proses NIK {nik}: {str(e)}")
                not_found_niks.append(nik)
                continue

    except Exception as e:
        log.append(f"🚨 ERROR GLOBAL: {str(e)}")
    finally:
        if not os.path.exists("logs"):
            os.makedirs("logs")
        log_path = f"logs/{output_filename}"
        if not_found_niks:
            pd.DataFrame({"NIK Tidak Ditemukan": not_found_niks}).to_excel(
                log_path, index=False
            )
        with open(f"logs/log_{timestamp}.txt", "w") as f:
            for line in log:
                f.write(line + "\n")
        return log, log_path


# ---------------------- STREAMLIT APP -----------------------
st.set_page_config(page_title="Input Otomatis Intervensi", layout="centered")
st.title("📤 Input Otomatis Intervensi UMKM")
st.markdown("Credit : Ripal")
st.markdown(
    "Follow Instagram @rifqiinfll sebagai ucapan terima kasih telah mempermudah magang kalian"
)
# Download Template
with open("Dataset112.xlsx", "rb") as f:
    st.download_button("📥 Download Template Excel", f, file_name="Dataset112.xlsx")

uploaded_file = st.file_uploader("📂 Upload Excel Intervensi", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 Mulai Proses Input"):
        with st.spinner("Sedang memproses..."):
            logs, error_path = process_excel(uploaded_file)
        st.success("✅ Proses selesai.")
        st.download_button(
            "📥 Download Laporan Error / NIK Tidak Ditemukan",
            open(error_path, "rb"),
            file_name=os.path.basename(error_path),
        )
        st.text_area("📄 Log Proses", "\n".join(logs), height=300)

# History log
if st.button("📚 Lihat History Log"):
    log_files = sorted(
        [f for f in os.listdir("logs") if f.endswith(".txt")], reverse=True
    )
    for file in log_files:
        with open(os.path.join("logs", file), "r") as f:
            st.text_area(f"📄 {file}", f.read(), height=200)
