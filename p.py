import tkinter as tk
from tkinter import messagebox
import threading
import queue
import time
import csv
import os
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Queue untuk menyimpan log dari thread automasi
log_queue = queue.Queue()

# Variabel global untuk driver
driver = None

# Fungsi untuk menampilkan log dalam GUI secara real-time
def update_log():
    try:
        while True:
            message = log_queue.get_nowait()
            log_text.config(state=tk.NORMAL)
            log_text.insert(tk.END, message + "\n")
            log_text.config(state=tk.DISABLED)
            log_text.yview(tk.END)  # Scroll ke bawah
    except queue.Empty:
        pass
    root.after(100, update_log)

# Fungsi untuk menambahkan log ke dalam queue
def log_to_gui(message):
    log_queue.put(message)

# Fungsi untuk memulai automasi
def start_automation():
    start_button.config(state=tk.DISABLED)
    stop_button.config(state=tk.NORMAL)
    log_text.delete(1.0, tk.END)
    threading.Thread(target=run_automation, daemon=True).start()

# Fungsi untuk menghentikan automasi
def stop_automation():
    global driver
    if driver:
        driver.quit()
    stop_button.config(state=tk.DISABLED)
    start_button.config(state=tk.NORMAL)
    log_to_gui("Automasi dihentikan oleh pengguna.")

# Fungsi utama untuk menjalankan automasi
def run_automation():
    global driver
    spk_gagal = []  # Menyimpan SPK yang gagal
    try:
        directory_path = "C:/autoSPK/"
        config_file_path = os.path.join(directory_path, "config.json")
        default_csv_path = os.path.join(directory_path, "spk_list.csv")

        if not os.path.exists(directory_path):
            os.makedirs(directory_path)
            log_to_gui(f"Direktori {directory_path} berhasil dibuat.")

        if os.path.exists(config_file_path):
            with open(config_file_path, 'r') as config_file:
                config = json.load(config_file)
            username = config.get("username")
            password = config.get("password")
            file_csv = config.get("csv_path", default_csv_path)
        else:
            log_to_gui("File konfigurasi tidak ditemukan.")
            return

        if not username or not password:
            log_to_gui("Username atau password tidak ditemukan dalam konfigurasi.")
            return

        if not os.path.exists(file_csv):
            log_to_gui(f"File CSV tidak ditemukan di {file_csv}.")
            return

        try:
            with open(file_csv, newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                spk_list = [row["No_SP"] for row in reader]
        except Exception as e:
            log_to_gui(f"Error membaca file CSV: {e}")
            return

        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

        driver.get("https://mobile.acl.co.id/aclweb/#/login")
        log_to_gui("Berhasil membuka halaman login.")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "userName"))).send_keys(username)
        driver.find_element(By.ID, "password").send_keys(password)
        log_to_gui("Berhasil memasukkan username dan password.")

        submit_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[@class='ui blue fluid vertical animated button']"))
        )
        driver.execute_script("arguments[0].click();", submit_button)
        log_to_gui("Berhasil login.")

        time.sleep(3)
        driver.get("https://mobile.acl.co.id/aclweb/#/penagihanspk")
        log_to_gui("Berhasil menuju halaman Penagihan SPK.")

        try:
            search_box = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//input[@placeholder='cari no. spk..']"))
            )
            log_to_gui("Kolom pencarian 'No. SPK' ditemukan.")
        except Exception as e:
            log_to_gui(f"Error: Kolom pencarian 'No. SPK' tidak ditemukan. Detail: {e}")
            return

        jumlah_dicentang = 0
        for nomor_spk in spk_list:
            try:
                search_box.clear()
                search_box.send_keys(nomor_spk)
                log_to_gui(f"Memasukkan nomor SPK {nomor_spk} ke dalam kolom pencarian.")
                
                search_box.send_keys(Keys.ENTER)
                time.sleep(5)

                spk_element = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[2]/div/div[2]/div[2]/div/div/table/tbody/tr/td[4]'))
                )

                if spk_element.text.strip() == nomor_spk:
                    checkbox = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div[2]/div/div[2]/div[2]/div/div/table/tbody/tr/td[2]/div/input'))
                    )
                    if not checkbox.is_selected():
                        driver.execute_script("arguments[0].click();", checkbox)
                        log_to_gui(f"Checkbox untuk nomor SPK {nomor_spk} berhasil dicentang.")
                        jumlah_dicentang += 1
                    else:
                        log_to_gui(f"Checkbox untuk nomor SPK {nomor_spk} sudah tercentang.")
                else:
                    log_to_gui(f"Checkbox ditemukan untuk nomor SPK {spk_element.text.strip()}, tetapi tidak sesuai dengan {nomor_spk}.")
                    spk_gagal.append(nomor_spk)  # Menyimpan SPK yang gagal dicentang
                time.sleep(1)

            except Exception as e:
                log_to_gui(f"Error pada nomor SPK {nomor_spk}: {e}")
                spk_gagal.append(nomor_spk)  # Menyimpan SPK yang gagal dicentang

        total_spk = len(spk_list)
        persentase_berhasil = (jumlah_dicentang / total_spk) * 100

        # Tampilkan pemberitahuan setelah selesai centang
        message = f"Automasi selesai! {jumlah_dicentang} SPK berhasil dicentang dari {total_spk} SPK ({persentase_berhasil:.2f}%)"
        if spk_gagal:
            message += f"\nSPK yang gagal dicentang: {', '.join(spk_gagal)}"
        
        messagebox.showinfo("Selesai", message)
        log_to_gui(f"Input sukses! Total SPK yang dicentang: {jumlah_dicentang} dari {total_spk} ({persentase_berhasil:.2f}%)")
        if spk_gagal:
            log_to_gui(f"SPK yang gagal dicentang: {', '.join(spk_gagal)}")
        log_to_gui("Automasi selesai. Browser akan tetap terbuka.")

    except Exception as e:
        log_to_gui(f"Kesalahan umum dalam skrip: {e}")
        messagebox.showerror("Error", f"Terjadi kesalahan: {e}")

# Fungsi untuk memuat nomor SPK dari file CSV ke dalam input
def load_spk_from_csv(spk_input_text):
    directory_path = "C:/autoSPK/"
    file_csv = os.path.join(directory_path, "spk_list.csv")
    if os.path.exists(file_csv):
        try:
            with open(file_csv, newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                spk_list = [row["No_SP"] for row in reader]
            # Masukkan nomor SPK ke dalam widget Text
            spk_input_text.delete(1.0, tk.END)  # Clear existing content
            spk_input_text.insert(tk.END, "\n".join(spk_list))  # Load SPK list
        except Exception as e:
            log_to_gui(f"Error membaca file CSV: {e}")
    else:
        log_to_gui("File CSV tidak ditemukan, SPK kosong.")

# Fungsi untuk membuka input SPK
def open_spk_input():
    spk_input_window = tk.Toplevel(root)
    spk_input_window.title("Input SPK")

    # Frame untuk input nomor SPK
    frame_spk_input = tk.Frame(spk_input_window)
    frame_spk_input.pack(padx=20, pady=20)

    spk_label = tk.Label(frame_spk_input, text="Masukkan No SPK:")
    spk_label.pack(pady=5)
    spk_input_text = tk.Text(frame_spk_input, height=10, width=40)
    spk_input_text.pack(pady=5)

    # Memuat nomor SPK yang tersimpan
    load_spk_from_csv(spk_input_text)

    # Tombol untuk menyimpan nomor SPK
    save_button = tk.Button(frame_spk_input, text="Simpan SPK", command=lambda: save_spk_to_csv(spk_input_text.get("1.0", tk.END).strip()))
    save_button.pack(pady=10)

# Fungsi untuk menyimpan nomor SPK ke file CSV
def save_spk_to_csv(spk_data):
    spk_list = spk_data.splitlines()
    directory_path = "C:/autoSPK/"
    file_csv = os.path.join(directory_path, "spk_list.csv")
    
    try:
        with open(file_csv, mode='w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["No_SP"])
            for spk in spk_list:
                writer.writerow([spk])
        log_to_gui(f"Data SPK berhasil disimpan di {file_csv}")
    except Exception as e:
        log_to_gui(f"Error menyimpan file CSV: {e}")

# Fungsi untuk mengganti username dan password login
def change_user_password():
    # Memuat username dan password lama jika ada
    directory_path = "C:/autoSPK/"
    config_file_path = os.path.join(directory_path, "config.json")
    username = password = ""
    
    if os.path.exists(config_file_path):
        try:
            with open(config_file_path, 'r') as config_file:
                config = json.load(config_file)
                username = config.get("username", "")
                password = config.get("password", "")
        except Exception as e:
            log_to_gui(f"Error membaca konfigurasi: {e}")
    
    # Menampilkan dialog input untuk username dan password baru
    def save_new_credentials():
        new_username = username_entry.get()
        new_password = password_entry.get()
        directory_path = "C:/autoSPK/"
        config_file_path = os.path.join(directory_path, "config.json")
        if not os.path.exists(directory_path):
            os.makedirs(directory_path)
        config = {"username": new_username, "password": new_password}
        try:
            with open(config_file_path, 'w') as config_file:
                json.dump(config, config_file)
            log_to_gui("Username dan password berhasil diperbarui.")
            change_window.destroy()
        except Exception as e:
            log_to_gui(f"Error menyimpan konfigurasi: {e}")

    change_window = tk.Toplevel(root)
    change_window.title("Ganti Username & Password")

    # Frame untuk input username dan password baru
    frame = tk.Frame(change_window)
    frame.pack(padx=20, pady=20)

    username_label = tk.Label(frame, text="Username:")
    username_label.pack(pady=5)
    username_entry = tk.Entry(frame)
    username_entry.insert(0, username)  # Memasukkan username lama
    username_entry.pack(pady=5)

    password_label = tk.Label(frame, text="Password:")
    password_label.pack(pady=5)
    password_entry = tk.Entry(frame, show="*")
    password_entry.insert(0, password)  # Memasukkan password lama
    password_entry.pack(pady=5)

    save_button = tk.Button(frame, text="Simpan", command=save_new_credentials)
    save_button.pack(pady=10)

# Membuat window utama
root = tk.Tk()
root.title("Automasi SPK")

# Frame untuk tombol utama
button_frame = tk.Frame(root)
button_frame.pack(pady=20, padx=10, fill=tk.X)

# Tombol untuk membuka input SPK (di sebelah kiri)
spk_button = tk.Button(button_frame, text="Input SPK", command=open_spk_input)
spk_button.pack(side=tk.LEFT, padx=10)

# Tombol untuk mengganti user dan password login
change_user_button = tk.Button(button_frame, text="Ganti User & Password", command=change_user_password)
change_user_button.pack(side=tk.LEFT, padx=10)

# Frame untuk tombol start/stop
control_frame = tk.Frame(root)
control_frame.pack(pady=20, fill=tk.X)

# Tombol untuk memulai automasi
start_button = tk.Button(control_frame, text="Start Automasi", command=start_automation)
start_button.pack(side=tk.TOP, pady=10)  # Posisikan di atas

# Tombol untuk menghentikan automasi
stop_button = tk.Button(control_frame, text="Stop Automasi", state=tk.DISABLED, command=stop_automation)
stop_button.pack(side=tk.TOP, pady=10)  # Posisikan di bawah tombol start

# Menampilkan log automasi
log_text = tk.Text(root, height=20, width=80, state=tk.DISABLED)
log_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Jalankan update log secara periodik
root.after(100, update_log)

# Jalankan aplikasi Tkinter
root.mainloop()
