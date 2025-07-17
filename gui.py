import tkinter as tk
from tkinter import messagebox, scrolledtext
import subprocess
# รันการทดสอบ pytest และแสดงผลลัพท์ในกล่องข้อความ
def run_tests():
    try:
        result = subprocess.run(["python", "-m", "pytest", "new2.py", "-v", "--headed"], capture_output=True, text=True)
        output_text.config(state=tk.NORMAL)
        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, result.stdout)
        output_text.config(state=tk.DISABLED)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to run tests: {e}")

# สร้างหน้าต่างหลัก
root = tk.Tk()
root.title("หมดชีวิตฉันให้เธอ")
root.geometry("600x400")
root.configure(bg="#DCDCDC")

# ส่วนหัว
header_label = tk.Label(root, text="Python TestSheet Runner", font=("Arial", 14, "bold"), bg="#f0f0f0")
header_label.pack(pady=10)

# เฟรมสำหรับปุ่ม
button_frame = tk.Frame(root, bg="#f0f0f0")
button_frame.pack(pady=5)

run_button = tk.Button(button_frame, text="วิ่งแบบพี่ตูน", command=run_tests, font=("Arial", 18), bg="#4CAF50", fg="white", padx=10, pady=5)
run_button.pack()

# กล่องแสดงผลลัพธ์
output_text = scrolledtext.ScrolledText(root, height=15, width=70, wrap=tk.WORD, font=("Courier", 10), bg="black", fg="green")
output_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
output_text.config(state=tk.DISABLED)

# เริ่ม GUI
root.mainloop()
