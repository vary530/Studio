import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox
from PIL import Image, ImageTk

class CoordinateGrabber:
    def __init__(self, root):
        self.root = root
        self.root.title("📍 物調表座標抓取小工具")
        
        self.coords = {}
        
        # 1. 選擇空白底圖
        messagebox.showinfo("第一步", "請選擇您的「空白物調表底圖 (JPG/PNG)」")
        self.file_path = filedialog.askopenfilename(
            title="選擇空白物調表底圖", 
            filetypes=[("Image files", "*.jpg *.jpeg *.png")]
        )
        
        if not self.file_path:
            self.root.destroy()
            return

        # 2. 載入圖片
        self.image = Image.open(self.file_path)
        self.tk_image = ImageTk.PhotoImage(self.image)

        # 3. 建立可滾動的畫布 (對付大圖片)
        self.canvas = tk.Canvas(root, width=1000, height=700, scrollregion=(0, 0, self.image.width, self.image.height))
        
        self.hbar = tk.Scrollbar(root, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.hbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.vbar = tk.Scrollbar(root, orient=tk.VERTICAL, command=self.canvas.yview)
        self.vbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.canvas.config(xscrollcommand=self.hbar.set, yscrollcommand=self.vbar.set)
        self.canvas.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_image)
        
        # 綁定滑鼠左鍵點擊事件
        self.canvas.bind("<Button-1>", self.on_click)

        # 綁定關閉視窗事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_click(self, event):
        # 取得考慮到滾動條的真實座標
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        
        # 彈出對話框詢問欄位名稱
        field_name = simpledialog.askstring(
            "輸入欄位名稱", 
            f"您點擊了座標 ({int(x)}, {int(y)})\n\n請輸入這個位置的「欄位名稱」\n(例如：案名、地址、售價)："
        )
        
        if field_name:
            self.coords[field_name] = (int(x), int(y))
            # 在圖上畫個紅點和藍字做記號，避免重複點擊
            self.canvas.create_oval(x-4, y-4, x+4, y+4, fill="red", outline="red")
            self.canvas.create_text(x+10, y, text=field_name, anchor=tk.W, fill="blue", font=("Arial", 12, "bold"))
            print(f"✅ 已記錄: '{field_name}' -> 座標 ({int(x)}, {int(y)})")

    def on_closing(self):
        # 當關閉視窗時，把結果漂亮地印在終端機
        print("\n" + "="*50)
        print("🎉 座標抓取完成！請將下方的字典複製起來備用：\n")
        print("COORDS_DICT = {")
        for k, v in self.coords.items():
            print(f"    '{k}': {v},")
        print("}")
        print("="*50 + "\n")
        messagebox.showinfo("完成", "座標已印出在您執行程式的「終端機(黑畫面)」視窗中，請去複製！")
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = CoordinateGrabber(root)
    root.mainloop()