import os
import sys
import threading
import io
import zipfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageFile

Image.MAX_IMAGE_PIXELS = None
ImageFile.LOAD_TRUNCATED_IMAGES = True


def compress_pptx(pptx_path, progress_callback=None):
    
    Args:
        pptx_path: PPT文件路径
        progress_callback: 进度回调函数(current, total)
        
    Returns:
        压缩后的文件字节数据
    """
    output_stream = io.BytesIO()
    
    try:
        with zipfile.ZipFile(pptx_path, 'r') as original_zip:
            all_files = original_zip.namelist()
            media_files = [f for f in all_files if f.startswith('ppt/media/')]
            other_files = [f for f in all_files if not f.startswith('ppt/media/')]
            
            total_files = len(all_files)
            processed_files = 0
            
            with zipfile.ZipFile(output_stream, 'w', zipfile.ZIP_DEFLATED) as new_zip:
                for file_path in other_files:
                    try:
                        file_content = original_zip.read(file_path)
                        new_zip.writestr(file_path, file_content)
                    except Exception as e:
                        print(f"处理非媒体文件时出错: {str(e)}")
                        continue
                    
                    processed_files += 1
                    if progress_callback:
                        progress_callback(processed_files, total_files)
                
                for media_path in media_files:
                    if media_path.lower().endswith(('.jpg', '.jpeg', '.png')):
                        try:
                            image_data = original_zip.read(media_path)
                            image_stream = io.BytesIO(image_data)
                            
                            image = Image.open(image_stream)
                            
                            img_format = image.format
                            if img_format == 'PNG':
                                if image.mode == 'RGBA':
                                    background = Image.new('RGB', image.size, (255, 255, 255))
                                    background.paste(image, mask=image.split()[3])
                                    image = background
                                img_format = 'JPEG'
                                media_path = media_path.rsplit('.', 1)[0] + '.jpg'
                            
                            max_size = 1920
                            width, height = image.size
                            if width > max_size or height > max_size:
                                ratio = max_size / max(width, height)
                                new_width = int(width * ratio)
                                new_height = int(height * ratio)
                                image = image.resize((new_width, new_height), Image.BILINEAR)
                            
                            compressed_stream = io.BytesIO()
                            if img_format == 'JPEG':
                                image.save(compressed_stream, format='JPEG', quality=85, optimize=True)
                            else:
                                image.save(compressed_stream, format=img_format, compress_level=9)
                            
                            compressed_data = compressed_stream.getvalue()
                            
                            new_zip.writestr(media_path, compressed_data)
                            
                            image.close()
                            image_stream.close()
                            compressed_stream.close()
                        except Exception as e:
                            print(f"处理图片出错: {e}")
                            try:
                                file_content = original_zip.read(media_path)
                                new_zip.writestr(media_path, file_content)
                            except:
                                pass
                    else:
                        try:
                            file_content = original_zip.read(media_path)
                            new_zip.writestr(media_path, file_content)
                        except:
                            pass
                    
                    processed_files += 1
                    if progress_callback:
                        progress_callback(processed_files, total_files)
    except Exception as e:
        print(f"压缩过程出错: {str(e)}")
        pass
    
    return output_stream.getvalue()


class PPTCompressorApp:
    """PPT压缩工具应用程序类 - Windows 7优化版本"""
    
    def __init__(self, root):
        """初始化"""
        self.root = root
        self.root.title("PPT压缩工具 - Windows 7 专用版 【预览版本】")
        self.root.geometry("600x510")
        self.root.resizable(True, True)
        
        self.font = ("SimHei", 10)
        
        self.file_path = tk.StringVar()
        self.original_size = 0
        self.compressed_data = None
        self.compression_in_progress = False
        self.user_confirmed = False
        
        self.create_ui()
    
    def create_ui(self):
        """用户界面"""
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        file_frame = tk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(file_frame, text="文件:", font=self.font).pack(side=tk.LEFT, padx=5)
        tk.Entry(file_frame, textvariable=self.file_path, width=45, font=self.font).pack(side=tk.LEFT, padx=5)
        tk.Button(file_frame, text="浏览...", command=self.browse_file, width=10).pack(side=tk.LEFT, padx=5)
        
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        self.compress_btn = tk.Button(btn_frame, text="开始压缩", command=self.start_compression, width=15)
        self.compress_btn.pack(side=tk.RIGHT, padx=10)
        
        self.save_btn = tk.Button(btn_frame, text="保存文件", command=self.save_compressed_file, width=15, state=tk.DISABLED)
        self.save_btn.pack(side=tk.RIGHT, padx=10)
        
        # 关于按钮
        tk.Button(btn_frame, text="关于", command=self.show_about, width=10).pack(side=tk.LEFT, padx=10)
        
        # 进度条
        progress_frame = tk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(progress_frame, text="压缩进度:", font=self.font).pack(anchor=tk.W, pady=2)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, length=500, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=2)
        
        self.progress_label = tk.Label(progress_frame, text="准备就绪", font=self.font)
        self.progress_label.pack(anchor=tk.W, pady=2)
        
        # 结果显示
        result_frame = tk.Frame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.original_size_var = tk.StringVar(value="原始文件大小: 未选择文件")
        tk.Label(result_frame, textvariable=self.original_size_var, font=self.font).pack(anchor=tk.W, pady=2)
        
        self.compressed_size_var = tk.StringVar(value="压缩后大小: 未压缩")
        tk.Label(result_frame, textvariable=self.compressed_size_var, font=self.font).pack(anchor=tk.W, pady=2)
        
        self.ratio_var = tk.StringVar(value="压缩率: 未计算")
        tk.Label(result_frame, textvariable=self.ratio_var, font=self.font).pack(anchor=tk.W, pady=2)
    
    def show_about(self):
        """窗口"""
        # 使用简单的messagebox显示关于信息，减少性能占用
        about_message = (
            "PPT压缩工具 - Windows 7专用版 【预览版本】\n\n"
            "版本: 1.0.0_win7-AlphaTest\n"
            "开发者: Yuyudifiesh\n\n"
            "本工具用于压缩PPT文件，解决在Windows 7系统下\n"
            "PPT文件打开时卡顿、无响应的问题。"
        )
        
        messagebox.showinfo("关于", about_message)
    
    def browse_file(self):
        """浏览并选择PPT文件"""
        file_types = ["PowerPoint文件", "*.pptx"]
        path = filedialog.askopenfilename(filetypes=[("PowerPoint文件", "*.pptx")])
        if path:
            self.file_path.set(path)
            self.original_size = os.path.getsize(path)
            self.original_size_var.set(f"原始文件大小: {self.original_size/1024/1024:.2f} MB")
            
            # 重置结果显示
            self.compressed_size_var.set("压缩后大小: 未压缩")
            self.ratio_var.set("压缩率: 未计算")
            self.save_btn.config(state=tk.DISABLED)
            
            # 重置进度条
            self.progress_var.set(0)
            self.progress_label.config(text="准备就绪")
    
    def update_progress(self, current, total):
        """更新进度条"""
        progress = (current / total) * 100
        self.progress_var.set(progress)
        self.progress_label.config(text=f"处理中: {progress:.1f}%")
    
    def start_compression(self):
        """开始压缩文件"""
        file_path = self.file_path.get()
        
        if not file_path:
            messagebox.showwarning("警告", "请先选择一个PPT文件")
            return
        
        if not os.path.isfile(file_path):
            messagebox.showwarning("警告", "选择的文件不存在")
            return
        
        if not file_path.lower().endswith('.pptx'):
            messagebox.showwarning("警告", "请选择PPPTX格式的文件")
            return
        
        warning_message = "程序将执行压缩操作\n\n建议先备份原始文件\n处理大文件可能需要较长时间"
        
        result = messagebox.askyesno("确认压缩", warning_message)
        if not result:
            return  # 用户取消操作
        
        self.compress_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.progress_label.config(text="开始处理...")
        self.compression_in_progress = True
        
        # 单线程压缩
        compression_thread = threading.Thread(target=self.perform_compression, args=(file_path,))
        compression_thread.daemon = True
        compression_thread.start()
    
    def perform_compression(self, file_path):
        """在后台线程中执行压缩操作"""
        try:
            # 执行压缩
            self.compressed_data = compress_pptx(file_path, self.update_progress)
            
            # 计算结果
            compressed_size = len(self.compressed_data)
            compression_ratio = (1 - compressed_size / self.original_size) * 100
            
            self.root.after(0, self.update_compression_result, compressed_size, compression_ratio)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"压缩失败: {str(e)}"))
            self.root.after(0, lambda: self.progress_label.config(text="压缩失败"))
        finally:
            self.compression_in_progress = False
            self.root.after(0, lambda: self.compress_btn.config(state=tk.NORMAL))
    
    def update_compression_result(self, compressed_size, compression_ratio):
        """更新压缩结果显示"""
        self.compressed_size_var.set(f"压缩后大小: {compressed_size/1024/1024:.2f} MB")
        self.ratio_var.set(f"压缩率: {compression_ratio:.2f}%")
        self.progress_label.config(text="压缩完成!")
        self.save_btn.config(state=tk.NORMAL)
        
        messagebox.showinfo("成功", f"PPT压缩完成！压缩率: {compression_ratio:.2f}%")
    
    def save_compressed_file(self):
        """保存压缩后的文件"""
        if not self.compressed_data:
            messagebox.showwarning("警告", "没有压缩后的文件数据")
            return
        
        original_path = self.file_path.get()
        if not original_path:
            messagebox.showwarning("警告", "找不到原始文件路径")
            return
            
        try:
            original_name = os.path.basename(original_path)
            name_without_ext = os.path.splitext(original_name)[0]
            
            default_save_path = os.path.join(os.path.dirname(original_path), f"{name_without_ext}_compressed.pptx")
            
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pptx",
                filetypes=[("PowerPoint文件", "*.pptx")],
                initialfile=f"{name_without_ext}_compressed.pptx",
                initialdir=os.path.dirname(original_path)
            )
            
            if save_path:
                try:
                    if isinstance(self.compressed_data, bytes):
                        with open(save_path, "wb") as f:
                            f.write(self.compressed_data)
                        
                        messagebox.showinfo("成功", f"文件已保存至:\n{save_path}")
                    else:
                        messagebox.showerror("错误", "压缩数据格式无效")
                except IOError as e:
                    messagebox.showerror("错误", f"保存文件失败: 无法写入文件\n{str(e)}")
                except Exception as e:
                    messagebox.showerror("错误", f"保存文件失败: {str(e)}")
        except Exception as e:
            messagebox.showerror("错误", f"处理文件路径时出错: {str(e)}")


if __name__ == "__main__":
    # 创建应用程序窗口并运行
    root = tk.Tk()
    app = PPTCompressorApp(root)
    root.mainloop()
