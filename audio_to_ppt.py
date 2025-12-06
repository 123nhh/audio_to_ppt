import os
import re
import time
import uuid  # 新增：用于生成唯一文件名
import threading  # 新增：用于打印锁
from concurrent.futures import ThreadPoolExecutor, as_completed # 新增：线程池
from io import BytesIO
from PIL import Image, ImageFilter, ImageEnhance
from mutagen import File
from mutagen.flac import FLAC
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# 手动定义常量
PP_ALIGN_LEFT = 1
PP_ALIGN_CENTER = 2
MSO_ANCHOR_MIDDLE = 3
MSO_AUTO_SIZE_NONE = 0
PP_LINE_SPACING_EXACTLY = 4

# 全局打印锁，防止多线程打印时文字错乱
print_lock = threading.Lock()

def safe_print(msg):
    with print_lock:
        print(msg)

class AudioToPPT:
    def __init__(self, audio_path, output_ppt_path):
        self.audio_path = audio_path
        self.output_ppt_path = os.path.abspath(output_ppt_path)
        
        # --- 🛡️ 关键修改：生成唯一会话ID ---
        # 防止不同线程操作同一个 temp_bg.jpg 导致冲突
        self.uid = uuid.uuid4().hex[:8] 
        self.temp_bg = f"temp_bg_{self.uid}.jpg"
        self.temp_cover = f"temp_cover_{self.uid}.jpg"
        self.temp_mask_top = f"temp_mask_top_{self.uid}.jpg"
        self.temp_mask_bottom = f"temp_mask_bottom_{self.uid}.jpg"
        
        self.metadata = {
            'title': '未知标题',
            'artist': '未知歌手',
            'lyrics': [],
            'cover_data': None
        }
        
        # --- 布局参数 ---
        self.SLIDE_W = Inches(13.333) 
        self.SLIDE_H = Inches(7.5)
        
        self.VISIBLE_LINES = 5          
        self.FIXED_LINE_HEIGHT_PT = 60  
        self.line_height_in = Inches(self.FIXED_LINE_HEIGHT_PT / 72.0) 
        
        self.NORMAL_FONT_SIZE = 36      
        self.SMALL_FONT_SIZE = 24       
        self.LONG_LINE_THRESHOLD = 14   
        
        self.LYRIC_LEFT = Inches(5.5)   
        self.LYRIC_WIDTH = Inches(7.5)  
        
        self.window_height = self.line_height_in * self.VISIBLE_LINES
        self.window_top = (self.SLIDE_H - self.window_height) / 2
        
        self.mask_top_h = self.window_top
        self.mask_bottom_h = self.SLIDE_H - (self.window_top + self.window_height)

    def clean_lyrics(self, raw_text):
        cleaned_lines = []
        lines = raw_text.split('\n')
        pattern = re.compile(r'\[\d{1,3}:\d{2}(?:\.\d{1,3})?\]')
        for line in lines:
            line_content = re.sub(pattern, '', line).strip()
            if not line_content: continue
            cleaned_lines.append(line_content)
        return cleaned_lines

    def extract_metadata(self):
        try:
            audio = File(self.audio_path)
            tags = audio.tags
            if tags:
                self.metadata['title'] = str(tags.get('TITLE', tags.get('TIT2', ['未知标题']))[0])
                self.metadata['artist'] = str(tags.get('ARTIST', tags.get('TPE1', ['未知歌手']))[0])
                
                raw_lyrics = ""
                if isinstance(audio, FLAC):
                    raw_lyrics = tags.get('lyrics', tags.get('unsyncedlyrics', ['']))[0]
                elif tags and hasattr(tags, 'getall'): 
                     uslt = tags.getall('USLT')
                     if uslt: raw_lyrics = uslt[0].text

                if raw_lyrics:
                    self.metadata['lyrics'] = self.clean_lyrics(raw_lyrics)
                
                if isinstance(audio, FLAC) and audio.pictures:
                    self.metadata['cover_data'] = audio.pictures[0].data
                elif hasattr(audio, 'tags') and 'APIC:' in audio.tags: 
                     for key in audio.tags.keys():
                         if key.startswith('APIC'):
                             self.metadata['cover_data'] = audio.tags[key].data
                             break
        except Exception as e:
            safe_print(f"      [警告] 元数据读取可能有误: {e}")

    def prepare_images(self):
        if not self.metadata['cover_data']: return None
        try:
            img = Image.open(BytesIO(self.metadata['cover_data'])).convert("RGB")
            
            # 背景处理
            bg_img = img.filter(ImageFilter.GaussianBlur(radius=40))
            bg_img = ImageEnhance.Brightness(bg_img).enhance(0.5) 
            
            target_w, target_h = 1280, 720
            bg_img = bg_img.resize((target_w, target_h))
            
            # 使用带 ID 的文件名保存
            bg_img.save(self.temp_bg)
            img.save(self.temp_cover)
            
            scale_y = target_h / self.SLIDE_H
            px_mask_top = int(self.mask_top_h * scale_y)
            px_mask_bottom_start = int((self.window_top + self.window_height) * scale_y)
            
            if px_mask_top < 1: px_mask_top = 1
            if px_mask_bottom_start >= target_h: px_mask_bottom_start = target_h - 1
            
            mask_top_img = bg_img.crop((0, 0, target_w, px_mask_top))
            mask_top_img.save(self.temp_mask_top)
            
            mask_bottom_img = bg_img.crop((0, px_mask_bottom_start, target_w, target_h))
            mask_bottom_img.save(self.temp_mask_bottom)
            
            return True
        except Exception as e:
            safe_print(f"      [跳过] 图片处理失败 (可能是封面图损坏): {e}")
            return False

    def generate_ppt(self):
        prs = Presentation()
        prs.slide_width = self.SLIDE_W
        prs.slide_height = self.SLIDE_H

        if not self.prepare_images():
            safe_print("      [跳过] 无法生成必要图片资源。")
            return

        lyrics = self.metadata['lyrics']
        if not lyrics:
            lyrics = ["(纯音乐或未检测到歌词)"]

        padding_count = self.VISIBLE_LINES // 2
        total_text_height = self.line_height_in * len(lyrics)

        for i in range(len(lyrics)):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            
            # 使用带 ID 的临时文件
            slide.shapes.add_picture(self.temp_bg, 0, 0, width=self.SLIDE_W, height=self.SLIDE_H)

            base_y = self.window_top
            current_top = base_y + (padding_count * self.line_height_in) - (i * self.line_height_in)
            
            safe_height = max(total_text_height * 2, Inches(1))
            
            lyric_box = slide.shapes.add_textbox(self.LYRIC_LEFT, current_top, self.LYRIC_WIDTH, safe_height)
            tf = lyric_box.text_frame
            tf.word_wrap = False 
            tf.auto_size = MSO_AUTO_SIZE_NONE 
            tf.clear()
            
            for line_text in lyrics:
                p = tf.add_paragraph()
                p.text = line_text
                p.font.bold = True
                p.font.name = "微软雅黑"
                p.font.color.rgb = RGBColor(255, 255, 255)
                p.alignment = PP_ALIGN_LEFT 
                
                if len(line_text) > self.LONG_LINE_THRESHOLD:
                    p.font.size = Pt(self.SMALL_FONT_SIZE)
                else:
                    p.font.size = Pt(self.NORMAL_FONT_SIZE)
                
                p.line_spacing_rule = PP_LINE_SPACING_EXACTLY 
                p.line_spacing = Pt(self.FIXED_LINE_HEIGHT_PT)

            try:
                slide.shapes.add_picture(self.temp_mask_top, 0, 0, width=self.SLIDE_W, height=self.mask_top_h)
                slide.shapes.add_picture(self.temp_mask_bottom, 0, self.window_top + self.window_height, width=self.SLIDE_W, height=self.mask_bottom_h)
            except Exception as e:
                safe_print(f"      [警告] 遮罩添加失败: {e}")

            img_size = Inches(4.0)
            slide.shapes.add_picture(self.temp_cover, Inches(1.2), Inches(1.5), width=img_size, height=img_size)
            
            tx = slide.shapes.add_textbox(Inches(1.2), Inches(1.5) + img_size + Inches(0.2), img_size, Inches(1.5))
            p1 = tx.text_frame.add_paragraph()
            p1.text = f"《{self.metadata['title']}》"
            p1.font.size = Pt(28)
            p1.font.bold = True
            p1.font.color.rgb = RGBColor(255, 255, 255)
            p1.alignment = PP_ALIGN_CENTER 
            
            p2 = tx.text_frame.add_paragraph()
            p2.text = f"{self.metadata['artist']}"
            p2.font.size = Pt(20)
            p2.font.color.rgb = RGBColor(220, 220, 220)
            p2.alignment = PP_ALIGN_CENTER 

        try:
            prs.save(self.output_ppt_path)
        except PermissionError:
            safe_print(f"      ❌ 保存失败！文件被占用: {self.output_ppt_path}")
            return

        # 清理当前线程生成的唯一临时文件
        for f in [self.temp_bg, self.temp_cover, self.temp_mask_top, self.temp_mask_bottom]:
            if os.path.exists(f): 
                try: os.remove(f)
                except: pass

# --- 单个文件处理函数 (供线程池调用) ---
def process_single_audio(filename, output_dir):
    try:
        file_base_name = os.path.splitext(filename)[0]
        output_path = os.path.join(output_dir, f"{file_base_name}.pptx")
        
        converter = AudioToPPT(filename, output_path)
        converter.extract_metadata()
        converter.generate_ppt()
        
        safe_print(f"✅ [完成] {filename}")
        return True
    except Exception as e:
        safe_print(f"❌ [失败] {filename} 错误: {e}")
        return False

# --- 批量程序 ---
def batch_process():
    output_dir = "output"
    if not os.path.exists(output_dir): os.makedirs(output_dir)
    
    audio_exts = ('.flac', '.mp3', '.wav', '.m4a')
    files = [f for f in os.listdir('.') if f.lower().endswith(audio_exts)]

    if not files:
        print("❌ 未找到音频文件。")
        return

    print(f"🔍 发现 {len(files)} 个文件，准备进行多线程处理...\n")

    # --- ⚡ 开启多线程处理 ---
    # max_workers=8 意味着同时处理4首歌。你可以根据电脑配置调整，不建议超过CPU核心数太多。
    start_time = time.time()
    
    with ThreadPoolExecutor(max_workers=4) as executor:
        # 提交所有任务
        futures = [executor.submit(process_single_audio, f, output_dir) for f in files]
        
        # 等待所有任务完成
        for future in as_completed(futures):
            # 这里可以处理返回值，目前我们主要依赖 print 输出状态
            pass

    end_time = time.time()
    print(f"\n🎉 全部处理完毕！耗时: {end_time - start_time:.2f} 秒")
    print(f"📂 输出目录: {output_dir}")

if __name__ == "__main__":
    batch_process()
