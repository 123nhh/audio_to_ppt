import os
import re
import time
import uuid
import json
import glob
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from PIL import Image, ImageFilter, ImageEnhance
from mutagen import File
from mutagen.flac import FLAC
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE

# --- 新增: OpenAI 库 ---
try:
    from openai import OpenAI
except ImportError:
    print("[错误] 缺少 openai 库，无法清洗歌词。请运行: pip install openai")
    OpenAI = None

# ==========================================
# 配置初始化逻辑
# ==========================================
CONFIG_FILE = "ai_config.json"
DEFAULT_KEY_PLACEHOLDER = "sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"

def init_ai_configuration():
    """
    初始化 AI 配置：
    1. 尝试读取配置文件
    2. 如果不存在或配置无效，则在命令行引导用户输入
    3. 保存配置到文件
    """
    default_config = {
        "enabled": True,
        "api_key": DEFAULT_KEY_PLACEHOLDER,
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-3.5-turbo",
        "max_retries": 3
    }

    # 1. 尝试读取现有配置
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved_config = json.load(f)
                for k, v in default_config.items():
                    if k not in saved_config:
                        saved_config[k] = v
                
                if saved_config["enabled"] and (not saved_config["api_key"] or saved_config["api_key"] == DEFAULT_KEY_PLACEHOLDER):
                    print("[提示] 配置文件存在，但 API Key 无效。")
                else:
                    return saved_config
        except Exception as e:
            print(f"[警告] 读取配置文件失败: {e}，将重新配置。")

    # 2. 进入交互式配置向导
    print("\n" + "="*50)
    print("       AI 歌词清洗功能配置向导")
    print("="*50)
    print("检测到这是你第一次运行，或者 AI 配置尚未完成。")
    print("为了去除歌词中的废话（作词、作曲、推广等），我们需要配置 AI 接口。\n")

    use_ai = input("是否开启 AI 歌词清洗功能? (y/n) [默认: y]: ").strip().lower()
    if use_ai == 'n':
        default_config["enabled"] = False
        print("[设置] 已关闭 AI 功能。")
    else:
        default_config["enabled"] = True
        while True:
            user_key = input("\n请输入你的 API Key (必填): ").strip()
            if user_key and user_key != DEFAULT_KEY_PLACEHOLDER:
                default_config["api_key"] = user_key
                break
            print("[错误] API Key 不能为空，请重新输入。")

        print(f"\n请输入接口地址 (Base URL)")
        print(f"如果你使用 OpenAI 官方，直接回车即可。")
        print(f"如果你使用 DeepSeek/Kimi/OneApi 等中转，请输入对应的 v1 地址。")
        user_url = input(f"地址 [默认: {default_config['base_url']}]: ").strip()
        if user_url:
            default_config["base_url"] = user_url

        user_model = input(f"\n请输入模型名称 [默认: {default_config['model']}]: ").strip()
        if user_model:
            default_config["model"] = user_model

    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(default_config, f, indent=4, ensure_ascii=False)
        print(f"\n[成功] 配置已保存至 {CONFIG_FILE}")
        print("="*50 + "\n")
    except Exception as e:
        print(f"[错误] 无法写入配置文件: {e}")

    return default_config

AI_CONFIG = init_ai_configuration()

# ==========================================
# 核心逻辑
# ==========================================

print_lock = threading.Lock()
def safe_print(msg):
    with print_lock:
        print(msg)

def call_ai_to_clean_lyrics(raw_text, log_tag):
    if not AI_CONFIG["enabled"] or not OpenAI:
        return raw_text
    
    if len(raw_text) < 10:
        return raw_text

    client = OpenAI(api_key=AI_CONFIG["api_key"], base_url=AI_CONFIG["base_url"])
    
    system_prompt = "你是一个歌词处理程序。"
    user_prompt = (
        "请严格执行以下操作：\n"
        "1. 如果歌词包含'纯音乐'、'Instrumental'或没有实际歌词内容，请仅回复: [PURE_MUSIC]\n"
        "2. 删除头部元数据（作词、作曲、编曲等）。\n"
        "3. 删除尾部宣传信息（统筹、出品、邮箱等）。\n"
        "4. 必须保留原有的换行格式。\n"
        "5. 不要输出任何解释，只输出结果。\n\n"
        "待处理文本：\n"
        f"{raw_text}"
    )

    for attempt in range(AI_CONFIG["max_retries"]):
        try:
            response = client.chat.completions.create(
                model=AI_CONFIG["model"],
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.1,
                timeout=20
            )
            cleaned_content = response.choices[0].message.content.strip()
            safe_print(f"[{log_tag}] [AI] 清洗成功 (尝试 {attempt+1}/{AI_CONFIG['max_retries']})")
            return cleaned_content

        except Exception as e:
            safe_print(f"[{log_tag}] [警告] AI调用失败 (尝试 {attempt+1}/{AI_CONFIG['max_retries']}): {e}")
            time.sleep(1)

    safe_print(f"[{log_tag}] [错误] AI多次重试失败，将使用原始歌词")
    return raw_text

class AudioToPPT:
    def __init__(self, audio_path, output_ppt_path):
        self.audio_path = os.path.abspath(audio_path)
        self.output_ppt_path = os.path.abspath(output_ppt_path)
        self.file_name = os.path.basename(audio_path)
        self.is_pure_music = False
        
        self.uid = uuid.uuid4().hex[:8] 
        self.temp_bg = f"temp_bg_{self.uid}.jpg"
        self.temp_cover = f"temp_cover_{self.uid}.jpg"
        self.temp_mask_top = f"temp_mask_top_{self.uid}.jpg"
        self.temp_mask_bottom = f"temp_mask_bottom_{self.uid}.jpg"
        
        self.metadata = {
            'title': '未知标题', 'artist': '未知歌手', 'lyrics': [], 'cover_data': None
        }
        
        self.SLIDE_W = Inches(13.333) 
        self.SLIDE_H = Inches(7.5)
        self.CENTER_Y = self.SLIDE_H / 2
        self.SCROLL_UNIT_HEIGHT = Inches(0.9) 
        self.TEXTBOX_W = Inches(8.0)
        self.TEXTBOX_X = Inches(4.5) 
        self.TEXTBOX_H = Inches(100)
        
        self.STYLE_ACTIVE = {'size': 40, 'bold': True, 'color': (255, 255, 255)}
        self.STYLE_NORMAL = {'size': 26, 'bold': False, 'color': (150, 150, 150)}

    def _log(self, msg):
        tag = self.metadata.get('title', '')
        if not tag or tag == '未知标题':
            tag = self.file_name
        safe_print(f"[{tag}] {msg}")

    def parse_lyrics_lines(self, text_content):
        cleaned_lines = []
        if not text_content: return cleaned_lines
        lines = text_content.split('\n')
        pattern = re.compile(r'\[\d{1,3}:\d{2}(?:\.\d{1,3})?\]')
        for line in lines:
            line_content = re.sub(pattern, '', line).strip()
            if line_content: 
                cleaned_lines.append(line_content)
        return cleaned_lines

    def extract_metadata(self):
        try:
            audio = File(self.audio_path)
            tags = audio.tags
            if tags:
                raw_title = str(tags.get('TITLE', tags.get('TIT2', ['未知标题']))[0])
                self.metadata['title'] = raw_title.replace("《", "").replace("》", "").strip()
                self.metadata['artist'] = str(tags.get('ARTIST', tags.get('TPE1', ['未知歌手']))[0])
                
                raw_lyrics_text = ""
                if isinstance(audio, FLAC):
                    raw_lyrics_text = tags.get('lyrics', tags.get('unsyncedlyrics', ['']))[0]
                elif tags and hasattr(tags, 'getall'): 
                     uslt = tags.getall('USLT')
                     if uslt: raw_lyrics_text = uslt[0].text

                if raw_lyrics_text:
                    self._log(f"[处理] 正在分析歌词...")
                    
                    if "纯音乐" in raw_lyrics_text or "Instrumental" in raw_lyrics_text:
                        self.is_pure_music = True
                        self._log("[识别] 检测到纯音乐标记 (元数据)")
                    else:
                        display_name = self.metadata['title'] if self.metadata['title'] != '未知标题' else self.file_name
                        ai_result = call_ai_to_clean_lyrics(raw_lyrics_text, display_name)
                        
                        if "[PURE_MUSIC]" in ai_result:
                            self.is_pure_music = True
                            self._log("[识别] AI 判定为纯音乐")
                        else:
                            final_lines = self.parse_lyrics_lines(ai_result)
                            self.metadata['lyrics'] = final_lines
                            if not final_lines:
                                self.is_pure_music = True
                                self._log("[识别] 清洗后无有效歌词，视为纯音乐")
                else:
                    self.is_pure_music = True
                
                if isinstance(audio, FLAC) and audio.pictures:
                    self.metadata['cover_data'] = audio.pictures[0].data
                elif hasattr(audio, 'tags') and 'APIC:' in audio.tags: 
                     for key in audio.tags.keys():
                         if key.startswith('APIC'):
                             self.metadata['cover_data'] = audio.tags[key].data
                             break
        except Exception as e:
            self._log(f"[警告] 元数据读取可能有误: {e}")

    def prepare_images(self):
        if not self.metadata['cover_data']: return None
        try:
            img = Image.open(BytesIO(self.metadata['cover_data'])).convert("RGB")
            bg_img = img.filter(ImageFilter.GaussianBlur(radius=60))
            bg_img = ImageEnhance.Brightness(bg_img).enhance(0.3) 
            target_w, target_h = 1280, 720
            bg_img = bg_img.resize((target_w, target_h))
            bg_img.save(self.temp_bg)
            img.save(self.temp_cover)
            mask_h = 200
            mask_top = bg_img.crop((0, 0, target_w, mask_h))
            mask_top.save(self.temp_mask_top)
            mask_bottom = bg_img.crop((0, target_h - mask_h, target_w, target_h))
            mask_bottom.save(self.temp_mask_bottom)
            return True
        except Exception as e:
            self._log(f"[跳过] 图片处理失败: {e}")
            return False

    def generate_ppt(self):
        if os.path.exists(self.output_ppt_path):
            try:
                os.remove(self.output_ppt_path)
                self._log(f"[清理] 已删除旧文件: {os.path.basename(self.output_ppt_path)}")
            except PermissionError:
                self._log(f"[错误] 无法删除旧文件，请先关闭PPT: {self.output_ppt_path}")
                return False

        prs = Presentation()
        prs.slide_width = self.SLIDE_W
        prs.slide_height = self.SLIDE_H

        if not self.prepare_images():
            self._log("[跳过] 无法生成必要图片资源。")
            return False

        # Slide 1: 封面
        slide_intro = prs.slides.add_slide(prs.slide_layouts[6])
        slide_intro.shapes.add_picture(self.temp_bg, 0, 0, width=self.SLIDE_W, height=self.SLIDE_H)
        try:
            slide_intro.shapes.add_picture(self.temp_mask_top, 0, 0, width=self.SLIDE_W, height=Inches(2.0))
            slide_intro.shapes.add_picture(self.temp_mask_bottom, 0, self.SLIDE_H - Inches(2.0), width=self.SLIDE_W, height=Inches(2.0))
        except: pass
        intro_cover_size = Inches(5.0)
        intro_cover_left = (self.SLIDE_W - intro_cover_size) / 2
        slide_intro.shapes.add_picture(self.temp_cover, intro_cover_left, Inches(0.8), width=intro_cover_size, height=intro_cover_size)
        tx_intro = slide_intro.shapes.add_textbox(0, Inches(6.0), self.SLIDE_W, Inches(1.5))
        p_title = tx_intro.text_frame.add_paragraph()
        p_title.text = self.metadata['title'] 
        p_title.font.size = Pt(36)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(255, 255, 255)
        p_title.alignment = PP_ALIGN.CENTER
        p_artist = tx_intro.text_frame.add_paragraph()
        p_artist.text = self.metadata['artist']
        p_artist.font.size = Pt(24)
        p_artist.font.color.rgb = RGBColor(180, 180, 180)
        p_artist.alignment = PP_ALIGN.CENTER

        # 纯音乐逻辑
        if self.is_pure_music or not self.metadata['lyrics']:
            self._log(f"[完成] 纯音乐模式：仅生成封面，跳过歌词页。")
            try:
                prs.save(self.output_ppt_path)
                self._clean_temp_files()
                return True
            except Exception as e:
                self._log(f"[错误] 保存失败: {e}")
                return False

        # Slide 2+: 歌词
        lyrics = self.metadata['lyrics']
        for current_idx in range(len(lyrics)):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(self.temp_bg, 0, 0, width=self.SLIDE_W, height=self.SLIDE_H)
            dynamic_top = self.CENTER_Y - (current_idx * self.SCROLL_UNIT_HEIGHT) - Inches(0.5)
            tb = slide.shapes.add_textbox(self.TEXTBOX_X, dynamic_top, self.TEXTBOX_W, self.TEXTBOX_H)
            tf = tb.text_frame
            tf.word_wrap = True 
            tf.auto_size = MSO_AUTO_SIZE.NONE
            tf.clear()
            for line_idx, line_text in enumerate(lyrics):
                p = tf.add_paragraph()
                p.text = line_text
                if line_idx == current_idx:
                    p.font.size = Pt(self.STYLE_ACTIVE['size'])
                    p.font.bold = self.STYLE_ACTIVE['bold']
                    p.font.color.rgb = RGBColor(*self.STYLE_ACTIVE['color'])
                else:
                    p.font.size = Pt(self.STYLE_NORMAL['size'])
                    p.font.bold = self.STYLE_NORMAL['bold']
                    p.font.color.rgb = RGBColor(*self.STYLE_NORMAL['color'])
                p.alignment = PP_ALIGN.CENTER
                p.space_before = Pt(0)
                p.space_after = Pt(30) 
            try:
                slide.shapes.add_picture(self.temp_mask_top, 0, 0, width=self.SLIDE_W, height=Inches(2.0))
                slide.shapes.add_picture(self.temp_mask_bottom, 0, self.SLIDE_H - Inches(2.0), width=self.SLIDE_W, height=Inches(2.0))
            except: pass
            small_cover_size = Inches(3.2)
            slide.shapes.add_picture(self.temp_cover, Inches(0.8), Inches(2.0), width=small_cover_size, height=small_cover_size)
            info_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.0) + small_cover_size + Inches(0.1), small_cover_size, Inches(1.5))
            p_song = info_box.text_frame.add_paragraph()
            p_song.text = self.metadata['title']
            p_song.font.size = Pt(20)
            p_song.font.bold = True
            p_song.font.color.rgb = RGBColor(255, 255, 255)
            p_song.alignment = PP_ALIGN.CENTER
            p_art = info_box.text_frame.add_paragraph()
            p_art.text = self.metadata['artist']
            p_art.font.size = Pt(14)
            p_art.font.color.rgb = RGBColor(180, 180, 180)
            p_art.alignment = PP_ALIGN.CENTER

        try:
            prs.save(self.output_ppt_path)
        except PermissionError:
            self._log(f"[错误] 保存失败！文件被占用: {self.output_ppt_path}")
            return False
        
        self._clean_temp_files()
        return True

    def _clean_temp_files(self):
        for f in [self.temp_bg, self.temp_cover, self.temp_mask_top, self.temp_mask_bottom]:
            if os.path.exists(f): 
                try: os.remove(f)
                except: pass

def process_single_audio(filename, output_dir):
    try:
        file_base_name = os.path.splitext(filename)[0]
        relative_output_path = os.path.join(output_dir, f"{file_base_name}.pptx")
        abs_output_path = os.path.abspath(relative_output_path)
        
        converter = AudioToPPT(filename, abs_output_path)
        converter.extract_metadata()
        success = converter.generate_ppt()
        
        final_name = converter.metadata.get('title', file_base_name)
        if final_name == '未知标题': final_name = file_base_name

        if success:
            safe_print(f"[{final_name}] [完成] PPT已生成")
            return True
        else:
            return False
    except Exception as e:
        safe_print(f"[{filename}] [失败] 错误: {e}")
        return False

def cleanup_residual_files():
    """
    清理逻辑：扫描目录中所有符合脚本生成的临时文件格式，并删除。
    防止程序崩溃后垃圾文件残留。
    """
    # 匹配模式：temp_xxx_8位hex.jpg
    # 例如: temp_bg_a1b2c3d4.jpg
    patterns = [
        "temp_bg_*.jpg", 
        "temp_cover_*.jpg", 
        "temp_mask_top_*.jpg", 
        "temp_mask_bottom_*.jpg"
    ]
    
    deleted_count = 0
    for pattern in patterns:
        for filepath in glob.glob(pattern):
            # 简单校验文件名，防止误删用户文件 (检查是否包含 'temp_' 和 '.jpg')
            if "temp_" in filepath and filepath.endswith(".jpg"):
                try:
                    os.remove(filepath)
                    deleted_count += 1
                except Exception:
                    pass
    
    if deleted_count > 0:
        print(f"\n[清理] 已自动清理 {deleted_count} 个残留的临时文件。")

def batch_process():
    output_dir = "output"
    if not os.path.exists(output_dir): os.makedirs(output_dir)
    audio_exts = ('.flac', '.mp3', '.wav', '.m4a')
    files = [f for f in os.listdir('.') if f.lower().endswith(audio_exts)]
    if not files:
        print("[错误] 未找到音频文件。")
        cleanup_residual_files() # 即使没文件也检查一下残留
        return
    
    print(f"[扫描] 发现 {len(files)} 个文件 | 模式：交互配置 + 纯音乐过滤 + 自动清理残留\n")
    
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(process_single_audio, f, output_dir) for f in files]
        for future in as_completed(futures): pass
    
    # --- 最终清理 ---
    cleanup_residual_files()
    
    print(f"\n[结束] 全部处理完毕！请查看 {output_dir} 文件夹。")

if __name__ == "__main__":
    batch_process()
