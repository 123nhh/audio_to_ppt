import os
import win32com.client
import time
import shutil

def get_ppt_app():
    """启动 PowerPoint 应用程序"""
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        return app
    except Exception as e:
        print(f"[致命错误] 无法启动 PowerPoint: {e}")
        return None

def main():
    # --- 路径配置 ---
    ppt_source_dir = os.path.abspath("output")     # PPT 来源
    music_source_dir = os.path.abspath("music")    # 音频来源
    target_dir = os.path.abspath("ppt_output")     # 输出目录
    
    if not os.path.exists(ppt_source_dir):
        print(f"[错误] 找不到 PPT 来源文件夹: {ppt_source_dir}")
        return

    # 创建输出文件夹
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)

    # 1. 扫描 PPT 文件
    ppt_files = [f for f in os.listdir(ppt_source_dir) if f.endswith(".pptx") and not f.startswith("~")]
    if not ppt_files:
        print("[错误] output 文件夹里没有 PPT 文件。")
        return

    print(f"\n{'='*40}")
    print(f"      PPT 合并与音频精准匹配工具")
    print(f"{'='*40}")

    for i, f in enumerate(ppt_files):
        print(f"[{i+1}] {f}")
    
    print("\n请选择 PPT 合并顺序 (输入数字序号，用空格分隔)。")
    selection = input("\n请输入: ").strip()
    
    selected_ppts = []
    if not selection:
        selected_ppts = ppt_files
    else:
        try:
            for p in selection.split():
                idx = int(p) - 1
                if 0 <= idx < len(ppt_files):
                    selected_ppts.append(ppt_files[idx])
        except ValueError:
            print("[错误] 输入格式无效。")
            return

    if not selected_ppts:
        print("[退出] 未选择任何文件。")
        return

    # 2. 确定需要匹配的音频名称列表（不含后缀的文件名）
    # 例如：选中了 "演讲.pptx"，则记录 "演讲"
    selected_base_names = [os.path.splitext(f)[0] for f in selected_ppts]

    # 3. 开始合并 PPT
    print(f"\n[任务 1/2] 正在合并 PPT...")
    ppt_app = get_ppt_app()
    if not ppt_app: return

    base_pres = None
    output_ppt_path = os.path.join(target_dir, "合并后.pptx")
    
    if os.path.exists(output_ppt_path):
        try:
            os.remove(output_ppt_path)
        except:
            output_ppt_path = os.path.join(target_dir, f"合并后_{int(time.time())}.pptx")

    try:
        first_ppt = os.path.join(ppt_source_dir, selected_ppts[0])
        base_pres = ppt_app.Presentations.Open(first_ppt, WithWindow=True)
        base_pres.SaveAs(output_ppt_path)
        
        if len(selected_ppts) > 1:
            for filename in selected_ppts[1:]:
                print(f"  -> 正在追加 PPT: {filename}")
                file_path = os.path.join(ppt_source_dir, filename)
                base_pres.Slides.InsertFromFile(file_path, base_pres.Slides.Count)
        
        base_pres.Save()
        base_pres.Close()
        print(f"✅ PPT 已成功合并至: {output_ppt_path}")

    except Exception as e:
        print(f"❌ PPT 合并出错: {e}")
        if base_pres: base_pres.Close()

    # 4. 精准复制对应的音频文件
    print(f"\n[任务 2/2] 正在匹配并复制音频...")
    if not os.path.exists(music_source_dir):
        print(f"[跳过] 找不到 music 文件夹。")
    else:
        audio_exts = (".mp3", ".wav", ".m4a", ".wma", ".aac", ".flac")
        all_music_files = os.listdir(music_source_dir)
        
        found_audio_count = 0
        for audio_file in all_music_files:
            # 获取音频文件的名字（不含后缀）和 后缀
            name_part, ext_part = os.path.splitext(audio_file)
            
            # 如果该音频的名字在选中的 PPT 名字列表中，且后缀是音频格式
            if name_part in selected_base_names and ext_part.lower() in audio_exts:
                src = os.path.join(music_source_dir, audio_file)
                dst = os.path.join(target_dir, audio_file)
                try:
                    shutil.copy2(src, dst)
                    print(f"  -> [匹配成功] 已复制音频: {audio_file}")
                    found_audio_count += 1
                except Exception as e:
                    print(f"  [失败] 无法复制 {audio_file}: {e}")
        
        if found_audio_count == 0:
            print("  [提示] 未在 music 文件夹中找到与选中 PPT 同名的音频文件。")
        else:
            print(f"✅ 音频同步完成，共复制 {found_audio_count} 个文件。")

    print(f"\n{'='*40}")
    print(f"全部完成！输出目录: ./ppt_output")
    print(f"{'='*40}")

if __name__ == "__main__":
    main()