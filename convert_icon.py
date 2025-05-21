from PIL import Image
import os
import sys

try:
    # 输入和输出文件路径
    input_image = r"d:\007\vscode\cursor\Image.png"
    output_icon = r"d:\007\vscode\cursor\new_electrochemistry_icon.ico"
    
    print(f"正在转换图像: {input_image}")
    print(f"目标图标文件: {output_icon}")
    
    # 检查输入文件是否存在
    if not os.path.exists(input_image):
        print(f"错误: 输入文件不存在: {input_image}")
        sys.exit(1)
    
    # 打开图像
    img = Image.open(input_image)
    print(f"已打开图像，大小: {img.size}, 格式: {img.format}")
    
    # 确保图像是正方形，为了更好的图标效果
    width, height = img.size
    size = max(width, height)
    new_img = Image.new("RGBA", (size, size), (255, 255, 255, 0))
    new_img.paste(img, ((size - width) // 2, (size - height) // 2))
    img = new_img
    
    # 调整为合适的图标尺寸（通常是 256x256 或更小的多种尺寸）
    icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    print(f"正在创建多尺寸图标: {', '.join([f'{s[0]}x{s[1]}' for s in icon_sizes])}")
    
    # 保存为 ico 文件
    img.save(output_icon, format='ICO', sizes=icon_sizes)
    
    # 验证文件是否已创建
    if os.path.exists(output_icon):
        print(f"图标已成功保存至 {output_icon}")
        print(f"文件大小: {os.path.getsize(output_icon)} 字节")
    else:
        print(f"错误: 图标文件未能创建: {output_icon}")

except Exception as e:
    print(f"发生错误: {str(e)}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
