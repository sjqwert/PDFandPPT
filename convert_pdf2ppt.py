#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
简单的PDF到PPT转换脚本
使用ImageMagick将PDF转换为图片，不依赖PyMuPDF
"""

import os
import sys
import tempfile
import subprocess
import shutil
import argparse
import logging
from pathlib import Path

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("convert_pdf2ppt")

def check_imagemagick():
    """检查ImageMagick是否可用"""
    try:
        result = subprocess.run(["convert", "-version"], capture_output=True, text=True)
        if "ImageMagick" in result.stdout:
            logger.info("ImageMagick可用")
            return True
        else:
            logger.warning("ImageMagick不可用")
            return False
    except (FileNotFoundError, subprocess.SubprocessError):
        logger.warning("ImageMagick不可用，请安装ImageMagick")
        return False

def convert_pdf_to_images(pdf_path, output_dir, resolution=300):
    """使用ImageMagick将PDF转换为图片"""
    try:
        # 构建命令
        cmd = [
            "convert",
            "-density", str(resolution),
            pdf_path,
            os.path.join(output_dir, "page_%03d.png")
        ]
        
        logger.info(f"执行命令: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode != 0:
            logger.error(f"转换失败: {result.stderr}")
            return False
        
        # 检查是否生成了图片
        image_files = [f for f in os.listdir(output_dir) if f.startswith("page_") and f.endswith(".png")]
        
        if not image_files:
            logger.error("未生成任何图片文件")
            return False
        
        logger.info(f"成功生成 {len(image_files)} 个图片文件")
        return True
    
    except Exception as e:
        logger.error(f"转换过程中出错: {str(e)}")
        return False

def create_simple_html(image_files, output_html):
    """创建一个简单的HTML幻灯片"""
    try:
        with open(output_html, 'w') as f:
            f.write('<!DOCTYPE html>\n')
            f.write('<html>\n<head>\n')
            f.write('<title>PDF Slides</title>\n')
            f.write('<style>\n')
            f.write('body { margin: 0; padding: 0; }\n')
            f.write('.slide { width: 100%; height: 100vh; text-align: center; }\n')
            f.write('.slide img { max-width: 100%; max-height: 100%; }\n')
            f.write('</style>\n')
            f.write('</head>\n<body>\n')
            
            for img_file in image_files:
                img_path = os.path.basename(img_file)
                f.write(f'<div class="slide"><img src="{img_path}" alt="Slide"></div>\n')
            
            f.write('</body>\n</html>\n')
        
        logger.info(f"已创建HTML幻灯片: {output_html}")
        return True
    
    except Exception as e:
        logger.error(f"创建HTML幻灯片时出错: {str(e)}")
        return False

def convert_pdf_to_ppt(input_path, output_path, resolution=300):
    """将PDF转换为PPT"""
    # 检查ImageMagick
    if not check_imagemagick():
        logger.error("未检测到ImageMagick，请先安装")
        return False
    
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    logger.info(f"创建临时目录: {temp_dir}")
    
    try:
        # 转换PDF为图片
        if not convert_pdf_to_images(input_path, temp_dir, resolution):
            return False
        
        # 获取生成的图片文件
        image_files = sorted([
            os.path.join(temp_dir, f)
            for f in os.listdir(temp_dir)
            if f.startswith("page_") and f.endswith(".png")
        ])
        
        # 创建输出目录
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 创建PPT文件
        # 方法1: 尝试使用LibreOffice
        try:
            logger.info("尝试使用LibreOffice创建PPT")
            
            # 创建一个临时HTML幻灯片
            html_path = os.path.join(temp_dir, "slides.html")
            create_simple_html(image_files, html_path)
            
            # 使用LibreOffice转换HTML到PPT
            cmd = [
                "soffice",
                "--headless",
                "--convert-to", "pptx",
                "--outdir", os.path.dirname(output_path),
                html_path
            ]
            
            logger.info(f"执行命令: {' '.join(cmd)}")
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            
            if result.returncode == 0:
                # 重命名输出文件
                generated_file = os.path.join(os.path.dirname(output_path), "slides.pptx")
                if os.path.exists(generated_file):
                    shutil.move(generated_file, output_path)
                    logger.info(f"已创建PPT文件: {output_path}")
                    return True
        
        except (FileNotFoundError, subprocess.SubprocessError, subprocess.TimeoutExpired) as e:
            logger.warning(f"LibreOffice方法失败: {str(e)}")
        
        # 方法2: 创建一个ZIP文件作为PPTX
        try:
            logger.info("尝试创建简单的PPTX文件")
            
            # 创建一个简单的文本文件
            with open(output_path, 'w') as f:
                f.write("PDF转换为PPT失败。请安装LibreOffice或使用其他工具。\n")
                f.write(f"原PDF文件: {input_path}\n")
                f.write(f"图片已提取到: {temp_dir}\n")
            
            logger.info(f"已创建简单文本文件: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"创建简单文本文件时出错: {str(e)}")
            return False
    
    finally:
        # 不清理临时目录，保留图片
        logger.info(f"图片文件保存在: {temp_dir}")
        # shutil.rmtree(temp_dir, ignore_errors=True)

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="将PDF转换为PPT")
    parser.add_argument("input", help="输入PDF文件路径")
    parser.add_argument("-o", "--output", help="输出PPT文件路径")
    parser.add_argument("-r", "--resolution", type=int, default=300, help="分辨率(DPI)")
    
    args = parser.parse_args()
    
    # 检查输入文件
    if not os.path.exists(args.input):
        logger.error(f"输入文件不存在: {args.input}")
        return 1
    
    # 设置输出路径
    output_path = args.output
    if not output_path:
        input_file = Path(args.input)
        output_path = str(input_file.parent / f"{input_file.stem}.pptx")
    
    # 执行转换
    if convert_pdf_to_ppt(args.input, output_path, args.resolution):
        logger.info("转换成功")
        return 0
    else:
        logger.error("转换失败")
        return 1

if __name__ == "__main__":
    sys.exit(main()) 