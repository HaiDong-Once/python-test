import os
import re
import base64
import requests
import json
import time
import logging
from git import Repo
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

# 设置日志
logger = logging.getLogger(__name__)

# Gitee API相关配置
GITEE_API_URL = "https://gitee.com/api/v5"
GITEE_REPO_OWNER = os.getenv("GITEE_REPO_OWNER", "comma-dong")
GITEE_REPO_NAME = os.getenv("GITEE_REPO_NAME", "image-projects")
GITEE_ACCESS_TOKEN = os.getenv("GITEE_ACCESS_TOKEN", "")
GITEE_BRANCH = "master"

def create_gitee_repo_if_not_exists():
    """如果仓库不存在，则创建仓库"""
    check_url = f"{GITEE_API_URL}/repos/{GITEE_REPO_OWNER}/{GITEE_REPO_NAME}"
    headers = {"Content-Type": "application/json;charset=UTF-8"}
    
    if GITEE_ACCESS_TOKEN:
        params = {"access_token": GITEE_ACCESS_TOKEN}
        response = requests.get(check_url, params=params, headers=headers)
        
        if response.status_code == 404:
            # 仓库不存在，创建仓库
            create_url = f"{GITEE_API_URL}/user/repos"
            data = {
                "access_token": GITEE_ACCESS_TOKEN,
                "name": GITEE_REPO_NAME,
                "description": "Automatically created repository for storing images",
                "private": False,
                "has_issues": False,
                "has_wiki": False
            }
            create_response = requests.post(create_url, json=data, headers=headers)
            if create_response.status_code == 201:
                print(f"成功创建仓库 {GITEE_REPO_NAME}")
            else:
                print(f"创建仓库失败: {create_response.text}")
                return False
    
    return True

def upload_image_to_gitee(image_path):
    """将图片上传到Gitee仓库并返回图片URL"""
    if not GITEE_ACCESS_TOKEN:
        raise ValueError("请设置Gitee访问令牌")
    
    # 确保仓库存在
    if not create_gitee_repo_if_not_exists():
        raise ValueError("无法访问或创建Gitee仓库")
    
    # 准备图片数据
    with open(image_path, 'rb') as f:
        image_data = f.read()
    
    # 对图片内容进行Base64编码
    base64_data = base64.b64encode(image_data).decode('utf-8')
    
    # 准备文件路径和提交信息
    filename = os.path.basename(image_path)
    timestamp = time.strftime("%Y%m%d%H%M%S")
    file_path = f"images/{timestamp}_{filename}"
    commit_message = f"Upload image {filename}"
    
    # 上传文件到Gitee
    url = f"{GITEE_API_URL}/repos/{GITEE_REPO_OWNER}/{GITEE_REPO_NAME}/contents/{file_path}"
    params = {
        "access_token": GITEE_ACCESS_TOKEN,
        "content": base64_data,
        "message": commit_message,
        "branch": GITEE_BRANCH
    }
    headers = {"Content-Type": "application/json;charset=UTF-8"}
    
    response = requests.post(url, json=params, headers=headers)
    
    if response.status_code == 201:
        result = response.json()
        return result.get("content").get("download_url")
    else:
        print(f"上传图片 {filename} 失败: {response.text}")
        return None

def update_image_links_in_md(md_file_path, local_image_dir, remote_image_urls):
    """更新Markdown文件中的图片链接"""
    logger.info(f"开始更新Markdown文件中的图片链接: {md_file_path}")
    
    try:
        with open(md_file_path, 'r', encoding='utf-8') as f:
            md_content = f.read()
        
        updated_content = md_content
        update_count = 0
        
        # 遍历所有本地图片路径和对应的远程URL
        for local_filename, remote_url in remote_image_urls.items():
            if not remote_url:
                continue
                
            # 查找图片引用模式
            # 匹配 ![任意文本](文件名.扩展名) 格式的图片引用
            pattern = re.compile(r'(!\[[^\]]*?\])(\(' + re.escape(local_filename) + r'\))', re.IGNORECASE)
            
            if pattern.search(updated_content):
                # 替换为远程URL
                updated_content = pattern.sub(r'\1(' + remote_url + ')', updated_content)
                update_count += 1
                logger.info(f"替换图片链接: {local_filename} -> {remote_url}")
            else:
                logger.warning(f"未找到图片引用: {local_filename}")
        
        # 如果内容有更新，写回文件
        if updated_content != md_content:
            with open(md_file_path, 'w', encoding='utf-8') as f:
                f.write(updated_content)
            logger.info(f"完成 {update_count} 个图片链接更新")
        else:
            logger.warning("未更新任何图片链接")
            
    except Exception as e:
        logger.error(f"更新图片链接时出错: {str(e)}", exc_info=True)

def upload_images_to_gitee(md_file_path, local_image_dir):
    """上传本地图片到Gitee并更新Markdown文件中的链接"""
    logger.info(f"开始处理图片上传: {md_file_path}, 图片目录: {local_image_dir}")
    
    # 确保token存在
    if not GITEE_ACCESS_TOKEN:
        logger.warning("Gitee访问令牌未设置，跳过图片上传")
        return {}
    
    try:
        # 查找Markdown文件中的所有本地图片引用
        with open(md_file_path, 'r', encoding='utf-8') as f:
            md_content = f.read()
        
        # 使用正则表达式查找图片引用
        # 匹配 ![任意文本](文件名.扩展名) 格式的图片引用
        img_pattern = r'!\[(.*?)\]\(([^/\)]+\.(png|jpg|jpeg|gif|svg|webp))\)'
        local_images = re.findall(img_pattern, md_content, re.IGNORECASE)
        
        if not local_images:
            logger.info("未在Markdown中找到图片引用，跳过上传")
            return {}
        
        logger.info(f"在Markdown中找到 {len(local_images)} 个图片引用")
        
        # 上传图片到Gitee并获取远程URL
        remote_image_urls = {}
        for alt_text, image_filename, ext in local_images:
            # 获取图片的绝对路径，图片和markdown在同一目录
            image_abs_path = os.path.join(local_image_dir, image_filename)
            
            # 检查文件是否存在
            if not os.path.exists(image_abs_path):
                logger.warning(f"图片文件不存在: {image_abs_path}")
                continue
            
            logger.info(f"开始上传图片: {image_filename}")
            try:
                # 上传图片到Gitee
                remote_url = upload_image_to_gitee(image_abs_path)
                if remote_url:
                    logger.info(f"图片上传成功: {remote_url}")
                    remote_image_urls[image_filename] = remote_url
                else:
                    logger.warning(f"图片上传失败: {image_filename}")
            except Exception as e:
                logger.error(f"上传图片 {image_filename} 时出错: {str(e)}", exc_info=True)
        
        # 更新Markdown文件中的图片链接
        if remote_image_urls:
            logger.info(f"更新Markdown文件中的图片链接: {md_file_path}")
            update_image_links_in_md(md_file_path, local_image_dir, remote_image_urls)
            logger.info("图片链接更新完成")
        
        return remote_image_urls
    
    except Exception as e:
        logger.error(f"处理图片上传时出错: {str(e)}", exc_info=True)
        return {} 