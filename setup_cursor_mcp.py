#!/usr/bin/env python3
"""
Cursor MCP配置设置脚本
自动生成Cursor MCP配置文件
"""

import json
import os
import subprocess
from pathlib import Path

def create_cursor_config():
    """创建Cursor MCP配置文件"""
    
    # 获取当前目录的绝对路径
    current_dir = Path(__file__).parent.absolute()
    user_mcp_path = Path.home() / ".cursor" / "mcp.json"

    # 尝试读取现有的 Cursor MCP 配置（用于复用 Context7 等第三方服务配置）
    existing_context7 = None
    try:
        if user_mcp_path.exists():
            with open(user_mcp_path, "r", encoding="utf-8") as f:
                existing = json.load(f)
            ctx = existing.get("mcpServers", {}).get("Context7")
            if ctx:
                existing_context7 = ctx
    except Exception:
        # 读取失败时忽略，不影响下方配置生成
        existing_context7 = None
    
    # 依据当前项目路径，生成与现有 c:\\Users\\<User>\\.cursor\\mcp.json 对齐的配置
    # 参考：excel-mcp 使用模块形式 + cwd；powerpoint-mcp 与 word-document-server 使用脚本路径
    config_servers = {}

    # 复用 Context7（如用户现有配置中存在）
    if existing_context7 is not None:
        config_servers["Context7"] = existing_context7

    # Office 系列服务（统一默认读写目录为 workspace）
    workspace_dir = str(current_dir / "workspace")
    config_servers["excel-mcp"] = {
        "command": "python",
        "args": ["-m", "excel_mcp", "stdio"],
        "cwd": str(current_dir / "excel-mcp-server-main"),
        "env": {
            "EXCEL_FILES_PATH": workspace_dir
        }
    }

    config_servers["powerpoint-mcp"] = {
        "command": "python",
        "args": [str(current_dir / "Office-PowerPoint-MCP-Server-main" / "ppt_mcp_server.py")],
        "env": {
            "PPT_TEMPLATE_PATH": workspace_dir
        }
    }

    config_servers["word-document-server"] = {
        "command": "python",
        "args": [str(current_dir / "Office-Word-MCP-Server-main" / "word_mcp_server.py")],
        "env": {
            "MCP_TRANSPORT": "stdio",
            "WORD_FILES_PATH": workspace_dir
        }
    }

    # 汇总配置
    config = {"mcpServers": config_servers}
    
    # 保存配置文件
    config_file = current_dir / "cursor_mcp_config.json"
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
    
    print("=" * 60)
    print("🎉 Cursor MCP配置已生成!")
    print("=" * 60)
    print(f"配置文件位置: {config_file}")
    print()
    print("请将以下内容添加到你的Cursor MCP配置中:")
    print("(通常在 ~/.cursor/mcp_settings.json 或类似位置)")
    print()
    print("-" * 60)
    print(json.dumps(config, indent=2, ensure_ascii=False))
    print("-" * 60)
    print()
    print("配置说明:")
    print("1. excel-mcp: Excel文件操作服务")
    print("2. powerpoint-mcp: PowerPoint演示文稿操作服务") 
    print("3. word-mcp: Word文档操作服务")
    print()
    print("使用说明:")
    print("1. 确保已安装所有依赖: pip install -r requirements.txt")
    print("2. 将上述配置添加到Cursor的MCP设置中")
    print("3. 重启Cursor")
    print("4. 在Cursor中就可以使用Office文件操作功能了!")
    print()
    print("文件目录:")
    print(f" {current_dir / 'workspace'}")

def create_directories():
    """创建必要的目录"""
    current_dir = Path(__file__).parent
    
    directories = [
        current_dir / 'workspace'
    ]
    
    print("创建必要的目录...")
    for directory in directories:
        directory.mkdir(exist_ok=True)
        print(f"✓ {directory}")

def install_dependencies():
    """安装项目依赖: pip install -r requirements.txt"""
    current_dir = Path(__file__).parent
    req_file = current_dir / 'requirements.txt'
    if not req_file.exists():
        print("未找到 requirements.txt，跳过依赖安装")
        return True
    print("安装依赖: pip install -r requirements.txt")
    try:
        # 使用与当前 Python 一致的解释器来安装依赖，直接显示输出到终端
        result = subprocess.run(
            [os.sys.executable, '-m', 'pip', 'install', '-r', str(req_file)],
            check=False
        )
        if result.returncode != 0:
            print("❌ 依赖安装失败")
            return False
        print("✓ 依赖安装完成")
        return True
    except Exception as e:
        print(f"❌ 依赖安装异常: {e}")
        return False

def check_dependencies():
    """检查依赖"""
    print("检查依赖...")
    
    required_packages = [
        ('mcp', 'mcp'),
        ('fastmcp', 'fastmcp'), 
        ('openpyxl', 'openpyxl'),
        ('python-pptx', 'pptx'),
        ('python-docx', 'docx'),
        ('Pillow', 'PIL'),
        ('msoffcrypto-tool', 'msoffcrypto'),
        ('docx2pdf', 'docx2pdf'),
        ('typer', 'typer')
    ]
    
    missing_packages = []
    for package_name, import_name in required_packages:
        try:
            __import__(import_name)
        except ImportError:
            missing_packages.append(package_name)
    
    if missing_packages:
        print(f"❌ 缺少依赖包: {', '.join(missing_packages)}")
        print("尝试自动安装缺少的依赖...")
        if not install_dependencies():
            print("请手动运行: pip install -r requirements.txt")
            return False
        # 重新校验
        return check_dependencies()
    else:
        print("✓ 所有依赖已安装")
        return True

def main():
    """主函数"""
    print("Office MCP Services - Cursor配置设置")
    print("=" * 60)
    
    # 创建目录
    create_directories()
    print()
    
    # 安装/检查依赖
    install_dependencies()
    if not check_dependencies():
        print("\n依赖未完全满足，请根据提示安装后重试")
        return
    print()
    
    # 创建配置
    create_cursor_config()

if __name__ == "__main__":
    main()
