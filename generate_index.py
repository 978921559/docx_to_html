import os
from pathlib import Path


def generate_index_html(directory):
    # 获取当前目录展示名称
    dir_name = directory.name if directory != Path.cwd() else "Root Directory"

    # 收集有效子目录链接（仅包含已生成index.html的）
    subdir_links = []
    for item in directory.iterdir():
        if item.is_dir() and (item / "index.html").exists():
            rel_path = os.path.relpath(item / "index.html", directory)
            subdir_links.append((item.name, rel_path, 'dir'))

    # 收集同级HTML文件（排除自身）
    file_links = []
    for item in directory.iterdir():
        if item.is_file() and item.suffix == ".html" and item.name != "index.html":
            file_links.append((item.stem, item.name, 'file'))

    # 合并并排序链接（目录在前，文件在后）
    all_links = sorted(subdir_links, key=lambda x: x[0].lower()) + \
                sorted(file_links, key=lambda x: x[0].lower())

    # 生成HTML内容
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{dir_name} - 索引</title>
    <style>
        :root {{
            --bg-color: #ffffff;
            --text-color: #2d3748;
            --accent-color: #4a5568;
            --border-color: #e2e8f0;
            --dir-color: #2b6cb0;
            --file-color: #718096;
        }}

        @media (prefers-color-scheme: dark) {{
            :root {{
                --bg-color: #1a202c;
                --text-color: #e2e8f0;
                --accent-color: #a0aec0;
                --border-color: #4a5568;
                --dir-color: #63b3ed;
                --file-color: #a0aec0;
            }}
        }}

        body {{
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            line-height: 1.6;
            margin: 2rem auto;
            max-width: 800px;
            padding: 0 1rem;
            color: var(--text-color);
            background-color: var(--bg-color);
        }}

        .header {{
            padding-bottom: 1.5rem;
            margin-bottom: 2rem;
            border-bottom: 2px solid var(--border-color);
        }}

        .title {{
            font-size: 1.875rem;
            margin: 0 0 0.5rem;
            color: var(--accent-color);
        }}

        .count {{
            color: var(--file-color);
            font-size: 0.875rem;
        }}

        .link-list {{
            list-style: none;
            padding: 0;
            margin: 0;
        }}

        .link-item {{
            padding: 0.75rem;
            margin: 0.5rem 0;
            border-radius: 0.375rem;
            transition: all 0.2s ease;
            background: var(--bg-color);
            border: 1px solid var(--border-color);
        }}

        .link-item:hover {{
            transform: translateX(4px);
            border-color: var(--dir-color);
        }}

        .link-item a {{
            text-decoration: none;
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }}

        .link-item[data-type="dir"] {{
            border-left: 4px solid var(--dir-color);
        }}

        .link-item[data-type="dir"] a::before {{
            content: "📁";
            font-size: 1.2em;
            color: var(--dir-color);
        }}

        .link-item[data-type="file"] {{
            border-left: 4px solid var(--file-color);
        }}

        .link-item[data-type="file"] a::before {{
            content: "📄";
            font-size: 1.2em;
            color: var(--file-color);
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1 class="title">{dir_name}</h1>
        <p class="count">共 {len(all_links)} 个项目</p>
    </div>

    <ul class="link-list">
        {"".join(
        f'<li class="link-item" data-type="{link_type}">'
        f'<a href="{path}">{name}</a>'
        f'</li>'
        for name, path, link_type in all_links
    )}
    </ul>
</body>
</html>"""

    # 写入文件
    (directory / "index.html").write_text(html_content, encoding="utf-8")


def main():
    # 获取所有需要处理的目录（按深度倒序处理）
    directories = []
    for root, _, _ in os.walk(os.getcwd()):
        directories.append(Path(root))

    # 按目录深度排序（从最深开始处理）
    directories.sort(key=lambda p: len(p.parts), reverse=True)

    for directory in directories:
        # 检查是否需要生成索引
        has_content = any(
            (item.is_dir() or
             (item.is_file() and item.suffix == ".html" and item.name != "index.html"))
            for item in directory.iterdir()
        )

        if has_content or directory == Path.cwd():
            generate_index_html(directory)
            print(f"生成目录索引：{directory}")


if __name__ == "__main__":
    main()