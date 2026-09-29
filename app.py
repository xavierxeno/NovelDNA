import streamlit as st
import chromadb
import ollama
from openai import OpenAI
import os, random, shutil
from pathlib import Path
import re

st.set_page_config(page_title="作品语料检索与风格解析器", layout="wide")
# ========== 全局连接 ==========
client = chromadb.PersistentClient(path="./chroma_db")
EMBED_MODEL = "bge-m3"
# ========== 基础工具函数 ==========
def list_authors():
    raw = Path("./raw")
    if not raw.exists():
        raw.mkdir()
    return sorted([p.name for p in raw.iterdir() if p.is_dir()])

def list_dna_authors():
    sd = Path("./setting")
    if not sd.exists():
        return []
    return sorted([p.name for p in sd.iterdir()
                   if p.is_dir() and any(f.name.startswith("Writing-DNA") for f in p.iterdir())])

def load_dna(author, book_title=None):
    """优先加载单本书DNA，回退到作者全局DNA"""
    if book_title and book_title != "全部作品":
        p = Path("./setting") / author / f"Writing-DNA_{book_title}.md"
        if p.exists():
            return p.read_text(encoding="utf-8")
    p = Path("./setting") / author / "Writing-DNA.md"
    if p.exists():
        return p.read_text(encoding="utf-8")
    return ""

def get_collection(author):
    return client.get_or_create_collection(name=f"novel_{author}")

def list_books_of_author(author):
    try:
        col = get_collection(author)
        data = col.get(include=["metadatas"])
        books = set()
        for m in data["metadatas"]:
            if m and "book_title" in m and m["book_title"]:
                books.add(m["book_title"])
        return sorted(books)
    except Exception:
        return []

def split_text(text, size=500, overlap=100):
    chunks, start = [], 0
    while start < len(text):
        end = start + size
        chunks.append(text[start:end])
        start = end - overlap
        if start >= len(text):
            break
    return chunks

def read_text_file(path):
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix in (".txt", ".md", ".text"):
        return path.read_text(encoding="utf-8", errors="ignore")
    elif suffix == ".docx":
        from docx import Document
        doc = Document(str(path))
        paras = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        paras.append(cell.text)
        return "\n\n".join(paras)
    elif suffix == ".doc":
        try:
            import win32com.client as win32
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(path.resolve()))
            text = doc.Content.Text
            doc.Close(False)
            word.Quit()
            return text
        except Exception as e:
            raise Exception(f"读取 .doc 失败：{e}\n请确认已安装 Microsoft Word")
    else:
        raise Exception(f"不支持的文件类型：{suffix}")

def ingest_author(author, book_title, source_dirs=None):
    col = get_collection(author)
    try:
        col.delete(where={})
    except Exception:
        pass
    dirs = source_dirs if source_dirs else [Path("./raw") / author]
    total = 0
    for d in dirs:
        d = Path(d)
        if not d.exists():
            continue
        for f in list(d.glob("*.txt")) + list(d.glob("*.md")) + \
                list(d.glob("*.docx")) + list(d.glob("*.doc")):
            text = read_text_file(f)
            chunks = split_text(text)
            for i, ch in enumerate(chunks):
                emb = ollama.embeddings(model=EMBED_MODEL, prompt=ch)["embedding"]
                col.add(
                    ids=[f"{f.stem}_{i}_{total}"],
                    documents=[ch],
                    embeddings=[emb],
                    metadatas={"source": f.name, "book_title": book_title}
                )
                total += 1
    return total

def delete_author(author, delete_raw=False):
    sd = Path("./setting") / author
    if sd.exists():
        shutil.rmtree(sd)
    try:
        client.delete_collection(f"novel_{author}")
    except Exception:
        pass
    if delete_raw:
        rd = Path("./raw") / author
        if rd.exists():
            shutil.rmtree(rd)

def load_skill(skill_dir:str):
    """加载skill，读取SKILL.md，提取system prompt，读取模板文件"""
    skill_root = Path(skill_dir)
    skill_md = skill_root / "SKILL.md"
    if not skill_md.exists():
        raise FileNotFoundError(f"找不到skill配置：{skill_md}，请确认skill目录创建完成")
    skill_text = skill_md.read_text(encoding="utf-8")
    # 提取System Prompt
    sys_pat = re.search(r"## System Prompt\s*\n(.*?)(?=\n##|\Z)", skill_text, re.DOTALL)
    system_prompt = sys_pat.group(1).strip() if sys_pat else ""
    # 提取模板路径
    tp_pat = re.search(r"## Template Path\s*\n(.*?)(?=\n##|\Z)", skill_text, re.DOTALL)
    template_rel_path = tp_pat.group(1).strip() if tp_pat else ""
    template_file = skill_root / template_rel_path
    if not template_file.exists():
        raise FileNotFoundError(f"找不到模板文件 {template_file}")
    template_content = template_file.read_text(encoding="utf-8")
    return system_prompt, template_content

#世界观分析函数
def extract_worldbuilding(author, book_title, distill_model_type,
                          api_client=None, api_model=None, ollama_model=None):
    # 加载世界观skill
    system_prompt, world_template = load_skill("./skill/worldbuilding-skill")
    # 读取该作品向量库
    col = get_collection(author)
    res = col.get(where={"book_title": book_title}, include=["documents"])
    docs = res.get("documents", [])
    if not docs:
        raise Exception(f"《{book_title}》还未上传入库，请先上传文本")
    # 采样原文片段，最多8段，控制总长度
    full_text = "\n".join(docs)
    step = max(3000, len(full_text) // 8)
    segs = []
    pos = 0
    while pos < len(full_text):
        segs.append(full_text[pos:pos+3000])
        pos += step
    sample_segs = segs[:8]
    corpus_text = "\n====原文片段分割线====\n".join(sample_segs)[:20000]
    user_msg = f"【原文片段】\n{corpus_text}\n\n【输出模板】\n{world_template}\n\n填充模板，生成《{book_title}》世界观设定文档。"
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_msg}
    ]
    # 区分云端API / Ollama本地模型，和你现有逻辑对齐
    if distill_model_type == "api":
        resp = api_client.chat.completions.create(model=api_model, messages=messages)
        return resp.choices[0].message.content
    else:
        resp = ollama.chat(model=ollama_model, messages=messages)
        return resp["message"]["content"]
# 角色资料卡分析函数（从skill加载prompt和模板）
def extract_character_card(author, book_title, character_name, distill_model_type,
                            api_client=None, api_model=None, ollama_model=None):
    # 加载角色卡skill
    system_prompt, char_template = load_skill("./skill/character-card-skill")

    # 读取该作品向量库
    col = get_collection(author)
    # RAG检索和该角色相关片段
    query_emb = ollama.embeddings(model=EMBED_MODEL, prompt=character_name)["embedding"]
    rag_res = col.query(query_embeddings=[query_emb], n_results=6,
                        where={"book_title": book_title})
    rag_docs = rag_res["documents"][0]
    if not rag_docs:
        raise Exception(f"《{book_title}》里没检索到和「{character_name}」相关的片段")
    corpus_text = "\n====原文片段分割线====\n".join(rag_docs)[:15000]

    user_msg = (f"角色名：{character_name}，作品《{book_title}》\n"
                f"【原文片段】\n{corpus_text}\n\n"
                f"【输出模板】\n{char_template}\n\n按模板生成角色资料卡。")
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_msg}
    ]
    # 区分云端API / Ollama本地模型
    if distill_model_type == "api":
        resp = api_client.chat.completions.create(model=api_model, messages=messages)
        return resp.choices[0].message.content
    else:
        resp = ollama.chat(model=ollama_model, messages=messages)
        return resp["message"]["content"]

# ==========分析函数（支持单本/全局两种模式） ==========
def distill_author(author, mode, book_title, distill_model_type,
                   api_client=None, api_model=None, ollama_model=None):
    """
    mode="book":   分析指定 book_title 的书（从向量库取素材）
    mode="author":分析作者全部作品（从 raw 文件取素材，原逻辑）
    """
    # ---- 1. 准分析素材 ----
    if mode == "book":
        if not book_title:
            raise Exception("单分析必须指定书名")
        col = get_collection(author)
        res = col.get(where={"book_title": book_title}, include=["documents"])
        docs = res.get("documents", [])
        if not docs:
            raise Exception(f"《{book_title}》还没入库，请先上传并入库这本书")
        full_text = "\n".join(docs)
        dna_filename = f"Writing-DNA_{book_title}.md"
    else:
        raw_dir = Path("./raw") / author
        files = list(raw_dir.glob("*.txt")) + list(raw_dir.glob("*.md")) + \
                list(raw_dir.glob("*.docx")) + list(raw_dir.glob("*.doc"))
        if not files:
            raise Exception(f"{author}/ 里没有文本文件")
        parts = []
        for f in files:
            parts.append(read_text_file(f))
        full_text = "\n".join(parts)
        dna_filename = "Writing-DNA.md"
    # 均匀采样最多 4 段，共 12000 字
    text_len = len(full_text)
    step = max(2000, text_len // 4)
    segs = []
    pos = 0
    while pos < text_len:
        segs.append(full_text[pos:pos+2000])
        pos += step
    sample = segs[:4]
    corpus = "\n====片段分隔====\n".join(sample)[:12000]
    # ---- 2. 加载模板 ----
    tpl_root = Path("./skill/writing-dna-skill/templates/author-corpus/zh")
    tpl_files = ["Writing-DNA.md", "语言DNA.md", "文章结构模板.md",
                 "写作视角与认知框架.md", "视觉风格指南.md"]
    tpl_text = ""
    for t in tpl_files:
        fp = tpl_root / t
        if fp.exists():
            tpl_text += f"\n====模板:{t}====\n{fp.read_text(encoding='utf-8')}\n"
    sys_prompt = """你是文学风格分析器。基于原文片段，严格依照模板生成Writing-DNA。
要求：
1. 只输出语言层风格规则（句式、节奏、词汇、修辞、叙事视角、对话写法）
2. 不要剧情总结、不要读后感、不要把具体人名/书名/世界观写进风格规则
3. 题材专有物只作为比喻素材举例，不作为风格主体"""
    user_prompt = f"【原文片段】\n{corpus}\n\n【输出模板】\n{tpl_text}\n\n按模板输出。"
    messages = [{"role":"system","content":sys_prompt},
                {"role":"user","content":user_prompt}]
    # ---- 3. 调用模型 ----
    if distill_model_type == "api":
        resp = api_client.chat.completions.create(model=api_model, messages=messages)
        dna_content = resp.choices[0].message.content
    else:
        resp = ollama.chat(model=ollama_model, messages=messages)
        dna_content = resp["message"]["content"]
    # ---- 4. 保存DNA ----
    out_dir = Path("./setting") / author
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / dna_filename).write_text(dna_content, encoding="utf-8")
    # 全分析才重新入库；单分析素材已在库中，不重复入库
    if mode == "author":
        chunk_count = ingest_author(author, book_title="(全部作品)")
    else:
        chunk_count = len(docs) if mode == "book" else 0
    return dna_content, chunk_count

def fuse_authors(author_a, author_b, new_name, distill_model_type,
                 api_client=None, api_model=None, ollama_model=None):
    dna_a = load_dna(author_a)
    dna_b = load_dna(author_b)
    if not dna_a or not dna_b:
        raise Exception("两个作者必须先完分析！")
    sys_prompt = """你是风格融合器。融合两份写作DNA，保留各自特点，冲突处标注优先级，只输出新规则。"""
    user_prompt = f"【A】\n{dna_a}\n\n【B】\n{dna_b}\n\n比较出「{new_name}」的DNA。"
    messages = [{"role":"system","content":sys_prompt},
                {"role":"user","content":user_prompt}]
    if distill_model_type == "api":
        resp = api_client.chat.completions.create(model=api_model, messages=messages)
        new_dna = resp.choices[0].message.content
    else:
        resp = ollama.chat(model=ollama_model, messages=messages)
        new_dna = resp["message"]["content"]
    out_dir = Path("./setting") / new_name
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "Writing-DNA.md").write_text(new_dna, encoding="utf-8")
    chunk_count = ingest_author(new_name, book_title="(比较)",
                                source_dirs=[Path("./raw")/author_a, Path("./raw")/author_b])
    return new_dna, chunk_count

# ========== 侧边栏（全部控制面板都放这里！） ==========
with st.sidebar:
    st.header("📖 多作者小说RAG")
    PROVIDERS = {
        "智谱GLM(免费)": {
            "base_url": "https://open.bigmodel.cn/api/paas/v4",
            "models": ["glm-4.7-flash", "glm-4-flash", "glm-4-air"]
        },
        "硅基流动": {
            "base_url": "https://api.siliconflow.cn/v1",
            "models": ["Qwen/Qwen2.5-7B-Instruct", "THUDM/glm-4-9b-chat",
                       "deepseek-ai/DeepSeek-V3"]
        },
        "通义千问": {
            "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
            "models": ["qwen-turbo", "qwen-plus", "qwen-long"]
        },
        "DeepSeek": {
            "base_url": "https://api.deepseek.com",
            "models": ["deepseek-chat", "deepseek-reasoner"]
        },
    }
    use_api = st.checkbox("使用云端API推理", value=False)
    work_model = None
    api_client = None
    api_model = ""
    if use_api:
        provider_name = st.selectbox("选择API服务商", list(PROVIDERS.keys()))
        cfg = PROVIDERS[provider_name]
        api_key = st.text_input(f"{provider_name} API Key", type="password")
        api_model = st.selectbox("模型", cfg["models"])
        if api_key:
            api_client = OpenAI(api_key=api_key, base_url=cfg["base_url"])
    else:
        try:
            local_models = [m["model"] for m in ollama.list()["models"]]
            work_model = st.selectbox("本地Ollama模型", local_models)
        except Exception:
            st.error("Ollama未启动"); work_model = None
    authors = list_authors()
    dna_authors = list_dna_authors()
    st.subheader("🎯 当前写作对象")
    cur_author = None
    if dna_authors:
        cur_author = st.selectbox("选择作者", dna_authors)
    else:
        st.info("还没有任何DNA，请分析")
    cur_book = "全部作品"
    if cur_author:
        books = list_books_of_author(cur_author)
        if books:
            cur_book = st.selectbox("限定作品", ["全部作品"] + books)
    st.divider()
    st.subheader("📤 上传文本")
    combined = []
    seen = set()
    for name in dna_authors:
        if name not in seen:
            seen.add(name); combined.append(name)
    for name in authors:
        if name not in seen:
            seen.add(name); combined.append(name)
    upload_target = st.selectbox("存入哪个作者", ["（新建作者）"] + combined, key="upload_target")
    if upload_target == "（新建作者）":
        new_author_name = st.text_input("新作者名", value="")
        target_author = new_author_name.strip()
    else:
        target_author = upload_target
    upload_book = st.text_input("作品名", value="")
    uploaded = st.file_uploader("拖入文件（.txt/.md/.docx/.doc）",
                                type=["txt","md","docx","doc"], accept_multiple_files=True)
    if uploaded and target_author and upload_book.strip():
        save_dir = Path("./raw") / target_author
        save_dir.mkdir(parents=True, exist_ok=True)
        for f in uploaded:
            (save_dir / f.name).write_bytes(f.getbuffer())
        st.success(f"已保存 {len(uploaded)} 个文件到 raw/{target_author}/")
        if st.button("（请按）保存并入库"):
            with st.status("入库中...", expanded=True) as s:
                try:
                    n = ingest_author(target_author, book_title=upload_book.strip())
                    s.update(label=f"✅ 入库 {n} 片段", state="complete")
                    st.rerun()
                except Exception as e:
                    s.update(label="❌ 失败", state="error"); st.error(str(e))
    elif uploaded and target_author and not upload_book.strip():
        st.warning("请先填作品名再上传")
    st.divider()
    st.subheader("🧪 分析")
    distill_mode = st.radio("分析模式", ["作者风格", "单本风格"], horizontal=True)
    if authors:
        sel_distill = st.selectbox("选作者文件夹", authors)
        distill_book = None
        if distill_mode == "单本风格":
            # 从该作者向量库里取已入库的书
            books_for_distill = list_books_of_author(sel_distill)
            if not books_for_distill:
                st.warning("该作者还没有入库的书，请先上传并入库")
            else:
                distill_book = st.selectbox("选要分析的书", books_for_distill)
        if st.button("开始分析"):
            with st.status("分析中...", expanded=True) as s:
                try:
                    if use_api:
                        _, n = distill_author(sel_distill,
                                              mode="book" if distill_mode=="单本风格" else "author",
                                              book_title=distill_book,
                                              distill_model_type="api",
                                              api_client=api_client, api_model=api_model)
                    else:
                        _, n = distill_author(sel_distill,
                                              mode="book" if distill_mode=="单本风格" else "author",
                                              book_title=distill_book,
                                              distill_model_type="local",
                                              ollama_model=work_model)
                    s.update(label=f"✅分析完成（{n} 片段参与）", state="complete")
                    st.rerun()
                except Exception as e:
                    s.update(label="❌ 失败", state="error"); st.error(str(e))
    st.divider()
        # ===== 世界观模块 =====
    st.subheader("🌍 世界观设定整理")
    # 初始化session状态标记
    if "show_world_doc" not in st.session_state:
        st.session_state.show_world_doc = False

    if cur_book and cur_book != "全部作品":
        if st.button("📄 生成本书世界观设定集", help="从当前选中作品语料提取世界观，自动保存md文件"):
            with st.status("正在解析原文，提取世界观设定...", expanded=True) as status:
                try:
                    if use_api:
                        world_result = extract_worldbuilding(
                            cur_author, cur_book, "api",
                            api_client=api_client,
                            api_model=api_model
                        )
                    else:
                        world_result = extract_worldbuilding(
                            cur_author, cur_book, "local",
                            ollama_model=work_model
                        )
                    # 写入md文件
                    save_dir = Path("./setting") / cur_author
                    save_dir.mkdir(parents=True, exist_ok=True)
                    save_file = save_dir / f"世界观_{cur_book}.md"
                    save_file.write_text(world_result, encoding="utf-8")
                    status.update(label="✅ 世界观整理完成，文件已保存", state="complete")
                    st.session_state.show_world_doc = True #生成完自动打开
                except Exception as err:
                    status.update(label="❌ 任务失败", state="error")
                    st.error(f"错误：{str(err)}")

        # 读取已保存世界观文档按钮
        world_md_path = Path("./setting") / cur_author / f"世界观_{cur_book}.md"
        if world_md_path.exists():
            if st.button("📖 读取已保存世界观文档"):
                # 切换显示状态
                st.session_state.show_world_doc = not st.session_state.show_world_doc

            # expander折叠框，支持手动收起
            with st.expander("📖 世界观文档", expanded=st.session_state.show_world_doc):
                load_text = world_md_path.read_text(encoding="utf-8")
                st.markdown(load_text)
        else:
            st.info("💡 暂未找到本书的世界观文档，请先生成")
    else:
        st.info("💡 请先在【限定作品】下拉框选中目标书籍，再生成世界观")
    
    # ===== 角色资料卡模块 =====
    st.divider()
    st.subheader("👤 角色资料卡")
    if "show_char_card" not in st.session_state:
        st.session_state.show_char_card = False

    if cur_book and cur_book != "全部作品":
        char_save_dir = Path("./setting") / cur_author
        char_save_dir.mkdir(parents=True, exist_ok=True)

        # 自动扫描当前本书已经存在的角色md，提取角色名
        existed_char_list = []
        for f in char_save_dir.glob(f"角色_{cur_book}_*.md"):
            filename = f.name
            prefix = f"角色_{cur_book}_"
            if filename.startswith(prefix) and filename.endswith(".md"):
                char_name_parsed = filename[len(prefix):-3]
                existed_char_list.append(char_name_parsed)
        existed_char_list = sorted(existed_char_list)

        # 下拉选择已有角色 + 手动输入框
        select_opt = ["（新建角色，手动输入）"] + existed_char_list
        selected_char = st.selectbox("📋 已生成角色列表", options=select_opt, key="char_select_box")
        if selected_char != "（新建角色，手动输入）":
            char_name = selected_char
        else:
            char_name = st.text_input("角色名称", placeholder="输入书中角色名字", key="char_name_input")

        char_file_path = char_save_dir / f"角色_{cur_book}_{char_name.strip()}.md"

        if char_name.strip():
            col_gen, col_del = st.columns(2)
            with col_gen:
                if st.button("👤 生成角色资料卡", help="从当前作品语料提取角色信息，保存md"):
                    with st.status("正在检索原文，生成角色资料...", expanded=True) as status:
                        try:
                            if use_api:
                                char_result = extract_character_card(
                                    cur_author, cur_book, char_name.strip(), "api",
                                    api_client=api_client, api_model=api_model
                                )
                            else:
                                char_result = extract_character_card(
                                    cur_author, cur_book, char_name.strip(), "local",
                                    ollama_model=work_model
                                )
                            save_file = char_save_dir / f"角色_{cur_book}_{char_name.strip()}.md"
                            save_file.write_text(char_result, encoding="utf-8")
                            status.update(label="✅ 角色资料卡生成完成，已保存", state="complete")
                            st.session_state.show_char_card = True
                        except Exception as err:
                            status.update(label="❌ 任务失败", state="error")
                            st.error(f"错误：{str(err)}")
            with col_del:
                if char_file_path.exists():
                    if st.button("🗑️ 删除角色卡", help="永久删除该角色md文件", type="secondary"):
                        char_file_path.unlink()
                        st.warning(f"已删除：{char_name} 的角色资料卡，请刷新页面更新角色列表")
                        st.session_state.show_char_card = False

            if char_file_path.exists():
                if st.button("📖 读取已保存角色资料卡", key="read_char_btn"):
                    st.session_state.show_char_card = not st.session_state.show_char_card
                with st.expander("👤 角色资料卡文档", expanded=st.session_state.show_char_card):
                    char_text = char_file_path.read_text(encoding="utf-8")
                    st.markdown(char_text)
            else:
                st.info("💡 暂无该角色资料卡，请先生成")
    else:
        st.info("💡 请先在【限定作品】下拉框选中目标书籍")

    # ===== 比较两个作者 =====
    st.subheader("🔀 比较两个作者")
    if len(dna_authors) >= 2:
        a1 = st.selectbox("作者1", dna_authors, key="fuse_a")
        a2 = st.selectbox("作者2", [x for x in dna_authors if x != a1], key="fuse_b")
        new_name = st.text_input("新作者名", value="name")
        if st.button("比较"):
            with st.status("比较中...", expanded=True) as s:
                try:
                    if use_api:
                        _, n = fuse_authors(a1, a2, new_name, "api",
                                            api_client=api_client, api_model=api_model)
                    else:
                        _, n = fuse_authors(a1, a2, new_name, "local",
                                            ollama_model=work_model)
                    s.update(label=f"✅ 比较完成（{n} 片段）", state="complete")
                    st.rerun()
                except Exception as e:
                    s.update(label="❌ 失败", state="error"); st.error(str(e))
    st.divider()
    # ===== 删除模块=====
    st.subheader("🗑️ 删除")
    if dna_authors:
        del_target = st.selectbox("选择要删除的", dna_authors, key="del_target")
        del_raw = st.checkbox("同时删原文文件", value=False)
        if st.button("确认删除", type="primary"):
            delete_author(del_target, delete_raw=del_raw)
            st.success(f"已删除 {del_target}"); st.rerun()
    st.divider()
    # ===== 检索数量滑块 =====
    retrieve_cnt = st.slider("检索片段数", 1, 8, 2)
# ========== 侧边栏 END ==========

# ========== 主页面 ==========
st.title("📝作品语料检索与风格解析器")

if cur_author:
    dna = load_dna(cur_author, cur_book)
    dna_label = cur_book if cur_book != "全部作品" else "作者全局"
    with st.expander(f"📄 查看 {cur_author} 的 Writing-DNA（{dna_label}）", expanded=False):
        st.markdown(dna if dna else "（还没有DNA，请先进行分析）")

if "messages" not in st.session_state:
    st.session_state.messages = {}
conv_key = f"{cur_author}::{cur_book}"
if conv_key not in st.session_state["messages"]:
    st.session_state.messages[conv_key] = []

for msg in st.session_state.messages[conv_key]:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

user_input = st.chat_input("需要检索什么...")
retrieved_text = ""
if user_input and cur_author:
    st.chat_message("user").markdown(user_input)
    st.session_state.messages[conv_key].append({"role":"user","content":user_input})
    dna = load_dna(cur_author, cur_book)
    col = get_collection(cur_author)
    # 手动用 bge-m3 生成 query 向量（避免 chroma 自带 embedding 维度不匹配）
    query_emb = ollama.embeddings(model=EMBED_MODEL, prompt=user_input)["embedding"]
    query_kwargs = {"query_embeddings":[query_emb], "n_results":retrieve_cnt}
    if cur_book != "全部作品":
        query_kwargs["where"] = {"book_title": cur_book}
    res = col.query(**query_kwargs)
    # 片段去重
    docs_raw = res["documents"][0] if res["documents"] else []
    unique_docs, seen_text = [], set()
    for doc in docs_raw:
        key = doc.strip().replace("\n","")
        if key not in seen_text:
            seen_text.add(key)
            unique_docs.append(doc)
    retrieved_text = "\n\n".join(unique_docs)
    book_note = "" if cur_book == "全部作品" else f"\n【本次只参考《{cur_book}》，禁止引入其他作品人物/设定】"
    system_prompt = f"""【硬性规则，优先级最高】
1. 只用下方参考片段内容，禁止引入其他作品、网络资料或脑补设定
2. 片段里没有的信息直接说"资料不足"，不要编造
【写作DNA】
{dna}
【规则】直接输出正文，不要解释{book_note}
【参考片段】
{retrieved_text}
"""
    msgs = [{"role":"system","content":system_prompt}] + st.session_state.messages[conv_key]
    response = ""
    with st.chat_message("assistant"):
        placeholder = st.empty()
        if use_api:
            if api_client is None:
                st.error("请先填 API Key")
            else:
                stream = api_client.chat.completions.create(model=api_model, messages=msgs, stream=True)
                for chunk in stream:
                    if chunk.choices and chunk.choices[0].delta.content:
                        response += chunk.choices[0].delta.content
                        placeholder.markdown(response)
        else:
            stream = ollama.chat(model=work_model, messages=msgs, stream=True)
            for chunk in stream:
                response += chunk["message"]["content"]
                placeholder.markdown(response)
    st.session_state.messages[conv_key].append({"role":"assistant","content":response})

with st.expander("🔍 调试：本次检索片段"):
    st.text_area("", retrieved_text, height=250)
