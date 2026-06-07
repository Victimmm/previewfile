<template>
  <div class="fp-wrap">
    <h2 class="fp-title">文件预览</h2>

    <!-- URL 输入区 -->
    <div class="fp-url-row">
      <div class="fp-input-group">
        <i class="ti ti-link" aria-hidden="true"></i>
        <input
          v-model="urlInput"
          class="fp-url-input"
          type="text"
          placeholder="输入文件 URL（支持 PDF、Word、Excel）"
          @keyup.enter="handleUrlAdd"
        />
      </div>
      <button class="fp-btn fp-btn--primary" @click="handleUrlAdd">
        <i class="ti ti-plus" aria-hidden="true"></i>添加
      </button>
    </div>

    <!-- 本地上传区 -->
    <div
      class="fp-drop-zone"
      :class="{ 'fp-drop-zone--active': isDragging }"
      @dragover.prevent="isDragging = true"
      @dragleave.prevent="isDragging = false"
      @drop.prevent="handleDrop"
      @click="triggerFileInput"
    >
      <input
        ref="fileInputRef"
        type="file"
        accept="*/*"
        multiple
        class="fp-hidden-input"
        @change="handleLocalFiles"
      />
      <i class="ti ti-cloud-upload fp-drop-icon" aria-hidden="true"></i>
      <p class="fp-drop-text">拖拽文件到此处，或 <span class="fp-drop-link">点击选择</span></p>
      <p class="fp-drop-hint">支持所有文件类型·仅 PDF / Word / Excel 可预览</p>
    </div>

    <!-- 文件列表 -->
    <transition-group name="fp-list" tag="div" class="fp-file-list" v-if="files.length">
      <div
        v-for="item in files"
        :key="item.id"
        class="fp-file-row"
        :class="{ 'fp-file-row--active': previewingId === item.id }"
      >
        <div class="fp-file-icon" :class="`fp-file-icon--${item.ext}`">
          <i :class="`ti ${extIcon(item.ext)}`" aria-hidden="true"></i>
        </div>
        <div class="fp-file-info">
          <span class="fp-file-name" :title="item.name">{{ item.name }}</span>
          <span class="fp-file-meta">{{ item.sizeLabel }} · {{ extLabel(item.ext) }}</span>
        </div>
        <div class="fp-file-actions">
          <button
            v-if="canPreview(item.ext)"
            class="fp-btn fp-btn--ghost"
            :class="{ 'fp-btn--active': previewingId === item.id }"
            @click="openPreview(item)"
            :title="previewingId === item.id ? '关闭预览' : '预览'"
          >
            <i :class="`ti ${previewingId === item.id ? 'ti-eye-off' : 'ti-eye'}`" aria-hidden="true"></i>
            {{ previewingId === item.id ? '关闭' : '预览' }}
          </button>
          <span v-else class="fp-no-preview">不支持预览</span>
          <button class="fp-btn fp-btn--danger" @click="removeFile(item.id)" title="删除">
            <i class="ti ti-trash" aria-hidden="true"></i>
          </button>
        </div>
      </div>
    </transition-group>

    <!-- 空状态 -->
    <div v-if="!files.length" class="fp-empty">
      <i class="ti ti-files fp-empty-icon" aria-hidden="true"></i>
      <p>暂无文件，请上传或输入 URL</p>
    </div>

    <!-- 预览区 -->
    <transition name="fp-preview">
      <div v-if="previewingId" class="fp-preview-panel">
        <div class="fp-preview-header">
          <span class="fp-preview-title">
            <i :class="`ti ${extIcon(previewingItem?.ext)}`" aria-hidden="true"></i>
            {{ previewingItem?.name }}
          </span>
          <button class="fp-btn fp-btn--ghost" @click="closePreview">
            <i class="ti ti-x" aria-hidden="true"></i>关闭
          </button>
        </div>

        <div class="fp-preview-body">
          <!-- 加载中 -->
          <div v-if="previewState === 'loading'" class="fp-state-box">
            <div class="fp-spinner"></div>
            <p>正在加载预览…</p>
          </div>
          <!-- 错误 -->
          <div v-else-if="previewState === 'error'" class="fp-state-box fp-state-box--error">
            <i class="ti ti-alert-circle" aria-hidden="true"></i>
            <p>{{ previewError }}</p>
          </div>
          <!-- 内容 -->
          <div
            v-else-if="previewState === 'done'"
            ref="previewRef"
            class="fp-preview-content"
          ></div>
        </div>
      </div>
    </transition>

    <!-- 库加载失败提示 -->
    <div v-if="libError" class="fp-lib-error">
      <i class="ti ti-alert-triangle" aria-hidden="true"></i>
      预览库加载失败，请检查 <code>/lib/</code> 目录下的依赖文件。
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, nextTick } from "vue";

/* ─── 库引用 ─── */
let pdfjsLib = null;
let XLSX = null;
let mammoth = null;
const libError = ref(false);

const loadScript = (src) =>
  new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.async = true;
    s.onload = resolve;
    s.onerror = () => reject(new Error(`加载失败: ${src}`));
    document.head.appendChild(s);
  });

onMounted(async () => {
  try {
    await loadScript("./lib/pdf.min.js");
    await loadScript("./lib/xlsx.full.min.js");
    await loadScript("./lib/mammoth.browser.min.js");
    pdfjsLib = window.pdfjsLib;
    XLSX = window.XLSX;
    mammoth = window.mammoth;
    if (!pdfjsLib || !XLSX || !mammoth) throw new Error("部分库未挂载到 window");
    pdfjsLib.GlobalWorkerOptions.workerSrc = "./lib/pdf.worker.js";
  } catch (e) {
    console.error("预览库加载失败", e);
    libError.value = true;
  }
});

/* ─── 文件列表 ─── */
let _id = 0;
const files = ref([]);
const urlInput = ref("");
const isDragging = ref(false);
const fileInputRef = ref(null);

const triggerFileInput = () => fileInputRef.value?.click();

const getExt = (name) => {
  const i = name.lastIndexOf(".");
  return i >= 0 ? name.substring(i + 1).toLowerCase() : "";
};

const formatSize = (bytes) => {
  if (!bytes) return "URL";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
};

const addFiles = (rawFiles) => {
  for (const f of rawFiles) {
    files.value.push({
      id: ++_id,
      name: f.name,
      ext: getExt(f.name),
      sizeLabel: formatSize(f.size),
      source: "local",
      file: f,
    });
  }
};

const handleLocalFiles = (e) => addFiles(Array.from(e.target.files));
const handleDrop = (e) => {
  isDragging.value = false;
  addFiles(Array.from(e.dataTransfer.files));
};

const handleUrlAdd = () => {
  const url = "http://127.0.0.1:5050/proxy?url=" + urlInput.value.trim();
  if (!url) return;
  const name = url.split("/").pop().split("?")[0] || "远程文件";
  const ext = getExt(name);
  files.value.push({
    id: ++_id,
    name,
    ext,
    sizeLabel: "URL",
    source: "url",
    url,
  });
  urlInput.value = "";
};

const removeFile = (id) => {
  if (previewingId.value === id) closePreview();
  files.value = files.value.filter((f) => f.id !== id);
};

/* ─── 图标 / 标签 ─── */
const extIcon = (ext) => {
  const map = {
    pdf: "ti-file-type-pdf",
    docx: "ti-file-type-doc",
    doc: "ti-file-type-doc",
    xlsx: "ti-file-spreadsheet",
    xls: "ti-file-spreadsheet",
  };
  return map[ext] || "ti-file";
};

const extLabel = (ext) => {
  const map = { pdf: "PDF", docx: "Word", doc: "Word", xlsx: "Excel", xls: "Excel" };
  return map[ext] || ext?.toUpperCase() || "文件";
};

const canPreview = (ext) => ["pdf", "docx", "doc", "xlsx", "xls"].includes(ext);

/* ─── 预览逻辑 ─── */
const previewingId = ref(null);
const previewState = ref("idle"); // idle | loading | done | error
const previewError = ref("");
const previewRef = ref(null);

const previewingItem = computed(() => files.value.find((f) => f.id === previewingId.value));

const closePreview = () => {
  previewingId.value = null;
  previewState.value = "idle";
  if (previewRef.value) previewRef.value.innerHTML = "";
};

const openPreview = async (item) => {
  if (previewingId.value === item.id) {
    closePreview();
    return;
  }
  previewingId.value = item.id;
  previewState.value = "loading";
  previewError.value = "";

  await nextTick();

  try {
    let arrayBuffer;
    if (item.source === "url") {
      const res = await fetch(item.url);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      arrayBuffer = await res.arrayBuffer();
    } else {
      arrayBuffer = await item.file.arrayBuffer();
    }

    switch (item.ext) {
      case "pdf":
        await renderPdf(arrayBuffer);
        break;
      case "docx":
      case "doc":
        await renderDocx(arrayBuffer);
        break;
      case "xlsx":
      case "xls":
        await renderExcel(arrayBuffer);
        break;
      default:
        throw new Error("不支持的格式");
    }

    previewState.value = "done";
  } catch (e) {
    console.error(e);
    previewState.value = "error";
    previewError.value = e.message || "预览失败，请检查文件格式或 URL 是否可访问";
  }
};

const renderPdf = async (arrayBuffer) => {
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  previewState.value = "done";
  await nextTick();
  const container = previewRef.value;
  container.innerHTML = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 1.5 });
    const canvas = document.createElement("canvas");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    canvas.style.cssText = "display:block;margin:0 auto 16px;border-radius:4px;";
    await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
    container.appendChild(canvas);
  }
};

const renderDocx = async (arrayBuffer) => {
  const result = await mammoth.convertToHtml(
    { arrayBuffer },
    {
      convertImage: mammoth.images.inline((img) =>
        img.read("base64").then((data) => ({ src: `data:${img.contentType};base64,${data}` }))
      ),
    }
  );
  previewState.value = "done";
  await nextTick();
  previewRef.value.innerHTML = result.value;
};

const renderExcel = async (arrayBuffer) => {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  previewState.value = "done";
  await nextTick();
  const container = previewRef.value;
  container.innerHTML = "";

  wb.SheetNames.forEach((sheetName, idx) => {
    const tab = document.createElement("div");
    tab.className = "fp-sheet";
    const label = document.createElement("div");
    label.className = "fp-sheet-label";
    label.textContent = sheetName;
    const html = XLSX.utils.sheet_to_html(wb.Sheets[sheetName]);
    const tableWrap = document.createElement("div");
    tableWrap.className = "fp-table-wrap";
    tableWrap.innerHTML = html;
    tab.appendChild(label);
    tab.appendChild(tableWrap);
    container.appendChild(tab);
  });
};
</script>

<style scoped>
/* ── 字体变量 ── */
.fp-wrap {
  --accent: #2563eb;
  --accent-light: #eff6ff;
  --accent-border: #bfdbfe;
  --danger: #dc2626;
  --danger-light: #fef2f2;
  --radius-sm: 6px;
  --radius-md: 10px;
  --radius-lg: 14px;
  font-family: "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
  padding: 28px 24px;
  max-width: 860px;
  margin: 0 auto;
  color: #1e293b;
}

/* ── 标题 ── */
.fp-title {
  font-size: 20px;
  font-weight: 600;
  margin: 0 0 20px;
  letter-spacing: -0.3px;
}

/* ── URL 输入行 ── */
.fp-url-row {
  display: flex;
  gap: 10px;
  margin-bottom: 14px;
}

.fp-input-group {
  flex: 1;
  display: flex;
  align-items: center;
  gap: 10px;
  border: 1px solid #e2e8f0;
  border-radius: var(--radius-md);
  padding: 0 12px;
  background: #fff;
  transition: border-color 0.15s, box-shadow 0.15s;
}

.fp-input-group:focus-within {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
}

.fp-input-group i {
  font-size: 17px;
  color: #94a3b8;
  flex-shrink: 0;
}

.fp-url-input {
  flex: 1;
  border: none;
  outline: none;
  font-size: 14px;
  padding: 10px 0;
  background: transparent;
  color: #1e293b;
}

.fp-url-input::placeholder {
  color: #94a3b8;
}

/* ── 拖拽区 ── */
.fp-drop-zone {
  border: 2px dashed #cbd5e1;
  border-radius: var(--radius-lg);
  padding: 28px 20px;
  text-align: center;
  cursor: pointer;
  transition: border-color 0.2s, background 0.2s;
  background: #f8fafc;
  margin-bottom: 20px;
  user-select: none;
}

.fp-drop-zone:hover,
.fp-drop-zone--active {
  border-color: var(--accent);
  background: var(--accent-light);
}

.fp-hidden-input {
  display: none;
}

.fp-drop-icon {
  font-size: 32px;
  color: #94a3b8;
  display: block;
  margin-bottom: 10px;
}

.fp-drop-zone:hover .fp-drop-icon,
.fp-drop-zone--active .fp-drop-icon {
  color: var(--accent);
}

.fp-drop-text {
  font-size: 14px;
  color: #475569;
  margin: 0 0 4px;
}

.fp-drop-link {
  color: var(--accent);
  font-weight: 500;
}

.fp-drop-hint {
  font-size: 12px;
  color: #94a3b8;
  margin: 0;
}

/* ── 文件列表 ── */
.fp-file-list {
  display: flex;
  flex-direction: column;
  gap: 8px;
  margin-bottom: 20px;
}

.fp-file-row {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 10px 14px;
  border: 1px solid #e2e8f0;
  border-radius: var(--radius-md);
  background: #fff;
  transition: border-color 0.15s, box-shadow 0.15s;
}

.fp-file-row:hover {
  border-color: #cbd5e1;
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.06);
}

.fp-file-row--active {
  border-color: var(--accent-border);
  background: var(--accent-light);
}

/* 文件类型图标 */
.fp-file-icon {
  width: 36px;
  height: 36px;
  border-radius: var(--radius-sm);
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 18px;
  flex-shrink: 0;
  background: #f1f5f9;
  color: #64748b;
}

.fp-file-icon--pdf {
  background: #fef2f2;
  color: #dc2626;
}

.fp-file-icon--docx,
.fp-file-icon--doc {
  background: #eff6ff;
  color: #2563eb;
}

.fp-file-icon--xlsx,
.fp-file-icon--xls {
  background: #f0fdf4;
  color: #16a34a;
}

.fp-file-info {
  flex: 1;
  min-width: 0;
}

.fp-file-name {
  display: block;
  font-size: 14px;
  font-weight: 500;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  color: #1e293b;
}

.fp-file-meta {
  display: block;
  font-size: 12px;
  color: #94a3b8;
  margin-top: 2px;
}

.fp-file-actions {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-shrink: 0;
}

.fp-no-preview {
  font-size: 12px;
  color: #94a3b8;
  padding: 0 4px;
}

/* ── 按钮 ── */
.fp-btn {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  padding: 6px 12px;
  border-radius: var(--radius-sm);
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  border: 1px solid transparent;
  transition: background 0.15s, border-color 0.15s, color 0.15s;
  white-space: nowrap;
}

.fp-btn i {
  font-size: 15px;
}

.fp-btn--primary {
  background: var(--accent);
  color: #fff;
  border-color: var(--accent);
}

.fp-btn--primary:hover {
  background: #1d4ed8;
  border-color: #1d4ed8;
}

.fp-btn--ghost {
  background: transparent;
  color: #475569;
  border-color: #e2e8f0;
}

.fp-btn--ghost:hover {
  background: #f1f5f9;
  border-color: #cbd5e1;
  color: #1e293b;
}

.fp-btn--ghost.fp-btn--active {
  background: var(--accent-light);
  color: var(--accent);
  border-color: var(--accent-border);
}

.fp-btn--danger {
  background: transparent;
  color: #ef4444;
  border-color: #fecaca;
}

.fp-btn--danger:hover {
  background: var(--danger-light);
  border-color: #fca5a5;
}

/* ── 空状态 ── */
.fp-empty {
  text-align: center;
  padding: 40px 20px;
  color: #94a3b8;
}

.fp-empty-icon {
  font-size: 40px;
  display: block;
  margin-bottom: 10px;
}

.fp-empty p {
  margin: 0;
  font-size: 14px;
}

/* ── 预览面板 ── */
.fp-preview-panel {
  border: 1px solid #e2e8f0;
  border-radius: var(--radius-lg);
  overflow: hidden;
  margin-top: 4px;
}

.fp-preview-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px 16px;
  border-bottom: 1px solid #e2e8f0;
  background: #f8fafc;
}

.fp-preview-title {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 14px;
  font-weight: 500;
  color: #334155;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.fp-preview-title i {
  font-size: 17px;
  flex-shrink: 0;
}

.fp-preview-body {
  min-height: 400px;
  background: #fff;
  overflow: auto;
}

/* ── 状态（加载/错误）── */
.fp-state-box {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 14px;
  min-height: 300px;
  color: #64748b;
  font-size: 14px;
}

.fp-state-box i {
  font-size: 36px;
}

.fp-state-box p {
  margin: 0;
  max-width: 320px;
  text-align: center;
}

.fp-state-box--error {
  color: #dc2626;
}

.fp-spinner {
  width: 36px;
  height: 36px;
  border: 3px solid #e2e8f0;
  border-top-color: var(--accent);
  border-radius: 50%;
  animation: fp-spin 0.7s linear infinite;
}

@keyframes fp-spin {
  to {
    transform: rotate(360deg);
  }
}

/* ── 预览内容 ── */
.fp-preview-content {
  padding: 24px 28px;
  font-size: 15px;
  line-height: 1.7;
}

.fp-preview-content :deep(img) {
  max-width: 100%;
  height: auto;
}

.fp-preview-content :deep(table) {
  border-collapse: collapse;
  width: 100%;
  font-size: 13px;
}

.fp-preview-content :deep(td),
.fp-preview-content :deep(th) {
  border: 1px solid #e2e8f0;
  padding: 6px 10px;
}

.fp-preview-content :deep(th) {
  background: #f8fafc;
  font-weight: 500;
}

/* Excel 多 sheet */
.fp-sheet + .fp-sheet {
  margin-top: 28px;
}

.fp-sheet-label {
  font-size: 13px;
  font-weight: 600;
  color: #16a34a;
  margin-bottom: 10px;
  padding: 4px 10px;
  background: #f0fdf4;
  border-radius: var(--radius-sm);
  display: inline-block;
}

.fp-table-wrap {
  overflow-x: auto;
}

/* ── 库加载错误 ── */
.fp-lib-error {
  margin-top: 16px;
  padding: 12px 16px;
  background: #fffbeb;
  border: 1px solid #fde68a;
  border-radius: var(--radius-md);
  font-size: 13px;
  color: #92400e;
  display: flex;
  align-items: center;
  gap: 8px;
}

.fp-lib-error i {
  font-size: 17px;
  flex-shrink: 0;
}

.fp-lib-error code {
  background: #fef3c7;
  padding: 1px 5px;
  border-radius: 3px;
  font-family: monospace;
}

/* ── 列表动画 ── */
.fp-list-enter-active,
.fp-list-leave-active {
  transition: all 0.2s ease;
}

.fp-list-enter-from,
.fp-list-leave-to {
  opacity: 0;
  transform: translateY(-6px);
}

/* ── 预览面板动画 ── */
.fp-preview-enter-active,
.fp-preview-leave-active {
  transition: opacity 0.2s, transform 0.2s;
}

.fp-preview-enter-from,
.fp-preview-leave-to {
  opacity: 0;
  transform: translateY(8px);
}
</style>