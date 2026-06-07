<template>
  <div class="container">
    <h2>文件预览</h2>

    <input type="file" accept=".docx,.pdf,.xlsx,.xls" @change="handleFile" />

    <div ref="previewRef" class="preview"></div>
  </div>
</template>

<script setup>import { ref, onMounted } from "vue";

const previewRef = ref(null);

let pdfjsLib = null;
let XLSX = null;
let mammoth = null;

/**
 * 动态加载JS
 */
const loadScript = (src) => {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");

    script.src = src;
    script.async = true;

    script.onload = () => resolve();

    script.onerror = () =>
      reject(new Error(`加载失败: ${src}`));

    document.head.appendChild(script);
  });
};

/**
 * 初始化
 */
onMounted(async () => {
  try {

    await loadScript(
      "./lib/pdf.min.js"
    );

    await loadScript(
      "./lib/xlsx.full.min.js"
    );

    await loadScript(
      "./lib/mammoth.browser.min.js"
    );


    pdfjsLib = window.pdfjsLib;
    XLSX = window.XLSX;
    mammoth = window.mammoth;

    console.log(
      "mammoth加载成功",
      mammoth
    );

    if (!pdfjsLib) {
      throw new Error(
        "pdfjsLib加载失败"
      );
    }

    if (!XLSX) {
      throw new Error(
        "XLSX加载失败"
      );
    }

    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "./lib/pdf.worker.js";

    console.log(
      "PDF加载成功",
      pdfjsLib
    );

    console.log(
      "Excel加载成功",
      XLSX
    );

  } catch (err) {

    console.error(
      "库加载失败",
      err
    );

  }
});

/**
 * 获取后缀
 */
const getFileExt = (fileName) => {

  const index =
    fileName.lastIndexOf(".");

  if (index === -1) {
    return "";
  }

  return fileName
    .substring(index + 1)
    .toLowerCase();
};

/**
 * 清空预览区域
 */
const clearPreview = () => {

  if (previewRef.value) {
    previewRef.value.innerHTML = "";
  }

};

/**
 * Word预览
 */
const previewDocx = async (file) => {

  clearPreview();

  const arrayBuffer =
    await file.arrayBuffer();

  const result =
    await mammoth.convertToHtml(
      {
        arrayBuffer
      },
      {
        convertImage:
          mammoth.images.inline(
            (image) => {

              return image
                .read("base64")
                .then((data) => {

                  return {
                    src:
                      `data:${image.contentType};base64,${data}`
                  };

                });

            }
          )
      }
    );

  previewRef.value.innerHTML =
    result.value;

};

/**
 * Excel预览
 */
const previewExcel = async (file) => {

  clearPreview();

  const arrayBuffer =
    await file.arrayBuffer();

  const workbook =
    XLSX.read(arrayBuffer, {
      type: "array"
    });

  const firstSheetName =
    workbook.SheetNames[0];

  const worksheet =
    workbook.Sheets[firstSheetName];

  const html =
    XLSX.utils.sheet_to_html(
      worksheet
    );

  previewRef.value.innerHTML =
    html;

};

/**
 * PDF预览
 */
const previewPdf = async (file) => {

  clearPreview();

  const arrayBuffer =
    await file.arrayBuffer();

  const loadingTask =
    pdfjsLib.getDocument({
      data: arrayBuffer
    });

  const pdf =
    await loadingTask.promise;

  for (
    let pageNum = 1;
    pageNum <= pdf.numPages;
    pageNum++
  ) {

    const page =
      await pdf.getPage(pageNum);

    const viewport =
      page.getViewport({
        scale: 1.5
      });

    const canvas =
      document.createElement(
        "canvas"
      );

    const context =
      canvas.getContext(
        "2d"
      );

    canvas.width =
      viewport.width;

    canvas.height =
      viewport.height;

    canvas.style.display =
      "block";

    canvas.style.margin =
      "0 auto 20px";

    await page.render({
      canvasContext:
        context,
      viewport
    }).promise;

    previewRef.value.appendChild(
      canvas
    );
  }
};

/**
 * 文件选择
 */
const handleFile = async (e) => {

  const file =
    e.target.files[0];

  if (!file) {
    return;
  }

  const ext =
    getFileExt(
      file.name
    );

  try {

    switch (ext) {

      case "docx":
        await previewDocx(
          file
        );
        break;

      case "pdf":
        await previewPdf(
          file
        );
        break;

      case "xlsx":
      case "xls":
        await previewExcel(
          file
        );
        break;

      default:

        clearPreview();

        previewRef.value.innerHTML =
          `
          <div class="message">
            当前仅支持 DOCX、PDF、Excel 文件预览
          </div>
          `;

    }

  } catch (err) {

    console.error(err);

    clearPreview();

    previewRef.value.innerHTML =
      `
      <div class="message error">
        文件预览失败
      </div>
      `;

  }

};
</script>

<style scoped>
.container {
  padding: 20px;
}

.preview {
  margin-top: 20px;
  min-height: 600px;
  border: 1px solid #ddd;
  padding: 20px;
  overflow: auto;
  background: #fafafa;
}

.message {
  text-align: center;
  padding: 40px;
  color: #666;
}

.error {
  color: red;
}

.preview table {
  border-collapse: collapse;
}

.preview td,
.preview th {
  border: 1px solid #ddd;
  padding: 8px;
}
</style>