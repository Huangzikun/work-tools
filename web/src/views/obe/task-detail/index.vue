<script setup lang="ts">
import { computed, h, onMounted, onUnmounted, ref, watch } from 'vue';
import { useRoute, useRouter } from 'vue-router';
import {
  NButton,
  NCard,
  NCheckbox,
  NDataTable,
  NDatePicker,
  NForm,
  NFormItem,
  NInput,
  NInputNumber,
  NModal,
  NProgress,
  NScrollbar,
  NSelect,
  NSpace,
  NTag,
  NText,
  NTree,
  NUpload
} from 'naive-ui';
import type { DataTableColumns, UploadFileInfo } from 'naive-ui';
import {
  cancelObeGrade,
  deleteObeTask,
  downloadObeExcel,
  downloadObeStudentFile,
  downloadObeTaskAll,
  downloadObeZip,
  fetchObeProgress,
  fetchObeRubric,
  fetchObeTaskDetail,
  fetchObeTaskTree,
  generateObeRubric,
  resolveObeAmbiguous,
  retryObeStudent,
  startObeGrade,
  uploadObeStudentFile,
  uploadObeStudentFiles
} from '@/service/api/obe';

defineOptions({ name: 'ObeTaskDetail' });

const route = useRoute();
const router = useRouter();

const taskId = computed(() => Number(route.query.taskId || 0));

function errMsg(err: any, fallback: string): string {
  return err?.response?.data?.msg || err?.msg || err?.message || fallback;
}

// ============ 任务详情数据 ============
const loading = ref(false);
const task = ref<Api.Obe.TaskSummary | null>(null);
const studentsByDirExp = ref<Record<string, Record<string, Api.Obe.Student[]>>>({});
const latestJobByDirExp = ref<Record<string, Record<string, Api.Obe.GradingJob>>>({});
const experimentsByDir = ref<Record<string, string[]>>({});
const tree = ref<Api.Obe.TreeNode | null>(null);

const activeDirType = ref<string>('');
const activeExperiment = ref<string>('');

const dirTypeOptions = computed(() => {
  if (!task.value) return [];
  return (task.value.studentDirTypes || []).map(d => ({ label: d, value: d }));
});

const experimentOptions = computed(() => {
  if (!activeDirType.value) return [];
  const labels = experimentsByDir.value[activeDirType.value] || [];
  return labels.map(label => ({
    label: label === 'default' ? '默认（未上传）' : label,
    value: label
  }));
});

const activeStudents = computed(() => {
  if (!activeDirType.value || !activeExperiment.value) return [];
  const exps = studentsByDirExp.value[activeDirType.value] || {};
  return exps[activeExperiment.value] || [];
});

const activeJob = computed(() => {
  if (!activeDirType.value || !activeExperiment.value) return null;
  const exps = latestJobByDirExp.value[activeDirType.value] || {};
  return exps[activeExperiment.value] || null;
});

async function loadDetail() {
  loading.value = true;
  try {
    const { data, error } = await fetchObeTaskDetail(taskId.value);
    if (error) {
      window.$message?.error(errMsg(error, '加载任务失败'));
      return;
    }
    if (!data) return;
    task.value = data.task;
    studentsByDirExp.value = data.studentsByDirAndExperiment;
    latestJobByDirExp.value = data.latestJobByDirAndExperiment;
    experimentsByDir.value = data.experimentsByDir;

    // 初始化默认 dir_type / experiment
    if (!activeDirType.value && (task.value?.studentDirTypes || []).length > 0) {
      activeDirType.value = task.value!.studentDirTypes[0];
    }
    // 如果当前 experiment 不在最新列表里，重置为第一个
    const exps = experimentsByDir.value[activeDirType.value] || [];
    if (exps.length > 0 && !exps.includes(activeExperiment.value)) {
      activeExperiment.value = exps[0];
    }
  } finally {
    loading.value = false;
  }
}

async function loadTree() {
  const { data, error } = await fetchObeTaskTree(taskId.value);
  if (error) return;
  if (data) tree.value = data.tree;
}

watch(taskId, () => {
  if (taskId.value > 0) {
    loadDetail();
    loadTree();
  }
});

// dir_type 切换时，重置 experiment 到该 dir_type 的第一个
watch(activeDirType, () => {
  const exps = experimentsByDir.value[activeDirType.value] || [];
  if (exps.length > 0) {
    activeExperiment.value = exps[0];
  } else {
    activeExperiment.value = '';
  }
});

onMounted(() => {
  if (taskId.value > 0) {
    loadDetail();
    loadTree();
  }
});

// ============ 上传学生文件 ============
const uploading = ref(false);
const uploadProgress = ref(0);

// 临时存最近一次上传结果，用于 ambiguous 解决（按 experiment 分组）
const lastUploadExperiments = ref<Api.Obe.ExperimentUploadResult[]>([]);

async function handleUploadFiles(files: File[]) {
  if (!activeDirType.value) {
    window.$message?.error('请先选择考核目录');
    return;
  }
  uploading.value = true;
  uploadProgress.value = 0;
  try {
    const result = await uploadObeStudentFiles(taskId.value, activeDirType.value, files, {
      onUploadProgress: e => {
        if (e.total) uploadProgress.value = Math.round((e.loaded / e.total) * 100);
      }
    });
    lastUploadExperiments.value = result.experiments;

    const totalMatched = result.experiments.reduce((sum, e) => sum + e.matched.length, 0);
    const totalAmbiguous = result.experiments.reduce((sum, e) => sum + e.ambiguous.length, 0);
    const totalUnmatched = result.experiments.reduce((sum, e) => sum + e.unmatched.length, 0);
    const totalReset = result.experiments.reduce(
      (sum, e) => sum + e.matched.filter(m => m.resetPreviousGrading).length,
      0
    );

    const resetHint = totalReset > 0 ? `，重置 ${totalReset} 个已批改学生（需重新批改）` : '';
    window.$message?.success(
      `上传完成（${result.experiments.length} 个实验）：成功匹配 ${totalMatched}，歧义 ${totalAmbiguous}，未识别 ${totalUnmatched}${resetHint}`
    );

    await loadDetail();

    // 自动切到最新上传的 experiment
    if (result.experiments.length > 0) {
      activeExperiment.value = result.experiments[result.experiments.length - 1].experimentLabel;
    }
  } catch (err) {
    window.$message?.error(err instanceof Error ? err.message : '上传失败');
  } finally {
    uploading.value = false;
    uploadProgress.value = 0;
  }
}

function onUploadChange(options: { fileList: UploadFileInfo[] }) {
  const ready = options.fileList.filter(f => f.file && f.status !== 'error');
  if (ready.length === 0) return;
  const files = ready.map(f => f.file!).filter(Boolean);
  if (files.length === 0) return;
  handleUploadFiles(files);
  setTimeout(() => {
    uploadFileList.value = [];
  }, 100);
}

const uploadFileList = ref<UploadFileInfo[]>([]);

// ============ 解决 ambiguous ============
const ambiguousModalVisible = ref(false);
const ambiguousExperimentLabel = ref('');
const ambiguousResolutions = ref<Array<{ fileName: string; filePath: string; studentId: string }>>([]);
const ambiguousItems = ref<Api.Obe.AmbiguousItem[]>([]);

function openAmbiguousModal() {
  // 找第一个有 ambiguous 的 experiment
  const expWithAmbiguous = lastUploadExperiments.value.find(e => e.ambiguous.length > 0);
  if (!expWithAmbiguous) {
    window.$message?.info('没有歧义文件需要解决');
    return;
  }
  ambiguousExperimentLabel.value = expWithAmbiguous.experimentLabel;
  ambiguousItems.value = expWithAmbiguous.ambiguous;
  ambiguousResolutions.value = expWithAmbiguous.ambiguous.map(a => ({
    fileName: a.fileName,
    filePath: a.filePath,
    studentId: ''
  }));
  ambiguousModalVisible.value = true;
}

// 给 ambiguous 文件提供候选学生 + 全部学生（兜底）
function buildCandidateOptions(idx: number) {
  const candidates = ambiguousItems.value[idx]?.candidates || [];
  const allInExp = (studentsByDirExp.value[activeDirType.value]?.[ambiguousExperimentLabel.value] || []).map(s => ({
    studentId: s.studentId,
    studentName: s.studentName
  }));
  const seen = new Set(candidates.map(c => c.studentId));
  for (const a of allInExp) if (!seen.has(a.studentId)) {
    candidates.push(a);
    seen.add(a.studentId);
  }
  return candidates.map(c => ({ label: `${c.studentId} ${c.studentName}`, value: c.studentId }));
}

async function submitAmbiguous() {
  const valid = ambiguousResolutions.value.filter(r => r.studentId);
  if (valid.length === 0) {
    window.$message?.warning('请至少选择一个归属');
    return;
  }
  const { data, error } = await resolveObeAmbiguous(
    taskId.value,
    activeDirType.value,
    ambiguousExperimentLabel.value,
    valid
  );
  if (error) {
    window.$message?.error(errMsg(error, '解决失败'));
    return;
  }
  window.$message?.success(`已解决 ${data?.resolved?.length || 0} 个`);
  ambiguousModalVisible.value = false;
  lastUploadExperiments.value = lastUploadExperiments.value.filter(
    e => e.experimentLabel !== ambiguousExperimentLabel.value || e.ambiguous.length === 0
  );
  await loadDetail();
}

// ============ 触发批改 ============
const gradeModalVisible = ref(false);
const gradeForm = ref({
  teacherName: '',
  signDate: Math.floor(Date.now() / 1000),
  rubricDimensions: [] as Api.Obe.RubricDimension[],
  freeText: '',
  scoreLevels: null as number[] | null,
  signPicture: null as File | null,
  overwriteGraded: false
});
const gradeFormRef = ref();
const signPictureFileList = ref<UploadFileInfo[]>([]);

// 示范评分标准模板（贴合「人工智能与创新设计」，教师可一键载入后修改）
const DEMO_RUBRIC_TEMPLATE = {
  dimensions: [
    { name: '实验目的与原理', maxScore: 15, criteria: '是否清晰阐述 AI 原理（CNN/Transformer/扩散模型）、应用场景与实验目标' },
    { name: '环境搭建与数据准备', maxScore: 15, criteria: '环境（框架版本/GPU）、数据来源与预处理是否完整可复现' },
    { name: '模型设计与实现', maxScore: 25, criteria: '模型结构/超参是否合理，代码是否完整可运行，创新点说明' },
    { name: '实验结果与分析', maxScore: 25, criteria: '指标（准确率/FID/loss）是否真实，有无对比/消融/可视化' },
    { name: '创新与反思', maxScore: 15, criteria: '有无独立思考、改进思路、局限性与伦理反思，非简单复现' },
    { name: '报告规范性', maxScore: 5, criteria: '格式、图表、引用、语言是否规范' }
  ],
  freeText:
    '本课程为「人工智能与创新设计」，重点考察对 AI 原理的理解与创新设计。\n评语结构：亮点（1-2句）→ 不足（具体到节/图）→ 改进建议。\n分数按 90/80/70/60/50/0 六档；仅跑通 demo 无分析或抄袭给 50 以下。\n评分须引用学生报告原文。',
  scoreLevels: [90, 80, 70, 60, 50, 0]
};

function addDimension() {
  gradeForm.value.rubricDimensions.push({ name: '', maxScore: 10, criteria: '' });
}
function removeDimension(idx: number) {
  gradeForm.value.rubricDimensions.splice(idx, 1);
}
function loadRubricTemplate() {
  gradeForm.value.rubricDimensions = DEMO_RUBRIC_TEMPLATE.dimensions.map(d => ({ ...d }));
  gradeForm.value.freeText = DEMO_RUBRIC_TEMPLATE.freeText;
  gradeForm.value.scoreLevels = [...DEMO_RUBRIC_TEMPLATE.scoreLevels];
  window.$message?.success('已载入示范模板');
}
async function loadLastRubric() {
  if (!task.value?.courseName) return;
  try {
    const { data: rec, error } = await fetchObeRubric(task.value.courseName);
    if (error) {
      window.$message?.error('载入失败');
      return;
    }
    if (rec) {
      gradeForm.value.rubricDimensions = rec.dimensions.map(d => ({
        name: d.name,
        maxScore: d.maxScore,
        criteria: d.criteria
      }));
      gradeForm.value.freeText = rec.freeText;
      gradeForm.value.scoreLevels = rec.scoreLevels ?? null;
      window.$message?.success('已载入上次标准');
    } else {
      window.$message?.info('该课程暂无保存的标准');
    }
  } catch {
    window.$message?.error('载入失败');
  }
}

const aiInput = ref('');
const aiGenerating = ref(false);

async function generateRubric() {
  if (!aiInput.value.trim()) {
    window.$message?.error('请先粘贴实验内容');
    return;
  }
  aiGenerating.value = true;
  try {
    const { data, error } = await generateObeRubric(aiInput.value.trim());
    if (error || !data) {
      window.$message?.error('生成失败');
      return;
    }
    gradeForm.value.rubricDimensions = data.dimensions.map(d => ({
      name: d.name,
      maxScore: d.max_score,
      criteria: d.criteria
    }));
    gradeForm.value.freeText = data.freeText;
    window.$message?.success('AI 已生成评分标准，可微调后批改');
  } catch {
    window.$message?.error('生成失败');
  } finally {
    aiGenerating.value = false;
  }
}

function openGradeModal() {
  if (!activeDirType.value) {
    window.$message?.error('请先选择考核目录');
    return;
  }
  if (!activeExperiment.value || activeExperiment.value === 'default') {
    window.$message?.error('请先选择一个具体的实验（"默认"占位实验不能直接批改，先上传 ZIP 自动生成实验）');
    return;
  }
  const matchedCount = activeStudents.value.filter(s => s.matched).length;
  if (matchedCount === 0) {
    window.$message?.error('该实验下没有已匹配的学生文件，请先上传');
    return;
  }
  gradeForm.value = {
    teacherName: task.value?.teacherName || '',
    signDate: Math.floor(Date.now() / 1000),
    rubricDimensions: [],
    freeText: '',
    scoreLevels: null,
    signPicture: null,
    overwriteGraded: false
  };
  signPictureFileList.value = [];
  aiInput.value = '';
  // 自动载入该课程上次的标准，方便复用；无则空着由教师填写
  if (task.value?.courseName) {
    fetchObeRubric(task.value.courseName)
      .then(({ data: rec }) => {
        if (rec) {
          gradeForm.value.rubricDimensions = rec.dimensions.map(d => ({
            name: d.name,
            maxScore: d.maxScore,
            criteria: d.criteria
          }));
          gradeForm.value.freeText = rec.freeText;
          gradeForm.value.scoreLevels = rec.scoreLevels ?? null;
        }
      })
      .catch(() => {});
  }
  gradeModalVisible.value = true;
}

function onSignPictureChange(options: { fileList: UploadFileInfo[] }) {
  signPictureFileList.value = options.fileList;
  gradeForm.value.signPicture = options.fileList[0]?.file ?? null;
}

async function submitGrade() {
  if (!gradeForm.value.teacherName.trim()) {
    window.$message?.error('请填写教师姓名');
    return;
  }
  if (!gradeForm.value.signDate) {
    window.$message?.error('请选择签名日期');
    return;
  }
  if (!gradeForm.value.freeText.trim() && !gradeForm.value.rubricDimensions.some(d => d.name.trim())) {
    window.$message?.error('请填写评分标准（总体要求或至少一个评分维度）');
    return;
  }
  if (!gradeForm.value.signPicture) {
    window.$message?.error('请上传签名图片');
    return;
  }

  const date = new Date(gradeForm.value.signDate * 1000);
  const signDateStr = `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;

  // 只提交有名字的有效维度
  const rubricDimensions = gradeForm.value.rubricDimensions.filter(d => d.name.trim());

  try {
    const result = await startObeGrade(taskId.value, {
      dirType: activeDirType.value,
      experimentLabel: activeExperiment.value,
      teacherName: gradeForm.value.teacherName.trim(),
      signDate: signDateStr,
      teacherPrompt: gradeForm.value.freeText.trim(),
      rubricDimensions,
      scoreLevels: gradeForm.value.scoreLevels,
      signPicture: gradeForm.value.signPicture,
      // overwriteGraded=true → skipGraded=false（全量重跑）；默认 false → skipGraded=true（增量）
      skipGraded: !gradeForm.value.overwriteGraded
    });
    window.$message?.success(`批改已开始（jobId=${result.jobId}）`);
    gradeModalVisible.value = false;
    startProgressPolling(result.jobId);
  } catch (err) {
    window.$message?.error(err instanceof Error ? err.message : '触发批改失败');
  }
}

// ============ 进度轮询 ============
const progress = ref<Api.Obe.Progress | null>(null);
const cancelling = ref(false);
let progressTimer: ReturnType<typeof setInterval> | null = null;

const progressPercent = computed(() => {
  const p = progress.value;
  if (!p || !p.total) return 0;
  return Math.round(((p.graded + p.failed) / p.total) * 100);
});

// NProgress 的 status：终态着色并停止动画；running 时不传（默认色 + active 动画）
const progressStatus = computed<'success' | 'error' | 'warning' | undefined>(() => {
  switch (progress.value?.status) {
    case 'completed':
      return 'success';
    case 'failed':
      return 'error';
    case 'cancelled':
      return 'warning';
    default:
      return undefined;
  }
});

function startProgressPolling(jobId: number) {
  stopProgressPolling();
  progress.value = { total: 0, graded: 0, failed: 0, status: 'running', jobId };
  progressTimer = setInterval(async () => {
    const { data, error } = await fetchObeProgress(
      taskId.value,
      activeDirType.value,
      activeExperiment.value,
      jobId
    );
    if (error) return;
    if (data) {
      progress.value = data;
      await loadDetail();
      if (data.status && data.status !== 'running') {
        stopProgressPolling();
        cancelling.value = false;
        if (data.status === 'cancelled') {
          window.$message?.warning(`批改已取消（已批改 ${data.graded}/${data.total}）`);
        } else {
          window.$message?.success(
            `批改完成（成功 ${data.graded}/${data.total}，失败 ${data.failed}）`
          );
        }
      }
    }
  }, 3000);
}

function stopProgressPolling() {
  if (progressTimer) {
    clearInterval(progressTimer);
    progressTimer = null;
  }
}

/** 取消批改：协作式，后端跑完当前学生后才会真正停，故按钮进入 loading 直到轮询到 cancelled。 */
function handleCancelGrade() {
  const jobId = progress.value?.jobId;
  if (!jobId) return;
  window.$dialog?.warning({
    title: '取消批改',
    content: '将停止批改剩余学生，已批改的成绩会保留。当前正在批改的学生会跑完后再停止（约 10-30 秒）。',
    positiveText: '确认取消',
    negativeText: '继续批改',
    onPositiveClick: async () => {
      cancelling.value = true;
      const { error } = await cancelObeGrade(taskId.value, jobId);
      if (error) {
        cancelling.value = false;
        window.$message?.error(errMsg(error, '取消失败'));
        return;
      }
      window.$message?.info('已请求取消，等待当前学生批改完成…');
    }
  });
}

onUnmounted(stopProgressPolling);

// ============ 单学生批改（上传后自动批改）============
/** 乐观更新：立即把学生置为「批改中」并清空分数/评语/错误，让用户马上看到状态变化。
 *  真实结果由 gradeOneStudent 的轮询拿到后端最终状态覆盖。 */
function setStudentGrading(studentPk: number) {
  for (const [dt, expMap] of Object.entries(studentsByDirExp.value)) {
    for (const [exp, list] of Object.entries(expMap)) {
      const idx = list.findIndex(s => s.id === studentPk);
      if (idx < 0) continue;
      const newList = list.slice();
      newList[idx] = {
        ...list[idx],
        gradeStatus: 'grading',
        lastScore: undefined,
        lastComment: undefined,
        lastGradeError: undefined
      };
      studentsByDirExp.value = {
        ...studentsByDirExp.value,
        [dt]: { ...expMap, [exp]: newList }
      };
      return;
    }
  }
}

/** 用单学生操作返回的最新数据就地刷新本地列表（整体替换触发 activeStudents 重算）。 */
function upsertStudent(updated: Api.Obe.Student) {
  const dirMap = studentsByDirExp.value[updated.dirType];
  const list = dirMap?.[updated.experimentLabel];
  if (!list) return;
  const idx = list.findIndex(s => s.id === updated.id);
  if (idx < 0) return;
  const newList = list.slice();
  newList[idx] = updated;
  studentsByDirExp.value = {
    ...studentsByDirExp.value,
    [updated.dirType]: { ...dirMap, [updated.experimentLabel]: newList }
  };
}

function findStudentInState(studentPk: number): Api.Obe.Student | undefined {
  for (const expMap of Object.values(studentsByDirExp.value)) {
    for (const list of Object.values(expMap)) {
      const s = list.find(x => x.id === studentPk);
      if (s) return s;
    }
  }
  return undefined;
}

/**
 * 单学生批改流程：乐观置「批改中」+ 置空分数评语 → 后台发起批改 → 轮询直到完成自动更新。
 *
 * 后端 retry 是同步阻塞（含 LLM + LibreOffice，可能数十秒），而 axios 实例全局 timeout=10s，
 * 所以不能 await retryObeStudent（必然超时）。改为：乐观更新 UI 后 fire-and-forget 发起请求，
 * 用轮询 loadDetail 拉取真实状态，直到该学生脱离 grading（批改完成 / 失败）。
 */
function gradeOneStudent(studentPk: number) {
  setStudentGrading(studentPk);

  // fire-and-forget：10s 后 axios 会超时 abort，但后端仍在跑；超时错误忽略，靠轮询拿结果
  retryObeStudent(taskId.value, studentPk).catch(() => {});

  const intervalMs = 3000;
  const maxRounds = 100; // 5 分钟超时保护
  let rounds = 0;
  const timer = window.setInterval(async () => {
    rounds += 1;
    await loadDetail();
    const s = findStudentInState(studentPk);
    const done = !s || s.gradeStatus !== 'grading' || rounds >= maxRounds;
    if (!done) return;
    window.clearInterval(timer);
    if (s?.gradeStatus === 'graded') {
      window.$message?.success(`${s.studentName} 批改完成：${s.lastScore} 分`);
    } else if (s?.gradeStatus === 'failed') {
      window.$message?.error(`${s.studentName} 批改失败：${s.lastGradeError || '未知错误'}`);
    }
  }, intervalMs);
}

// ============ 单学生上传 / 下载 ============
async function handleDownloadStudent(row: Api.Obe.Student) {
  try {
    await downloadObeStudentFile(taskId.value, row.id);
    window.$message?.success('已开始下载');
  } catch (err) {
    window.$message?.error(err instanceof Error ? err.message : '下载失败');
  }
}

const uploadStudentModalVisible = ref(false);
const uploadStudentTarget = ref<Api.Obe.Student | null>(null);
const uploadStudentFileList = ref<UploadFileInfo[]>([]);
const uploadingStudent = ref(false);

function openUploadStudentModal(row: Api.Obe.Student) {
  uploadStudentTarget.value = row;
  uploadStudentFileList.value = [];
  uploadStudentModalVisible.value = true;
}

function onUploadStudentFileChange(options: { fileList: UploadFileInfo[] }) {
  // max=1：始终只保留最新选择的那一个
  uploadStudentFileList.value = options.fileList.slice(-1);
}

async function submitUploadStudent() {
  if (!uploadStudentTarget.value) return;
  const ready = uploadStudentFileList.value.filter(f => f.file);
  if (ready.length === 0) {
    window.$message?.warning('请先选择文件');
    return;
  }
  const file = ready[0].file!;
  uploadingStudent.value = true;
  try {
    const result = await uploadObeStudentFile(taskId.value, uploadStudentTarget.value.id, file);
    upsertStudent(result.student); // 先刷新为「已上传」（matched=true）
    uploadStudentModalVisible.value = false;
    window.$message?.success(`已上传「${result.fileName}」，开始批改...`);
    gradeOneStudent(result.student.id); // 上传后直接进入批改流程
  } catch (err) {
    window.$message?.error(err instanceof Error ? err.message : '上传失败');
  } finally {
    uploadingStudent.value = false;
  }
}

// ============ 下载 ============
async function handleDownloadZip() {
  if (!activeExperiment.value) {
    window.$message?.error('请先选择实验');
    return;
  }
  try {
    await downloadObeZip(taskId.value, activeDirType.value, activeExperiment.value);
  } catch (err) {
    window.$message?.error(err instanceof Error ? err.message : '下载失败');
  }
}

const downloading = ref(false);

async function handleDownloadAll() {
  downloading.value = true;
  try {
    await downloadObeTaskAll(taskId.value);
    window.$message?.success('已开始下载');
  } catch (err) {
    window.$message?.error(err instanceof Error ? err.message : '下载失败');
  } finally {
    downloading.value = false;
  }
}

async function handleDownloadExcel() {
  if (!activeExperiment.value || activeExperiment.value === 'default') {
    window.$message?.error('请先选择一个有批改记录的实验');
    return;
  }
  try {
    await downloadObeExcel(
      taskId.value,
      activeDirType.value,
      activeExperiment.value,
      activeJob.value?.id
    );
  } catch (err) {
    window.$message?.error(err instanceof Error ? err.message : '下载失败');
  }
}

// ============ 删除 ============
async function handleDelete() {
  const { error } = await deleteObeTask(taskId.value);
  if (error) {
    window.$message?.error(errMsg(error, '删除失败'));
    return;
  }
  window.$message?.success('已删除');
  router.push({ name: 'obe_tasks' });
}

// ============ 表格列 ============
function gradeStatusTag(s: string) {
  switch (s) {
    case 'pending':
      return h(NTag, { type: 'default', size: 'small' }, { default: () => '未批改' });
    case 'grading':
      return h(NTag, { type: 'info', size: 'small' }, { default: () => '批改中' });
    case 'graded':
      return h(NTag, { type: 'success', size: 'small' }, { default: () => '已批改' });
    case 'failed':
      return h(NTag, { type: 'error', size: 'small' }, { default: () => '失败' });
    default:
      return h(NTag, { type: 'default', size: 'small' }, { default: () => s });
  }
}

const studentColumns = computed<DataTableColumns<Api.Obe.Student>>(() => [
  { title: '学号', key: 'studentId', width: 120 },
  { title: '姓名', key: 'studentName', width: 90 },
  { title: '班级', key: 'studentClass', width: 140 },
  {
    title: '上传',
    key: 'matched',
    width: 90,
    render: row => (row.matched ? h(NTag, { type: 'success', size: 'small' }, { default: () => '已上传' }) : '-')
  },
  { title: '批改状态', key: 'gradeStatus', width: 100, render: row => gradeStatusTag(row.gradeStatus) },
  { title: '分数', key: 'lastScore', width: 70, render: row => (row.lastScore != null ? `${row.lastScore}` : '-') },
  { title: '评语', key: 'lastComment', minWidth: 200, ellipsis: { tooltip: true } },
  {
    title: '操作',
    key: 'actions',
    width: 220,
    render: row =>
      h(
        NSpace,
        { size: 4, wrap: false },
        {
          default: () => [
            row.matched
              ? h(
                  NButton,
                  { size: 'small', tertiary: true, onClick: () => handleDownloadStudent(row) },
                  { default: () => '下载' }
                )
              : null,
            h(
              NButton,
              { size: 'small', tertiary: true, onClick: () => openUploadStudentModal(row) },
              { default: () => '上传' }
            ),
            row.matched
              ? h(
                  NButton,
                  { size: 'small', tertiary: true, onClick: () => openUploadStudentModal(row) },
                  { default: () => '重新批改' }
                )
              : null
          ]
        }
      )
  }
]);

const treeData = computed(() => {
  if (!tree.value) return [];
  return [convertTreeNode(tree.value)];
});

function convertTreeNode(node: Api.Obe.TreeNode): any {
  return {
    key: node.key,
    label: node.label,
    children: node.children?.map(convertTreeNode) || [],
    isLeaf: node.type === 'file'
  };
}

// 是否有歧义待解决
const hasAmbiguousToResolve = computed(() =>
  lastUploadExperiments.value.some(e => e.ambiguous.length > 0)
);
</script>

<template>
  <NSpace vertical :size="16">
    <!-- 任务信息 -->
    <NCard :bordered="false" class="card-wrapper">
      <NSpace justify="space-between" align="center">
        <div>
          <NSpace align="center" :size="12">
            <span class="text-18px font-medium">{{ task?.className }} 《{{ task?.courseName }}》</span>
            <NTag size="small">{{ task?.status }}</NTag>
          </NSpace>
          <div class="text-12px opacity-60 mt-2">
            教师：{{ task?.teacherName }} · 学生数：{{ task?.studentCount }} · 创建：{{
              task?.createdAt ? new Date(task.createdAt).toLocaleString('zh-CN') : '-'
            }}
          </div>
        </div>
        <NSpace>
          <NButton @click="router.push({ name: 'obe_tasks' })">返回列表</NButton>
          <NButton type="primary" :loading="downloading" @click="handleDownloadAll">打包下载整个任务</NButton>
          <NButton type="error" tertiary @click="handleDelete">删除任务</NButton>
        </NSpace>
      </NSpace>
    </NCard>

    <!-- 目录树 -->
    <NCard :bordered="false" class="card-wrapper" title="目录结构">
      <NScrollbar style="max-height: 320px">
        <NTree
          v-if="treeData.length > 0"
          :data="treeData"
          key-field="key"
          label-field="label"
          block-line
          expand-on-click
          :default-expanded-keys="['root']"
        />
        <NText v-else depth="3">目录树为空</NText>
      </NScrollbar>
    </NCard>

    <!-- 目录 + 实验选择 + 操作 -->
    <NCard :bordered="false" class="card-wrapper">
      <NSpace align="center" :size="12" style="margin-bottom: 12px" wrap>
        <span>考核目录：</span>
        <NSelect
          v-model:value="activeDirType"
          :options="dirTypeOptions"
          style="width: 200px"
        />
        <span>实验：</span>
        <NSelect
          v-model:value="activeExperiment"
          :options="experimentOptions"
          style="width: 320px"
        />
        <NUpload
          v-model:file-list="uploadFileList"
          multiple
          accept=".doc,.docx,.zip"
          :default-upload="false"
          :show-file-list="false"
          @change="onUploadChange"
        >
          <NButton :loading="uploading" type="primary">上传学生文件（ZIP/DOCX）</NButton>
        </NUpload>
        <NButton
          v-if="hasAmbiguousToResolve"
          type="warning"
          @click="openAmbiguousModal"
        >
          解决歧义文件
        </NButton>
        <NButton type="primary" @click="openGradeModal">开始批改</NButton>
        <NButton @click="handleDownloadZip">下载 ZIP</NButton>
        <NButton @click="handleDownloadExcel">下载成绩 Excel</NButton>
      </NSpace>

      <NProgress
        v-if="uploading && uploadProgress > 0"
        type="line"
        :percentage="uploadProgress"
        :show-indicator="true"
        style="margin-bottom: 12px"
      />

      <div
        v-if="progress && (progress.status === 'running' || progress.status === 'cancelled')"
        style="margin-bottom: 12px"
      >
        <NProgress
          type="line"
          :percentage="progressPercent"
          :status="progressStatus"
          :show-indicator="true"
        />
        <NSpace align="center" justify="space-between" style="margin-top: 8px">
          <NText depth="2" class="text-13px">
            批改进度：{{ progress.graded + progress.failed }}/{{ progress.total }}
            （成功 {{ progress.graded }}，失败 {{ progress.failed }}）<template v-if="progress.currentStudent">
              · 当前：{{ progress.currentStudent }}
            </template>
          </NText>
          <NButton
            v-if="progress.status === 'running'"
            type="error"
            tertiary
            size="small"
            :loading="cancelling"
            @click="handleCancelGrade"
          >
            取消批改
          </NButton>
          <NTag v-else-if="progress.status === 'cancelled'" type="warning" size="small">已取消</NTag>
        </NSpace>
      </div>

      <NDataTable
        :columns="studentColumns"
        :data="activeStudents"
        :loading="loading"
        :max-height="500"
      />
    </NCard>

    <!-- 最近上传结果（按实验分组） -->
    <NCard
      v-if="lastUploadExperiments.length > 0"
      :bordered="false"
      class="card-wrapper"
      title="最近一次上传结果（按实验分组）"
    >
      <NSpace vertical :size="12">
        <div v-for="exp in lastUploadExperiments" :key="exp.experimentLabel">
          <NText strong>{{ exp.experimentLabel }}</NText>
          <span class="ml-2 text-12px opacity-60">
            （文件 {{ exp.fileCount }}，匹配 {{ exp.matched.length }}，歧义 {{ exp.ambiguous.length }}，未识别 {{ exp.unmatched.length }}）
          </span>
          <div class="text-12px opacity-70 ml-4">
            <div v-if="exp.matched.length > 0">
              <NText>已匹配：</NText>
              {{ exp.matched.map(m => `${m.studentName}(${m.fileName})`).join('、') }}
            </div>
            <div v-if="exp.unmatched.length > 0">
              <NText>未识别：</NText>
              {{ exp.unmatched.map(u => u.fileName).join('、') }}
            </div>
          </div>
        </div>
      </NSpace>
    </NCard>

    <!-- 批改弹窗 -->
    <NModal
      v-model:show="gradeModalVisible"
      preset="card"
      :title="`开始批改 - ${activeExperiment}`"
      style="width: 720px"
      :mask-closable="false"
    >
      <NForm ref="gradeFormRef" label-placement="left" label-width="100">
        <NFormItem label="教师姓名" required>
          <NInput v-model:value="gradeForm.teacherName" placeholder="如：黄子坤" />
        </NFormItem>
        <NFormItem label="签名日期" required>
          <NDatePicker v-model:value="gradeForm.signDate" type="date" style="width: 100%" />
        </NFormItem>
        <NFormItem label="签名图片" required>
          <NUpload
            v-model:file-list="signPictureFileList"
            :max="1"
            accept=".png,.jpg,.jpeg"
            :default-upload="false"
            @change="onSignPictureChange"
          >
            <NButton>选择 PNG/JPG</NButton>
          </NUpload>
        </NFormItem>
        <NFormItem label="评分维度" :show-label="false">
          <NSpace vertical :size="8" style="width: 100%">
            <div style="background: rgba(99, 102, 241, 0.06); border: 1px solid rgba(99, 102, 241, 0.25); border-radius: 6px; padding: 8px">
              <NText depth="2" style="font-size: 13px">
                AI 智能生成（粘贴实验内容，生成「积极评分 + 严谨扣分」的标准）
              </NText>
              <NInput
                v-model:value="aiInput"
                type="textarea"
                :autosize="{ minRows: 2, maxRows: 5 }"
                placeholder="粘贴实验任务、考察点、预期产出……AI 会据此生成 4-6 个评分维度 + 总体要求"
                style="margin-top: 4px"
              />
              <NSpace :size="8" align="center" style="margin-top: 6px">
                <NButton size="tiny" type="primary" :loading="aiGenerating" @click="generateRubric">
                  AI 生成评分标准
                </NButton>
                <NText depth="3" style="font-size: 12px">生成后填入下方表单，可微调，不影响签名/日期</NText>
              </NSpace>
            </div>
            <NSpace justify="space-between" align="center">
              <NText depth="2" style="font-size: 13px">
                评分维度（可选，留空则按总体自由文本评）
              </NText>
              <NSpace :size="8">
                <NButton size="tiny" tertiary @click="loadRubricTemplate">载入示范模板</NButton>
                <NButton size="tiny" tertiary @click="loadLastRubric">载入上次（按课程）</NButton>
              </NSpace>
            </NSpace>
            <div
              v-for="(dim, idx) in gradeForm.rubricDimensions"
              :key="idx"
              style="border: 1px solid var(--n-border-color, #e0e0e6); border-radius: 6px; padding: 8px"
            >
              <NSpace :size="8" align="center" wrap>
                <NInput
                  v-model:value="dim.name"
                  placeholder="维度名（如：实验步骤）"
                  style="width: 200px"
                />
                <NInputNumber
                  v-model:value="dim.maxScore"
                  :min="1"
                  :max="100"
                  placeholder="满分"
                  style="width: 120px"
                />
                <NButton size="tiny" quaternary type="error" @click="removeDimension(idx)">
                  删除
                </NButton>
              </NSpace>
              <NInput
                v-model:value="dim.criteria"
                type="textarea"
                :autosize="{ minRows: 1, maxRows: 3 }"
                placeholder="该维度评分说明（教师标准）"
                style="margin-top: 6px"
              />
            </div>
            <NButton size="small" dashed block @click="addDimension">+ 添加维度</NButton>
          </NSpace>
        </NFormItem>
        <NFormItem label="总体要求" required>
          <NInput
            v-model:value="gradeForm.freeText"
            type="textarea"
            :autosize="{ minRows: 4, maxRows: 10 }"
            placeholder="总体批改要求 / 评语风格 / 分数档位说明（如：分数按 90/80/70/60/50 五档）"
          />
        </NFormItem>
        <NFormItem label=" ">
          <NCheckbox v-model:checked="gradeForm.overwriteGraded">
            覆盖已批改学生（不勾选时只批改「未批改 / 失败」的学生，已批改的跳过）
          </NCheckbox>
        </NFormItem>
      </NForm>
      <template #footer>
        <NSpace justify="end">
          <NButton @click="gradeModalVisible = false">取消</NButton>
          <NButton type="primary" @click="submitGrade">开始批改</NButton>
        </NSpace>
      </template>
    </NModal>

    <!-- 解决歧义弹窗 -->
    <NModal
      v-model:show="ambiguousModalVisible"
      preset="card"
      :title="`解决歧义文件归属 - ${ambiguousExperimentLabel}`"
      style="width: 640px"
      :mask-closable="false"
    >
      <NSpace vertical :size="12">
        <div v-for="(item, idx) in ambiguousItems" :key="idx">
          <NText strong>{{ item.fileName }}</NText>
          <NSelect
            v-model:value="ambiguousResolutions[idx].studentId"
            :options="buildCandidateOptions(idx)"
            filterable
            placeholder="选择归属学生"
            style="margin-top: 4px"
          />
        </div>
      </NSpace>
      <template #footer>
        <NSpace justify="end">
          <NButton @click="ambiguousModalVisible = false">取消</NButton>
          <NButton type="primary" @click="submitAmbiguous">提交</NButton>
        </NSpace>
      </template>
    </NModal>

    <!-- 单学生上传弹窗 -->
    <NModal
      v-model:show="uploadStudentModalVisible"
      preset="card"
      :title="`上传报告 - ${uploadStudentTarget?.studentName || ''}`"
      style="width: 480px"
      :mask-closable="false"
    >
      <NSpace vertical :size="12">
        <NText depth="2">
          将文件直接指派给该学生，无需文件名匹配。支持 .docx / .doc / .zip（zip 取内含主报告）。
        </NText>
        <NUpload
          v-model:file-list="uploadStudentFileList"
          :max="1"
          accept=".doc,.docx,.zip"
          :default-upload="false"
          @change="onUploadStudentFileChange"
        >
          <NButton>选择文件</NButton>
        </NUpload>
        <NText v-if="uploadStudentTarget?.gradeStatus === 'graded'" type="warning">
          该学生已批改，上传新文件会覆盖旧文件并自动重新批改。
        </NText>
      </NSpace>
      <template #footer>
        <NSpace justify="end">
          <NButton @click="uploadStudentModalVisible = false">取消</NButton>
          <NButton type="primary" :loading="uploadingStudent" @click="submitUploadStudent">确认上传</NButton>
        </NSpace>
      </template>
    </NModal>
  </NSpace>
</template>

<style scoped></style>
