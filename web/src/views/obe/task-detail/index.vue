<script setup lang="ts">
import { computed, h, onMounted, onUnmounted, ref, watch } from 'vue';
import { useRoute, useRouter } from 'vue-router';
import {
  NAlert,
  NButton,
  NCard,
  NCheckbox,
  NDataTable,
  NDatePicker,
  NForm,
  NFormItem,
  NInput,
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
  deleteObeTask,
  downloadObeExcel,
  downloadObeTaskAll,
  downloadObeZip,
  fetchObeProgress,
  fetchObeTaskDetail,
  fetchObeTaskTree,
  resolveObeAmbiguous,
  retryObeStudent,
  startObeGrade,
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
  teacherPrompt: '',
  signPicture: null as File | null,
  overwriteGraded: false
});
const gradeFormRef = ref();
const signPictureFileList = ref<UploadFileInfo[]>([]);

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
    teacherPrompt: '',
    signPicture: null,
    overwriteGraded: false
  };
  signPictureFileList.value = [];
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
  if (!gradeForm.value.teacherPrompt.trim()) {
    window.$message?.error('请填写教师要求');
    return;
  }
  if (!gradeForm.value.signPicture) {
    window.$message?.error('请上传签名图片');
    return;
  }

  const date = new Date(gradeForm.value.signDate * 1000);
  const signDateStr = `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;

  try {
    const result = await startObeGrade(taskId.value, {
      dirType: activeDirType.value,
      experimentLabel: activeExperiment.value,
      teacherName: gradeForm.value.teacherName.trim(),
      signDate: signDateStr,
      teacherPrompt: gradeForm.value.teacherPrompt.trim(),
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
let progressTimer: ReturnType<typeof setInterval> | null = null;

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
        window.$message?.success(
          `批改完成（成功 ${data.graded}/${data.total}，失败 ${data.failed}）`
        );
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

onUnmounted(stopProgressPolling);

// ============ 单学生重试 ============
async function handleRetry(studentPk: number) {
  window.$message?.info('正在重试...');
  const { error } = await retryObeStudent(taskId.value, studentPk);
  if (error) {
    window.$message?.error(errMsg(error, '重试失败'));
    return;
  }
  window.$message?.success('重试完成');
  await loadDetail();
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
    width: 100,
    render: row =>
      row.matched
        ? h(NButton, { size: 'small', tertiary: true, onClick: () => handleRetry(row.id) }, { default: () => '重试' })
        : null
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

      <NAlert
        v-if="progress && progress.status === 'running'"
        type="info"
        :show-icon="true"
        style="margin-bottom: 12px"
      >
        批改进度：{{ progress.graded + progress.failed }}/{{ progress.total }}
        （成功 {{ progress.graded }}，失败 {{ progress.failed }}），当前：
        {{ progress.currentStudent || '...' }}
      </NAlert>

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
      style="width: 560px"
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
        <NFormItem label="教师要求" required>
          <NInput
            v-model:value="gradeForm.teacherPrompt"
            type="textarea"
            :autosize="{ minRows: 4, maxRows: 10 }"
            placeholder="请填写本次实验的批改要求"
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
  </NSpace>
</template>

<style scoped></style>
