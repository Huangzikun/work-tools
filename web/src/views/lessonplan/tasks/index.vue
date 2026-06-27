<script setup lang="ts">
import { computed, h, onMounted, onUnmounted, ref } from 'vue';
import { useRouter } from 'vue-router';
import { NButton, NPopconfirm, NProgress, NSpace, NTag, NText } from 'naive-ui';
import type { DataTableColumns } from 'naive-ui';
import {
  deleteLessonPlanTask,
  downloadLessonPlan,
  fetchLessonPlanProgress,
  fetchLessonPlanTasks,
  regenerateLessonPlan
} from '@/service/api/lessonPlan';

defineOptions({ name: 'LessonPlanTasks' });

const router = useRouter();

const loading = ref(false);
const data = ref<Api.LessonPlan.TaskSummary[]>([]);
const page = ref(1);
const size = ref(20);
const total = ref(0);

function errMsg(err: any, fallback: string): string {
  return err?.response?.data?.msg || err?.msg || err?.message || fallback;
}

async function loadTasks() {
  loading.value = true;
  try {
    const { data: respData, error } = await fetchLessonPlanTasks(page.value, size.value);
    if (error) {
      window.$message?.error(errMsg(error, '加载任务失败'));
      return;
    }
    data.value = respData?.list ?? [];
    total.value = respData?.total ?? 0;
  } finally {
    loading.value = false;
  }
}

function statusTag(s: string): { label: string; type: 'default' | 'info' | 'success' | 'warning' | 'error' } {
  switch (s) {
    case 'pending':
      return { label: '等待中', type: 'default' };
    case 'parsing':
      return { label: '解析中', type: 'info' };
    case 'generating':
      return { label: '生成中', type: 'info' };
    case 'building':
      return { label: '构建文档', type: 'info' };
    case 'completed':
      return { label: '已完成', type: 'success' };
    case 'failed':
      return { label: '失败', type: 'error' };
    default:
      return { label: s, type: 'default' };
  }
}

const ACTIVE_STATUSES = new Set(['pending', 'parsing', 'generating', 'building']);

function progressPercent(row: Api.LessonPlan.TaskSummary): number {
  if (!row.progressTotal) return 0;
  return Math.min(100, Math.round((row.progressDone / row.progressTotal) * 100));
}

let pollTimer: number | null = null;

function startPolling() {
  stopPolling();
  pollTimer = window.setInterval(async () => {
    const activeIds = data.value.filter(t => ACTIVE_STATUSES.has(t.status)).map(t => t.id);
    if (activeIds.length === 0) return;
    await Promise.all(
      activeIds.map(async id => {
        const { data: prog, error } = await fetchLessonPlanProgress(id);
        if (error || !prog) return;
        const target = data.value.find(t => t.id === id);
        if (!target) return;
        target.status = (prog.status as Api.LessonPlan.TaskStatus) || target.status;
        target.progressDone = prog.done ?? target.progressDone;
        target.progressTotal = prog.total ?? target.progressTotal;
        target.progressLabel = prog.label ?? target.progressLabel;
        if (target.status === 'completed' || target.status === 'failed') {
          // 完成后刷新一次拿 output_file / error_summary
          await loadTasks();
        }
      })
    );
  }, 3000);
}

function stopPolling() {
  if (pollTimer !== null) {
    window.clearInterval(pollTimer);
    pollTimer = null;
  }
}

async function handleDelete(taskId: number) {
  const { error } = await deleteLessonPlanTask(taskId);
  if (error) {
    window.$message?.error(errMsg(error, '删除失败'));
    return;
  }
  window.$message?.success('已删除');
  await loadTasks();
}

async function handleRegenerate(taskId: number) {
  const { error } = await regenerateLessonPlan(taskId);
  if (error) {
    window.$message?.error(errMsg(error, '重新生成失败'));
    return;
  }
  window.$message?.success('已开始重新生成');
  await loadTasks();
}

async function handleDownload(taskId: number) {
  try {
    await downloadLessonPlan(taskId);
  } catch (err) {
    window.$message?.error(errMsg(err, '下载失败'));
  }
}

const columns = computed<DataTableColumns<Api.LessonPlan.TaskSummary>>(() => [
  { title: 'ID', key: 'id', width: 60 },
  { title: '大纲文件', key: 'syllabusName', minWidth: 180, ellipsis: { tooltip: true } },
  {
    title: '课程名',
    key: 'courseName',
    minWidth: 120,
    render: row => row.courseInfo?.课程名称 || '-'
  },
  { title: '教案数', key: 'totalLessons', width: 80, align: 'center' },
  { title: '教师', key: 'teacherName', width: 90, render: row => row.teacherInfo?.授课教师 || '-' },
  {
    title: '状态',
    key: 'status',
    width: 110,
    render: row => {
      const s = statusTag(row.status);
      return h(NTag, { type: s.type, size: 'small', round: true }, { default: () => s.label });
    }
  },
  {
    title: '进度',
    key: 'progress',
    width: 160,
    render: row => {
      const percent = progressPercent(row);
      return h(
        'div',
        { class: 'flex flex-col' },
        {
          default: () => [
            h(NProgress, {
              type: 'line',
              percentage: percent,
              showIndicator: false,
              status: row.status === 'failed' ? 'error' : row.status === 'completed' ? 'success' : 'default'
            }),
            h(
              NText,
              { depth: 3, class: 'text-12px' },
              { default: () => row.progressLabel || `${row.progressDone}/${row.progressTotal}` }
            )
          ]
        }
      );
    }
  },
  {
    title: '创建时间',
    key: 'createdAt',
    width: 180,
    render: row => (row.createdAt ? new Date(row.createdAt).toLocaleString('zh-CN') : '-')
  },
  {
    title: '操作',
    key: 'actions',
    width: 240,
    fixed: 'right',
    render: row =>
      h(NSpace, { size: 8 }, {
        default: () => [
          row.status === 'completed'
            ? h(
                NButton,
                { size: 'small', type: 'primary', onClick: () => handleDownload(row.id) },
                { default: () => '下载' }
              )
            : null,
          row.status === 'failed' || row.status === 'completed'
            ? h(
                NPopconfirm,
                { onPositiveClick: () => handleRegenerate(row.id) },
                {
                  trigger: () =>
                    h(NButton, { size: 'small', type: 'warning', tertiary: true }, { default: () => '重新生成' }),
                  default: () => '将基于原大纲和参数重新生成，旧文件会备份到 retry/ 目录。'
                }
              )
            : null,
          !ACTIVE_STATUSES.has(row.status)
            ? h(
                NPopconfirm,
                { onPositiveClick: () => handleDelete(row.id) },
                {
                  trigger: () =>
                    h(NButton, { size: 'small', type: 'error', tertiary: true }, { default: () => '删除' }),
                  default: () => '确认删除该任务？大纲、输出文件和重试备份将被清除。'
                }
              )
            : null
        ]
      })
  }
]);

function handlePageChange(p: number) {
  page.value = p;
  loadTasks();
}

onMounted(async () => {
  await loadTasks();
  startPolling();
});

onUnmounted(stopPolling);
</script>

<template>
  <NSpace vertical :size="16">
    <NCard :bordered="false" class="card-wrapper">
      <NSpace justify="space-between" align="center">
        <div class="text-16px font-medium">教案生成任务列表</div>
        <NButton type="primary" @click="router.push({ name: 'lessonplan_generate' })">新建生成</NButton>
      </NSpace>
    </NCard>

    <NCard :bordered="false" class="card-wrapper">
      <NDataTable
        :columns="columns"
        :data="data"
        :loading="loading"
        :pagination="{
          page,
          pageSize: size,
          itemCount: total,
          showSizePicker: false,
          onChange: handlePageChange
        }"
        remote
        :scroll-x="1200"
      />
    </NCard>
  </NSpace>
</template>

<style scoped></style>
