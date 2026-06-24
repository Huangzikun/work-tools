<script setup lang="ts">
import { h, onMounted, ref } from 'vue';
import { useRouter } from 'vue-router';
import { NButton, NPopconfirm, NSpace, NTag } from 'naive-ui';
import type { DataTableColumns } from 'naive-ui';
import { deleteObeTask, fetchObeTasks } from '@/service/api/obe';

defineOptions({ name: 'ObeTasks' });

const router = useRouter();

const loading = ref(false);
const data = ref<Api.Obe.TaskSummary[]>([]);
const page = ref(1);
const size = ref(20);
const total = ref(0);

function errMsg(err: any, fallback: string): string {
  return err?.response?.data?.msg || err?.msg || err?.message || fallback;
}

async function loadTasks() {
  loading.value = true;
  try {
    const { data: respData, error } = await fetchObeTasks(page.value, size.value);
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

async function handleDelete(taskId: number) {
  const { error } = await deleteObeTask(taskId);
  if (error) {
    window.$message?.error(errMsg(error, '删除失败'));
    return;
  }
  window.$message?.success('已删除');
  await loadTasks();
}

function goDetail(taskId: number) {
  router.push({ name: 'obe_task-detail', query: { taskId: String(taskId) } });
}

function statusLabel(s: string): { label: string; type: 'default' | 'info' | 'success' | 'warning' } {
  switch (s) {
    case 'created':
      return { label: '已创建', type: 'default' };
    case 'uploading':
      return { label: '上传中', type: 'info' };
    case 'ready':
      return { label: '待批改', type: 'info' };
    case 'grading':
      return { label: '批改中', type: 'warning' };
    case 'graded':
      return { label: '已批改', type: 'success' };
    default:
      return { label: s, type: 'default' };
  }
}

const columns: DataTableColumns<Api.Obe.TaskSummary> = [
  { title: 'ID', key: 'id', width: 60 },
  { title: '班级', key: 'className', minWidth: 180 },
  { title: '课程', key: 'courseName', minWidth: 120 },
  { title: '教师', key: 'teacherName', width: 100 },
  { title: '学生数', key: 'studentCount', width: 80, align: 'center' },
  {
    title: '考核目录',
    key: 'studentDirTypes',
    minWidth: 200,
    render: row => row.studentDirTypes.join('、')
  },
  {
    title: '状态',
    key: 'status',
    width: 100,
    render: row => {
      const s = statusLabel(row.status);
      return h(NTag, { type: s.type, size: 'small', round: true }, { default: () => s.label });
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
    width: 180,
    fixed: 'right',
    render: row =>
      h(NSpace, { size: 8 }, {
        default: () => [
          h(NButton, { size: 'small', type: 'primary', onClick: () => goDetail(row.id) }, { default: () => '查看' }),
          h(
            NPopconfirm,
            { onPositiveClick: () => handleDelete(row.id) },
            {
              trigger: () => h(NButton, { size: 'small', type: 'error', tertiary: true }, { default: () => '删除' }),
              default: () => '确认删除该任务？所有上传文件与批改记录将被清除。'
            }
          )
        ]
      })
  }
];

function handlePageChange(p: number) {
  page.value = p;
  loadTasks();
}

onMounted(loadTasks);
</script>

<template>
  <NSpace vertical :size="16">
    <NCard :bordered="false" class="card-wrapper">
      <NSpace justify="space-between" align="center">
        <div class="text-16px font-medium">OBE 任务列表</div>
        <NButton type="primary" @click="router.push({ name: 'obe_mkdir' })">新建任务</NButton>
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
