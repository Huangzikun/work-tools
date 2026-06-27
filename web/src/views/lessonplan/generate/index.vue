<script setup lang="ts">
import { computed, reactive, ref } from 'vue';
import type { UploadFileInfo } from 'naive-ui';
import { useRouter } from 'vue-router';
import { useFormRules, useNaiveForm } from '@/hooks/common/form';
import { generateLessonPlan } from '@/service/api/lessonPlan';
import { useAuthStore } from '@/store/modules/auth';
import { fetchLessonPlanDefaultSystemPrompt } from '@/service/api/lessonPlan';

defineOptions({
  name: 'LessonPlanGenerate'
});

const router = useRouter();
const authStore = useAuthStore();
const { formRef, validate, restoreValidation } = useNaiveForm();
const { defaultRequiredRule } = useFormRules();

const COURSE_TYPE_OPTIONS = ['专业教育课程', '通识教育课程', '协同拓展课程', '实践创新课程'].map(v => ({
  label: v,
  value: v
}));
const COURSE_NATURE_OPTIONS = ['必修', '选修'].map(v => ({ label: v, value: v }));

interface FormModel {
  totalLessons: number | null;
  batchSize: number;
  courseName: string;
  courseNameEn: string;
  credit: string;
  theoryHours: string;
  practiceHours: string;
  onlineHours: string;
  major: string;
  prereq: string;
  semester: string;
  teachingRoom: string;
  teacherName: string;
  unit: string;
  writeDate: string;
  courseType: string;
  courseNature: string;
  systemPrompt: string;
}

const today = new Date().toISOString().slice(0, 10);

const model = reactive<FormModel>({
  totalLessons: 4,
  batchSize: 2,
  courseName: '',
  courseNameEn: '',
  credit: '',
  theoryHours: '',
  practiceHours: '',
  onlineHours: '',
  major: '',
  prereq: '',
  semester: '',
  teachingRoom: '',
  teacherName: authStore.userInfo.userName || '',
  unit: '',
  writeDate: today,
  courseType: '专业教育课程',
  courseNature: '必修',
  systemPrompt: ''
});

const rules = computed<Record<string, App.Global.FormRule[]>>(() => ({
  totalLessons: [defaultRequiredRule],
  courseName: [defaultRequiredRule],
  teacherName: [defaultRequiredRule]
}));

const fileList = ref<UploadFileInfo[]>([]);
const syllabusFile = ref<File | null>(null);

function handleFileChange(data: { fileList: UploadFileInfo[] }) {
  fileList.value = data.fileList;
  const newest = data.fileList[data.fileList.length - 1];
  syllabusFile.value = newest?.file ?? null;
}

const submitting = ref(false);
const uploadPercent = ref(0);
const uploadText = ref('');

const advancedCollapsed = ref(true);

async function loadDefaultSystemPrompt() {
  const { data, error } = await fetchLessonPlanDefaultSystemPrompt();
  if (!error && data) {
    model.systemPrompt = data;
  }
}

loadDefaultSystemPrompt();

function buildCourseInfo(): Api.LessonPlan.CourseInfo {
  return {
    课程名称: model.courseName.trim(),
    英文名称: model.courseNameEn.trim(),
    学分: model.credit.trim(),
    理论学时: model.theoryHours.trim(),
    实践学时: model.practiceHours.trim(),
    线上学时: model.onlineHours.trim(),
    适用专业: model.major.trim(),
    先修课程: model.prereq.trim(),
    开课学期: model.semester.trim(),
    课程教研室: model.teachingRoom.trim()
  };
}

function buildTeacherInfo(): Api.LessonPlan.TeacherInfo {
  return {
    授课教师: model.teacherName.trim(),
    所属单位: model.unit.trim(),
    撰写日期: model.writeDate.trim(),
    课程类型: model.courseType,
    课程性质: model.courseNature
  };
}

async function handleSubmit() {
  await validate();
  if (!syllabusFile.value) {
    window.$message?.error('请上传教学大纲 docx');
    return;
  }
  if (!model.totalLessons || model.totalLessons < 1 || model.totalLessons > 200) {
    window.$message?.error('教案数必须在 1-200 之间');
    return;
  }

  submitting.value = true;
  uploadPercent.value = 0;
  uploadText.value = '上传中 0%';

  try {
    const result = await generateLessonPlan(
      {
        syllabus: syllabusFile.value,
        syllabusName: syllabusFile.value.name,
        totalLessons: model.totalLessons,
        batchSize: model.batchSize,
        courseInfo: buildCourseInfo(),
        teacherInfo: buildTeacherInfo(),
        systemPrompt: model.systemPrompt.trim() || undefined
      },
      {
        onUploadProgress: e => {
          if (e.total) {
            const percent = Math.round((e.loaded / e.total) * 100);
            uploadPercent.value = Math.min(percent, 100);
            uploadText.value = `上传中 ${percent}%`;
          }
        }
      }
    );

    window.$message?.success('任务已创建，后台正在生成');
    router.push({ name: 'lessonplan_tasks' });
    void result;
  } catch (err) {
    const e = err as { response?: { data?: { msg?: string } }; message?: string };
    window.$message?.error(e?.response?.data?.msg || e?.message || '创建任务失败');
  } finally {
    submitting.value = false;
    uploadPercent.value = 0;
    uploadText.value = '';
  }
}

function handleReset() {
  restoreValidation();
  model.courseName = '';
  model.courseNameEn = '';
  model.credit = '';
  model.theoryHours = '';
  model.practiceHours = '';
  model.onlineHours = '';
  model.major = '';
  model.prereq = '';
  model.semester = '';
  model.teachingRoom = '';
  model.teacherName = authStore.userInfo.userName || '';
  model.unit = '';
  model.writeDate = today;
  model.courseType = '专业教育课程';
  model.courseNature = '必修';
  model.totalLessons = 4;
  model.batchSize = 2;
  fileList.value = [];
  syllabusFile.value = null;
}
</script>

<template>
  <NSpace vertical :size="16">
    <NCard :bordered="false" class="card-wrapper">
      <NAlert type="info" :show-icon="true">
        上传教学大纲 docx，填写课程基本信息后由 AI 自动生成完整教案。「教案数」决定生成多少份教案，每份对应一次课（约 5 课时 / 200 分钟）。提交后可在任务列表查看进度。
      </NAlert>
    </NCard>

    <NCard :bordered="false" class="card-wrapper" title="上传教学大纲">
      <NUpload
        v-model:file-list="fileList"
        :max="1"
        accept=".doc,.docx"
        :default-upload="false"
        @change="handleFileChange"
      >
        <NUploadDragger>
          <NSpace vertical align="center" :size="8">
            <NIcon size="36" :depth="3">
              <svg viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z" />
              </svg>
            </NIcon>
            <NText>点击或拖拽上传教学大纲 (.doc / .docx)</NText>
            <NText depth="3" class="text-12px">AI 会读取大纲内容并按课时数自动分配章节</NText>
          </NSpace>
        </NUploadDragger>
      </NUpload>
    </NCard>

    <NCard :bordered="false" class="card-wrapper" title="生成参数">
      <NForm
        ref="formRef"
        :model="model"
        :rules="rules"
        label-placement="left"
        label-width="140"
        require-mark-placement="right-hanging"
      >
        <div class="form-grid-2">
          <NFormItem label="教案数" path="totalLessons">
            <NInputNumber v-model:value="model.totalLessons" :min="1" :max="200" class="w-full" />
            <NText depth="3" class="ml-8px text-12px whitespace-nowrap">要生成多少份教案</NText>
          </NFormItem>
          <NFormItem label="批量大小">
            <div class="flex items-center w-full">
              <NInputNumber v-model:value="model.batchSize" :min="1" :max="5" class="flex-1" />
              <NText depth="3" class="ml-8px text-12px whitespace-nowrap">每批生成数</NText>
            </div>
          </NFormItem>
        </div>
      </NForm>
    </NCard>

    <NCard :bordered="false" class="card-wrapper" title="课程基本信息（首页）">
      <NForm label-placement="left" label-width="140">
        <div class="form-grid-2">
          <NFormItem label="课程名称" required>
            <NInput v-model:value="model.courseName" placeholder="如：数据结构" />
          </NFormItem>
          <NFormItem label="英文名称">
            <NInput v-model:value="model.courseNameEn" placeholder="如：Data Structures" />
          </NFormItem>
          <NFormItem label="学分">
            <NInput v-model:value="model.credit" placeholder="如：4" />
          </NFormItem>
          <NFormItem label="理论学时">
            <NInput v-model:value="model.theoryHours" placeholder="如：64" />
          </NFormItem>
          <NFormItem label="实践学时">
            <NInput v-model:value="model.practiceHours" placeholder="如：0" />
          </NFormItem>
          <NFormItem label="线上学时">
            <NInput v-model:value="model.onlineHours" placeholder="如：0" />
          </NFormItem>
          <NFormItem label="适用专业">
            <NInput v-model:value="model.major" placeholder="如：计算机科学与技术" />
          </NFormItem>
          <NFormItem label="先修课程">
            <NInput v-model:value="model.prereq" placeholder="如：程序设计基础" />
          </NFormItem>
          <NFormItem label="开课学期">
            <NInput v-model:value="model.semester" placeholder="如：第4学期" />
          </NFormItem>
          <NFormItem label="课程教研室">
            <NInput v-model:value="model.teachingRoom" placeholder="如：软件工程" />
          </NFormItem>
        </div>
      </NForm>
    </NCard>

    <NCard :bordered="false" class="card-wrapper" title="教师信息（首页）">
      <NForm label-placement="left" label-width="140">
        <div class="form-grid-2">
          <NFormItem label="授课教师" required>
            <NInput v-model:value="model.teacherName" placeholder="教师姓名" />
          </NFormItem>
          <NFormItem label="所属单位">
            <NInput v-model:value="model.unit" placeholder="如：信息工程学院" />
          </NFormItem>
          <NFormItem label="撰写日期">
            <NInput v-model:value="model.writeDate" placeholder="如：2026-06-26" />
          </NFormItem>
          <NFormItem label="课程类型">
            <NSelect v-model:value="model.courseType" :options="COURSE_TYPE_OPTIONS" />
          </NFormItem>
          <NFormItem label="课程性质">
            <NSelect v-model:value="model.courseNature" :options="COURSE_NATURE_OPTIONS" />
          </NFormItem>
        </div>
      </NForm>
    </NCard>

    <NCard :bordered="false" class="card-wrapper">
      <NCollapse>
        <NCollapseItem title="高级设置（自定义 AI 提示词）" name="advanced">
          <NForm label-placement="top">
            <NFormItem label="System Prompt">
              <NInput
                v-model:value="model.systemPrompt"
                type="textarea"
                :autosize="{ minRows: 8, maxRows: 20 }"
                placeholder="留空将使用内置默认提示词"
              />
            </NFormItem>
          </NForm>
        </NCollapseItem>
      </NCollapse>
    </NCard>

    <NCard :bordered="false" class="card-wrapper">
      <NSpace>
        <NButton type="primary" :loading="submitting" @click="handleSubmit">开始生成</NButton>
        <NButton :disabled="submitting" @click="handleReset">重置</NButton>
      </NSpace>

      <div v-if="submitting && uploadText" class="mt-12px">
        <NProgress type="line" :percentage="uploadPercent" :show-indicator="false" />
        <NText depth="3" class="text-12px">{{ uploadText }}</NText>
      </div>
    </NCard>
  </NSpace>
</template>

<style scoped>
.form-grid-2 {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 24px;
}
</style>
