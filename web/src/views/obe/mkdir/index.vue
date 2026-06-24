<script setup lang="ts">
import { computed, reactive, ref } from 'vue';
import type { UploadFileInfo } from 'naive-ui';
import { useRouter } from 'vue-router';
import { useFormRules, useNaiveForm } from '@/hooks/common/form';
import { fetchObeMkdir } from '@/service/api/obe';
import { $t } from '@/locales';

defineOptions({
  name: 'ObeMkdir'
});

const router = useRouter();
const { formRef, validate, restoreValidation } = useNaiveForm();
const { defaultRequiredRule } = useFormRules();

interface FormModel {
  className: string;
  courseName: string;
  teacherName: string;
  fixedDirTypes: string[];
  studentDirTypes: string[];
}

const model = reactive<FormModel>({
  className: '',
  courseName: '',
  teacherName: '',
  fixedDirTypes: ['教学课件', '教学教案'],
  studentDirTypes: ['课程考核', '实验实训报告']
});

const presetStudentDirOptions = computed(() => [
  { label: $t('page.obeMkdir.presetStudentDir.courseAssessment'), value: '课程考核' },
  { label: $t('page.obeMkdir.presetStudentDir.labReport'), value: '实验实训报告' },
  { label: $t('page.obeMkdir.presetStudentDir.finalExam'), value: '期末试卷' },
  { label: $t('page.obeMkdir.presetStudentDir.dailyHomework'), value: '平时作业' },
  { label: $t('page.obeMkdir.presetStudentDir.courseDesign'), value: '课程设计' }
]);

const fixedDirOptions = computed(() => [
  { label: '教学课件', value: '教学课件' },
  { label: '教学教案', value: '教学教案' }
]);

const rules = computed<Record<keyof FormModel, App.Global.FormRule[]>>(() => ({
  className: [defaultRequiredRule],
  courseName: [defaultRequiredRule],
  teacherName: [defaultRequiredRule],
  fixedDirTypes: [],
  studentDirTypes: [
    {
      validator: (_rule, value: string[]) => {
        if (!value || value.length === 0) {
          if (model.fixedDirTypes.length === 0) {
            return new Error($t('page.obeMkdir.bothDirTypesEmpty'));
          }
        }
        return true;
      },
      trigger: 'change'
    }
  ]
}));

// 文件上传（不走自动上传，本地持有 File 对象）
const fileList = ref<UploadFileInfo[]>([]);
const rosterFile = ref<File | null>(null);

function handleFileChange(data: { fileList: UploadFileInfo[] }) {
  fileList.value = data.fileList;
  const newest = data.fileList[data.fileList.length - 1];
  rosterFile.value = newest?.file ?? null;
}

// 提交
const submitting = ref(false);
const progress = ref(0);
const progressText = ref('');

async function handleSubmit() {
  await validate();
  if (!rosterFile.value) {
    window.$message?.error($t('page.obeMkdir.rosterRequired'));
    return;
  }
  if (model.studentDirTypes.length === 0 && model.fixedDirTypes.length === 0) {
    window.$message?.error($t('page.obeMkdir.bothDirTypesEmpty'));
    return;
  }

  submitting.value = true;
  progress.value = 0;
  progressText.value = $t('page.obeMkdir.uploadProgress', { percent: 0 });

  try {
    const result = await fetchObeMkdir(
      {
        className: model.className.trim(),
        courseName: model.courseName.trim(),
        teacherName: model.teacherName.trim(),
        fixedDirTypes: model.fixedDirTypes,
        studentDirTypes: model.studentDirTypes,
        roster: rosterFile.value
      },
      {
        onUploadProgress: e => {
          if (e.total) {
            const percent = Math.round((e.loaded / e.total) * 100);
            progress.value = Math.min(percent, 100);
            progressText.value = $t('page.obeMkdir.uploadProgress', { percent });
          }
        }
      }
    );

    window.$message?.success($t('page.obeMkdir.success'));
    // 跳转到任务详情页（展示目录结构 + 提供批改入口）
    router.push({ name: 'obe_task-detail', query: { taskId: String(result.taskId) } });
  } catch (err) {
    window.$message?.error(err instanceof Error ? err.message : '生成失败');
  } finally {
    submitting.value = false;
    progress.value = 0;
    progressText.value = '';
  }
}

function handleReset() {
  restoreValidation();
  model.className = '';
  model.courseName = '';
  model.teacherName = '';
  model.fixedDirTypes = ['教学课件', '教学教案'];
  model.studentDirTypes = ['课程考核', '实验实训报告'];
  fileList.value = [];
  rosterFile.value = null;
}
</script>

<template>
  <NSpace vertical :size="16">
    <NCard :bordered="false" class="card-wrapper">
      <NAlert type="info" :show-icon="true">
        {{ $t('page.obeMkdir.subtitle') }}
      </NAlert>
    </NCard>

    <NCard :bordered="false" class="card-wrapper" :title="$t('page.obeMkdir.title')">
      <NForm
        ref="formRef"
        :model="model"
        :rules="rules"
        label-placement="left"
        label-width="220"
        require-mark-placement="right-hanging"
      >
        <NFormItem :label="$t('page.obeMkdir.classNameLabel')" path="className">
          <NInput
            v-model:value="model.className"
            :placeholder="$t('page.obeMkdir.classNamePlaceholder')"
            clearable
          />
        </NFormItem>

        <NFormItem :label="$t('page.obeMkdir.courseNameLabel')" path="courseName">
          <NInput
            v-model:value="model.courseName"
            :placeholder="$t('page.obeMkdir.courseNamePlaceholder')"
            clearable
          />
        </NFormItem>

        <NFormItem :label="$t('page.obeMkdir.teacherNameLabel')" path="teacherName">
          <NInput
            v-model:value="model.teacherName"
            :placeholder="$t('page.obeMkdir.teacherNamePlaceholder')"
            clearable
          />
        </NFormItem>

        <NFormItem :label="$t('page.obeMkdir.fixedDirTypesLabel')" path="fixedDirTypes">
          <NCheckboxGroup v-model:value="model.fixedDirTypes">
            <NSpace>
              <NCheckbox
                v-for="opt in fixedDirOptions"
                :key="opt.value"
                :value="opt.value"
                :label="opt.label"
              />
            </NSpace>
          </NCheckboxGroup>
        </NFormItem>

        <NFormItem :label="$t('page.obeMkdir.studentDirTypesLabel')" path="studentDirTypes">
          <NSelect
            v-model:value="model.studentDirTypes"
            multiple
            filterable
            tag
            :options="presetStudentDirOptions"
            :placeholder="$t('page.obeMkdir.studentDirTypesPlaceholder')"
          />
        </NFormItem>

        <NFormItem :label="$t('page.obeMkdir.rosterLabel')" path="roster">
          <NSpace vertical :size="8" class="w-full">
            <NUpload
              v-model:file-list="fileList"
              :max="1"
              accept=".xls,.xlsx,.html,.htm"
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
                  <NText>{{ $t('page.obeMkdir.rosterDragHint') }}</NText>
                </NSpace>
              </NUploadDragger>
            </NUpload>
            <NText depth="3" class="text-12px">{{ $t('page.obeMkdir.rosterHint') }}</NText>
          </NSpace>
        </NFormItem>

        <NFormItem label=" ">
          <NSpace>
            <NButton type="primary" :loading="submitting" @click="handleSubmit">
              {{ $t('page.obeMkdir.submit') }}
            </NButton>
            <NButton :disabled="submitting" @click="handleReset">
              {{ $t('page.obeMkdir.reset') }}
            </NButton>
          </NSpace>
        </NFormItem>

        <NFormItem v-if="submitting && progressText" label=" ">
          <NSpace vertical :size="4" class="w-full">
            <NProgress type="line" :percentage="progress" :show-indicator="false" />
            <NText depth="3" class="text-12px">{{ progressText }}</NText>
          </NSpace>
        </NFormItem>
      </NForm>
    </NCard>
  </NSpace>
</template>

<style scoped></style>
