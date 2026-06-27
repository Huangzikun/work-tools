import axios from 'axios';
import type { AxiosProgressEvent } from 'axios';
import { getServiceBaseURL } from '@/utils/service';
import { localStg } from '@/utils/storage';
import { request } from '../request';

const isHttpProxy = import.meta.env.DEV && import.meta.env.VITE_HTTP_PROXY === 'Y';
const { baseURL } = getServiceBaseURL(import.meta.env, isHttpProxy);

export interface LessonPlanGenerateProgress {
  onUploadProgress?: (e: AxiosProgressEvent) => void;
}

export interface GenerateParams {
  syllabus: File;
  syllabusName?: string;
  totalLessons: number;
  totalHours: number;
  batchSize: number;
  courseInfo: Api.LessonPlan.CourseInfo;
  teacherInfo: Api.LessonPlan.TeacherInfo;
  systemPrompt?: string;
  userPrompt?: string;
}

export function fetchLessonPlanDefaultPrompts() {
  return request<{ systemPrompt: string; userPrompt: string }>({
    url: '/lesson-plan/default-prompt',
    method: 'get'
  });
}

export async function generateLessonPlan(
  params: GenerateParams,
  progress?: LessonPlanGenerateProgress
): Promise<{ taskId: number; status: string }> {
  const form = new FormData();
  form.append('syllabus', params.syllabus);
  if (params.syllabusName) form.append('syllabusName', params.syllabusName);
  form.append('totalLessons', String(params.totalLessons));
  form.append('totalHours', String(params.totalHours));
  form.append('batchSize', String(params.batchSize));
  form.append('courseInfo', JSON.stringify(params.courseInfo));
  form.append('teacherInfo', JSON.stringify(params.teacherInfo));
  if (params.systemPrompt) form.append('systemPrompt', params.systemPrompt);
  if (params.userPrompt) form.append('userPrompt', params.userPrompt);

  const token = localStg.get('token');
  const resp = await axios.post<App.Service.Response<{ taskId: number; status: string }>>(
    `${baseURL}/lesson-plan/generate`,
    form,
    {
      headers: token ? { Authorization: `Bearer ${token}` } : {},
      onUploadProgress: progress?.onUploadProgress
    }
  );

  if (String(resp.data.code) !== import.meta.env.VITE_SERVICE_SUCCESS_CODE) {
    throw new Error(resp.data.msg || '创建任务失败');
  }
  return resp.data.data;
}

export function fetchLessonPlanTasks(page = 1, size = 20) {
  return request<Api.LessonPlan.TaskPage>({
    url: '/lesson-plan/tasks',
    method: 'get',
    params: { page, size }
  });
}

export function fetchLessonPlanTask(taskId: number) {
  return request<Api.LessonPlan.TaskSummary>({
    url: `/lesson-plan/tasks/${taskId}`,
    method: 'get'
  });
}

export function fetchLessonPlanProgress(taskId: number) {
  return request<Api.LessonPlan.Progress>({
    url: `/lesson-plan/tasks/${taskId}/progress`,
    method: 'get'
  });
}

export function regenerateLessonPlan(
  taskId: number,
  systemPrompt?: string,
  userPrompt?: string
) {
  const form = new FormData();
  if (systemPrompt) form.append('systemPrompt', systemPrompt);
  if (userPrompt) form.append('userPrompt', userPrompt);
  return request<{ taskId: number; status: string }>({
    url: `/lesson-plan/tasks/${taskId}/regenerate`,
    method: 'post',
    data: form,
    headers: { 'Content-Type': 'multipart/form-data' }
  });
}

export function deleteLessonPlanTask(taskId: number) {
  return request<{ taskId: number; deleted: boolean }>({
    url: `/lesson-plan/tasks/${taskId}`,
    method: 'delete'
  });
}

export async function downloadLessonPlan(taskId: number): Promise<void> {
  const token = localStg.get('token');
  const resp = await axios.get(`${baseURL}/lesson-plan/tasks/${taskId}/download`, {
    headers: token ? { Authorization: `Bearer ${token}` } : {},
    responseType: 'blob'
  });

  const ct = String(resp.headers['content-type'] ?? '');
  if (ct.includes('application/json')) {
    const text = await (resp.data as Blob).text();
    try {
      const json = JSON.parse(text) as { msg?: string };
      throw new Error(json.msg || '下载失败');
    } catch (err) {
      if (err instanceof Error && err.message !== '下载失败') throw err;
      throw new Error('下载失败');
    }
  }

  const cd = String(resp.headers['content-disposition'] ?? '');
  const m = cd.match(/filename\*=UTF-8''([^;]+)/i);
  const filename = m ? decodeURIComponent(m[1]) : '教案.docx';

  const url = URL.createObjectURL(resp.data as Blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
