import axios from 'axios';
import type { AxiosProgressEvent } from 'axios';
import { getServiceBaseURL } from '@/utils/service';
import { localStg } from '@/utils/storage';
import { request } from '../request';

const isHttpProxy = import.meta.env.DEV && import.meta.env.VITE_HTTP_PROXY === 'Y';
const { baseURL } = getServiceBaseURL(import.meta.env, isHttpProxy);

// ============ 目录生成（POST /obe/mkdir） ============

export interface ObeMkdirProgress {
  onUploadProgress?: (e: AxiosProgressEvent) => void;
}

/**
 * 触发后端目录生成并持久化。返回 taskId 和目录树。
 *
 * 注意：不用项目的 `request` 封装，因为它会从 response.data.data 取值，
 * 但我们需要在上传过程中拿到 progress 事件，所以直接用 axios。
 * 后端响应是 JSON，但 axios transform 会原样返回 response.data。
 */
export async function fetchObeMkdir(
  params: Api.Obe.MkdirRequest,
  progress?: ObeMkdirProgress
): Promise<Api.Obe.MkdirResult> {
  const form = new FormData();
  form.append('className', params.className);
  form.append('courseName', params.courseName);
  form.append('teacherName', params.teacherName);
  form.append('fixedDirTypes', JSON.stringify(params.fixedDirTypes));
  form.append('studentDirTypes', JSON.stringify(params.studentDirTypes));
  form.append('roster', params.roster);

  const token = localStg.get('token');
  const resp = await axios.post<App.Service.Response<Api.Obe.MkdirResult>>(
    `${baseURL}/obe/mkdir`,
    form,
    {
      headers: token ? { Authorization: `Bearer ${token}` } : {},
      onUploadProgress: progress?.onUploadProgress
    }
  );

  if (String(resp.data.code) !== import.meta.env.VITE_SERVICE_SUCCESS_CODE) {
    throw new Error(resp.data.msg || '生成失败');
  }
  return resp.data.data;
}

// ============ 任务列表 ============

export function fetchObeTasks(page = 1, size = 20) {
  return request<Api.Obe.TaskPage>({
    url: '/obe/tasks',
    method: 'get',
    params: { page, size }
  });
}

export function fetchObeTaskDetail(taskId: number) {
  return request<Api.Obe.TaskDetail>({
    url: `/obe/tasks/${taskId}`,
    method: 'get',
    // 加时间戳防缓存：单学生批改轮询时需要每次拿最新状态
    params: { _t: Date.now() }
  });
}

export function fetchObeTaskTree(taskId: number) {
  return request<{ tree: Api.Obe.TreeNode }>({
    url: `/obe/tasks/${taskId}/tree`,
    method: 'get'
  });
}

export function deleteObeTask(taskId: number) {
  return request<{ taskId: number; deleted: boolean }>({
    url: `/obe/tasks/${taskId}`,
    method: 'delete'
  });
}

// ============ 上传学生文件 ============

export interface ObeUploadProgress {
  onUploadProgress?: (e: AxiosProgressEvent) => void;
}

export async function uploadObeStudentFiles(
  taskId: number,
  dirType: string,
  files: File[],
  progress?: ObeUploadProgress
): Promise<Api.Obe.UploadResult> {
  const form = new FormData();
  form.append('dirType', dirType);
  files.forEach(f => form.append('files', f));

  const token = localStg.get('token');
  const resp = await axios.post<App.Service.Response<Api.Obe.UploadResult>>(
    `${baseURL}/obe/tasks/${taskId}/upload`,
    form,
    {
      headers: token ? { Authorization: `Bearer ${token}` } : {},
      onUploadProgress: progress?.onUploadProgress
    }
  );

  if (String(resp.data.code) !== import.meta.env.VITE_SERVICE_SUCCESS_CODE) {
    throw new Error(resp.data.msg || '上传失败');
  }
  return resp.data.data;
}

export function resolveObeAmbiguous(
  taskId: number,
  dirType: string,
  experimentLabel: string,
  resolutions: Api.Obe.AmbiguousResolution[]
) {
  return request<{ resolved: Api.Obe.MatchedItem[]; failed: Api.Obe.AmbiguousResolution[] }>({
    url: `/obe/tasks/${taskId}/upload/resolve`,
    method: 'post',
    data: { dirType, experimentLabel, resolutions }
  });
}

// ============ 触发批改 ============

export interface ObeGradeProgress {
  onUploadProgress?: (e: AxiosProgressEvent) => void;
}

export async function startObeGrade(
  taskId: number,
  params: Api.Obe.GradeRequest,
  progress?: ObeGradeProgress
): Promise<Api.Obe.GradeResult> {
  const form = new FormData();
  form.append('dirType', params.dirType);
  form.append('experimentLabel', params.experimentLabel);
  form.append('teacherName', params.teacherName);
  form.append('signDate', params.signDate);
  form.append('teacherPrompt', params.teacherPrompt);
  if (params.systemPrompt) form.append('systemPrompt', params.systemPrompt);
  if (params.rubricDimensions && params.rubricDimensions.length) {
    form.append(
      'rubricDimensions',
      JSON.stringify(
        params.rubricDimensions.map(d => ({ name: d.name, max_score: d.maxScore, criteria: d.criteria }))
      )
    );
  }
  if (params.scoreLevels && params.scoreLevels.length) {
    form.append('scoreLevels', JSON.stringify(params.scoreLevels));
  }
  form.append('signPicture', params.signPicture);
  form.append('skipGraded', params.skipGraded === false ? 'false' : 'true');

  const token = localStg.get('token');
  const resp = await axios.post<App.Service.Response<Api.Obe.GradeResult>>(
    `${baseURL}/obe/tasks/${taskId}/grade`,
    form,
    {
      headers: token ? { Authorization: `Bearer ${token}` } : {},
      onUploadProgress: progress?.onUploadProgress
    }
  );

  if (String(resp.data.code) !== import.meta.env.VITE_SERVICE_SUCCESS_CODE) {
    throw new Error(resp.data.msg || '触发批改失败');
  }
  return resp.data.data;
}

export function fetchObeProgress(
  taskId: number,
  dirType: string,
  experimentLabel?: string,
  jobId?: number
) {
  const params: Record<string, string | number> = { dirType };
  if (experimentLabel) params.experimentLabel = experimentLabel;
  if (jobId) params.jobId = jobId;
  return request<Api.Obe.Progress>({
    url: `/obe/tasks/${taskId}/progress`,
    method: 'get',
    params
  });
}

export function retryObeStudent(taskId: number, studentPk: number) {
  return request<Api.Obe.Student>({
    url: `/obe/tasks/${taskId}/students/${studentPk}/retry`,
    method: 'post'
  });
}

// ============ 评分标准（按课程保存/载入） ============

export function fetchObeRubric(courseName: string) {
  return request<Api.Obe.RubricSnapshot | null>({
    url: '/obe/rubric',
    method: 'get',
    params: { courseName }
  });
}

export function saveObeRubric(payload: {
  courseName: string;
  dimensions: Api.Obe.RubricDimension[];
  freeText: string;
  scoreLevels?: number[] | null;
}) {
  return request<Api.Obe.RubricSnapshot>({ url: '/obe/rubric', method: 'post', data: payload });
}

export function generateObeRubric(experimentContent: string) {
  return request<{
    dimensions: Array<{ name: string; max_score: number; criteria: string }>;
    freeText: string;
  }>({ url: '/obe/rubric/generate', method: 'post', data: { experimentContent } });
}

export function cancelObeGrade(taskId: number, jobId: number) {
  return request<{ jobId: number; status: string }>({
    url: `/obe/tasks/${taskId}/jobs/${jobId}/cancel`,
    method: 'post'
  });
}

// ============ 下载 ============

export async function downloadObeZip(
  taskId: number,
  dirType: string,
  experimentLabel?: string
): Promise<void> {
  const token = localStg.get('token');
  const resp = await axios.get(`${baseURL}/obe/tasks/${taskId}/download/zip`, {
    params: experimentLabel ? { dirType, experimentLabel } : { dirType },
    headers: token ? { Authorization: `Bearer ${token}` } : {},
    responseType: 'blob'
  });

  // 后端错误时可能返回 JSON
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
  const filename = m ? decodeURIComponent(m[1]) : `${dirType}.zip`;

  const url = URL.createObjectURL(resp.data as Blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

export async function downloadObeExcel(
  taskId: number,
  dirType: string,
  experimentLabel: string,
  jobId?: number
): Promise<void> {
  const token = localStg.get('token');
  const params: Record<string, string | number> = { dirType, experimentLabel };
  if (jobId) params.jobId = jobId;
  const resp = await axios.get(`${baseURL}/obe/tasks/${taskId}/download/excel`, {
    params,
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
  const filename = m ? decodeURIComponent(m[1]) : `${dirType}_成绩.xlsx`;

  const url = URL.createObjectURL(resp.data as Blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

export async function downloadObeTaskAll(taskId: number): Promise<void> {
  const token = localStg.get('token');
  const resp = await axios.get(`${baseURL}/obe/tasks/${taskId}/download/all`, {
    headers: token ? { Authorization: `Bearer ${token}` } : {},
    responseType: 'blob'
  });

  // 后端错误时可能返回 JSON
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
  const filename = m ? decodeURIComponent(m[1]) : `${taskId}.zip`;

  const url = URL.createObjectURL(resp.data as Blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ============ 单学生上传 / 下载 ============

export interface ObeStudentUploadResult {
  student: Api.Obe.Student;
  fileName: string;
  resetPreviousGrading: boolean;
}

/** 单学生上传：把一个文件直接指派给指定学生（无需文件名匹配）。 */
export async function uploadObeStudentFile(
  taskId: number,
  studentPk: number,
  file: File,
  progress?: ObeUploadProgress
): Promise<ObeStudentUploadResult> {
  const form = new FormData();
  form.append('file', file);
  const token = localStg.get('token');
  const resp = await axios.post<App.Service.Response<ObeStudentUploadResult>>(
    `${baseURL}/obe/tasks/${taskId}/students/${studentPk}/upload`,
    form,
    {
      headers: token ? { Authorization: `Bearer ${token}` } : {},
      onUploadProgress: progress?.onUploadProgress
    }
  );

  if (String(resp.data.code) !== import.meta.env.VITE_SERVICE_SUCCESS_CODE) {
    throw new Error(resp.data.msg || '上传失败');
  }
  return resp.data.data;
}

/** 单学生下载：优先批改后文件，回退原始上传件。 */
export async function downloadObeStudentFile(taskId: number, studentPk: number): Promise<void> {
  const token = localStg.get('token');
  const resp = await axios.get(`${baseURL}/obe/tasks/${taskId}/students/${studentPk}/download`, {
    headers: token ? { Authorization: `Bearer ${token}` } : {},
    responseType: 'blob'
  });

  // 后端错误时可能返回 JSON
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
  const filename = m ? decodeURIComponent(m[1]) : `student_${studentPk}.docx`;

  const url = URL.createObjectURL(resp.data as Blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
