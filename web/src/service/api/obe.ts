import axios from 'axios';
import type { AxiosProgressEvent } from 'axios';
import { getServiceBaseURL } from '@/utils/service';
import { localStg } from '@/utils/storage';

const isHttpProxy = import.meta.env.DEV && import.meta.env.VITE_HTTP_PROXY === 'Y';
const { baseURL } = getServiceBaseURL(import.meta.env, isHttpProxy);

export interface ObeMkdirProgress {
  onUploadProgress?: (e: AxiosProgressEvent) => void;
  onDownloadProgress?: (e: AxiosProgressEvent) => void;
}

/**
 * 触发后端目录生成并下载 ZIP。
 *
 * 注意：不能用项目里的 `request` 封装（其 transform 会从 response.data.data 取值，
 * 对 Blob 响应不适用），所以这里直接用 axios + responseType: 'blob'。
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
  const resp = await axios.post(`${baseURL}/obe/mkdir`, form, {
    headers: token ? { Authorization: `Bearer ${token}` } : {},
    responseType: 'blob',
    onUploadProgress: progress?.onUploadProgress,
    onDownloadProgress: progress?.onDownloadProgress
  });

  // 兜底：若后端返回 JSON 错误（content-type 不是 zip），Blob 里其实是 JSON
  const ct = String(resp.headers['content-type'] ?? '');
  if (ct.includes('application/json')) {
    const text = await (resp.data as Blob).text();
    try {
      const json = JSON.parse(text) as { msg?: string };
      throw new Error(json.msg || '生成失败');
    } catch (err) {
      if (err instanceof Error && err.message !== '生成失败' && !(err instanceof SyntaxError)) {
        throw err;
      }
      throw new Error('生成失败');
    }
  }

  // 解析 Content-Disposition: attachment; filename*=UTF-8''<percent-encoded>
  const cd = String(resp.headers['content-disposition'] ?? '');
  const m = cd.match(/filename\*=UTF-8''([^;]+)/i);
  const filename = m ? decodeURIComponent(m[1]) : 'OBE目录.zip';

  return { blob: resp.data as Blob, filename };
}
