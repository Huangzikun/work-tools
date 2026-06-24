declare namespace Api {
  namespace Obe {
    // ============ mkdir ============

    interface MkdirRequest {
      className: string;
      courseName: string;
      teacherName: string;
      fixedDirTypes: string[];
      studentDirTypes: string[];
      roster: File;
    }

    interface MkdirResult {
      taskId: number;
      tree: TreeNode;
      studentCount: number;
      className: string;
      courseName: string;
      teacherName: string;
      fixedDirTypes: string[];
      studentDirTypes: string[];
    }

    // ============ 任务列表 / 详情 ============

    interface TaskSummary {
      id: number;
      className: string;
      courseName: string;
      teacherName: string;
      fixedDirTypes: string[];
      studentDirTypes: string[];
      studentCount: number;
      status: string;
      createdAt: string;
    }

    interface TaskPage {
      total: number;
      list: TaskSummary[];
      page: number;
      size: number;
    }

    interface Student {
      id: number;
      taskId: number;
      dirType: string;
      experimentLabel: string;
      studentId: string;
      studentName: string;
      studentClass?: string;
      studentDirName: string;
      uploadedFile?: string;
      matched: boolean;
      gradeStatus: 'pending' | 'grading' | 'graded' | 'failed';
      lastScore?: number;
      lastComment?: string;
      lastGradedFile?: string;
      lastGradeError?: string;
      lastGradedAt?: string;
    }

    interface GradingJob {
      id: number;
      taskId: number;
      dirType: string;
      experimentLabel: string;
      teacherName: string;
      signDate: string;
      signPicturePath: string;
      teacherPrompt: string;
      status: 'running' | 'completed' | 'failed' | 'cancelled';
      total: number;
      graded: number;
      failed: number;
      startedAt?: string;
      finishedAt?: string;
      errorSummary?: string;
    }

    interface TaskDetail {
      task: TaskSummary;
      // {dirType: {experimentLabel: Student[]}}
      studentsByDirAndExperiment: Record<string, Record<string, Student[]>>;
      // {dirType: {experimentLabel: GradingJob}}
      latestJobByDirAndExperiment: Record<string, Record<string, GradingJob>>;
      // {dirType: experimentLabel[]}
      experimentsByDir: Record<string, string[]>;
    }

    // ============ 目录树 ============

    interface TreeNode {
      key: string;
      label: string;
      type: 'dir' | 'file';
      size?: number;
      children?: TreeNode[];
    }

    // ============ 上传匹配（按 experiment 维度） ============

    interface MatchedItem {
      studentId: string;
      studentName: string;
      studentClass?: string;
      fileName: string;
      filePath: string;
    }

    interface AmbiguousItem {
      fileName: string;
      filePath: string;
      reason: string;
      candidates: Array<{ studentId: string; studentName: string }>;
    }

    interface UnmatchedItem {
      fileName: string;
      filePath?: string;
    }

    interface ExperimentUploadResult {
      experimentLabel: string;
      fileCount: number;
      matched: MatchedItem[];
      ambiguous: AmbiguousItem[];
      unmatched: UnmatchedItem[];
    }

    interface UploadResult {
      experiments: ExperimentUploadResult[];
    }

    interface AmbiguousResolution {
      fileName: string;
      filePath: string;
      studentId: string;
    }

    // ============ 触发批改 / 进度 ============

    interface GradeRequest {
      dirType: string;
      experimentLabel: string;
      teacherName: string;
      signDate: string;
      teacherPrompt: string;
      systemPrompt?: string;
      signPicture: File;
    }

    interface GradeResult {
      jobId: number;
      status: string;
    }

    interface Progress {
      jobId?: number;
      status?: string;
      total: number;
      graded: number;
      failed: number;
      currentStudent?: string | null;
      currentStudentId?: string | null;
      lastUpdate?: string;
    }
  }
}
