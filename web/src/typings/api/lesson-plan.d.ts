declare namespace Api {
  namespace LessonPlan {
    interface CourseInfo {
      课程名称: string;
      英文名称?: string;
      学分?: string;
      理论学时?: string;
      实践学时?: string;
      线上学时?: string;
      适用专业?: string;
      先修课程?: string;
      开课学期?: string;
      课程教研室?: string;
      [key: string]: string | undefined;
    }

    interface TeacherInfo {
      授课教师: string;
      所属单位?: string;
      撰写日期?: string;
      课程类型?: string;
      课程性质?: string;
      [key: string]: string | undefined;
    }

    type TaskStatus =
      | 'pending'
      | 'parsing'
      | 'generating'
      | 'building'
      | 'completed'
      | 'failed';

    interface TaskSummary {
      id: number;
      syllabusName: string;
      totalLessons: number;
      batchSize: number;
      courseInfo: CourseInfo;
      teacherInfo: TeacherInfo;
      status: TaskStatus;
      progressTotal: number;
      progressDone: number;
      progressLabel?: string;
      outputFile?: string;
      errorSummary?: string;
      createdAt?: string;
      startedAt?: string;
      finishedAt?: string;
    }

    interface TaskPage {
      total: number;
      list: TaskSummary[];
      page: number;
      size: number;
    }

    interface Progress {
      taskId: number;
      status: TaskStatus | 'none';
      total: number;
      done: number;
      label?: string;
      lastUpdate?: string;
    }
  }
}
