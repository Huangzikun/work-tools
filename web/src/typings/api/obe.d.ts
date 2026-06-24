declare namespace Api {
  namespace Obe {
    interface MkdirRequest {
      className: string;
      courseName: string;
      teacherName: string;
      fixedDirTypes: string[];
      studentDirTypes: string[];
      roster: File;
    }

    interface MkdirResult {
      blob: Blob;
      filename: string;
    }
  }
}
