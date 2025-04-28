declare module "combine-word" {
  interface CombineWordOptions {
    pageBreak?: boolean;
    title?: string;
    subject?: string;
    author?: string;
    keywords?: string;
    description?: string;
    lastModifiedBy?: string;
    vision?: string;
  }

  class CombineWord {
    constructor(
      options: CombineWordOptions,
      files: Array<Buffer | ArrayBuffer>
    );

    docx: JSZip | null;

    save(type: string, callback: (fileData: any) => void): void;
  }

  export = CombineWord;
}
