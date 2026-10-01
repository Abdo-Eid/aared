/** pptx-browser ships plain JS with no types; only the surface the bench uses. */
declare module "pptx-browser" {
  export class PptxRenderer {
    readonly slideCount: number;
    load(
      source: File | Blob | ArrayBuffer | Uint8Array,
      onProgress?: (progress: number, message: string) => void,
    ): Promise<void>;
    renderSlide(index: number, canvas: HTMLCanvasElement, width?: number): Promise<void>;
    destroy(): void;
  }
}
