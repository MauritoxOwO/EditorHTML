type TesseractWorker = {
  recognize: (image: string) => Promise<{ data?: { text?: string } }>;
};

type OcrWorkerRequest = {
  id: string;
  src: string;
};

type OcrWorkerResponse = {
  id: string;
  ok: boolean;
  text?: string;
  error?: string;
};

let workerPromise: Promise<TesseractWorker> | null = null;

self.onmessage = (event: MessageEvent<OcrWorkerRequest>) => {
  void recognize(event.data);
};

async function recognize(request: OcrWorkerRequest): Promise<void> {
  try {
    const worker = await getWorker();
    const result = await worker.recognize(request.src);
    postResult({
      id: request.id,
      ok: true,
      text: result.data?.text ?? "",
    });
  } catch (error) {
    postResult({
      id: request.id,
      ok: false,
      error: error instanceof Error ? error.message : String(error),
    });
  }
}

async function getWorker(): Promise<TesseractWorker> {
  if (!workerPromise) {
    workerPromise = import("tesseract.js").then(async (module) => {
      const tesseract = module as unknown as {
        createWorker: (
          langs?: string | string[],
          oem?: number,
          options?: { logger?: (message: unknown) => void }
        ) => Promise<TesseractWorker>;
      };
      return tesseract.createWorker(["spa", "eng"], 1, {
        logger: () => undefined,
      });
    });
  }

  return workerPromise;
}

function postResult(response: OcrWorkerResponse): void {
  self.postMessage(response);
}

export {};
