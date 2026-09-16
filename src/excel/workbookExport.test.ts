import test from "node:test";
import assert from "node:assert/strict";
import { MAX_SLICE_BYTES, measureWorkbookExport } from "./workbookExport";

/** Макет Office: книга из нескольких срезов. */
function fakeOffice(options: {
  size: number;
  sliceCount: number;
  sliceLength?: number;
  failAtSlice?: number;
  unavailable?: boolean;
}) {
  let closed = false;
  const file = {
    size: options.size,
    sliceCount: options.sliceCount,
    getSliceAsync: (index: number, callback: (result: any) => void) => {
      if (options.failAtSlice === index) {
        callback({ status: "failed", error: { message: "срез потерян" } });
        return;
      }
      callback({ status: "succeeded", value: { data: { length: options.sliceLength ?? options.size / options.sliceCount } } });
    },
    closeAsync: (callback?: (result: any) => void) => { closed = true; callback?.({}); }
  };
  (globalThis as any).Office = {
    FileType: { Compressed: "compressed" },
    context: {
      document: options.unavailable ? {} : {
        getFileAsync: (_type: string, _opts: unknown, callback: (result: any) => void) =>
          callback({ status: "succeeded", value: file })
      }
    }
  };
  return { wasClosed: () => closed };
}

test("a complete export reports size, slices and timing", async () => {
  fakeOffice({ size: 8000, sliceCount: 2, sliceLength: 4000 });
  const measurement = await measureWorkbookExport();

  assert.equal(measurement.available, true);
  assert.equal(measurement.sizeBytes, 8000);
  assert.equal(measurement.sliceCount, 2);
  assert.equal(measurement.complete, true);
  assert.equal(measurement.reason, undefined);
  assert.equal(typeof measurement.totalMs, "number");
});

test("slices that do not add up to the stated size are not called complete", async () => {
  // Копия из таких срезов была бы неполной, и молчать об этом нельзя.
  fakeOffice({ size: 8000, sliceCount: 2, sliceLength: 1000 });
  const measurement = await measureWorkbookExport();

  assert.equal(measurement.complete, false);
  assert.match(measurement.reason ?? "", /2000 байт вместо 8000/);
});

test("a lost slice is reported without pretending the export worked", async () => {
  fakeOffice({ size: 8000, sliceCount: 2, failAtSlice: 1 });
  const measurement = await measureWorkbookExport();

  assert.equal(measurement.available, true);
  assert.equal(measurement.complete, false);
  assert.match(measurement.reason ?? "", /срез потерян/);
});

test("an Excel without file export says so instead of throwing", async () => {
  fakeOffice({ size: 0, sliceCount: 0, unavailable: true });
  const measurement = await measureWorkbookExport();

  assert.equal(measurement.available, false);
  assert.match(measurement.reason ?? "", /getFileAsync/);
});

test("the file handle is always closed, including after a failure", async () => {
  const ok = fakeOffice({ size: 100, sliceCount: 1, sliceLength: 100 });
  await measureWorkbookExport();
  assert.equal(ok.wasClosed(), true, "после успеха");

  const failed = fakeOffice({ size: 100, sliceCount: 1, failAtSlice: 0 });
  await measureWorkbookExport();
  assert.equal(failed.wasClosed(), true, "после сбоя");
});

test("an oversized workbook is not read into memory", async () => {
  let sliceReads = 0;
  fakeOffice({ size: 500 * 1024 * 1024, sliceCount: 200 });
  const office = (globalThis as any).Office.context.document;
  const original = office.getFileAsync;
  office.getFileAsync = (type: string, opts: unknown, callback: (result: any) => void) =>
    original(type, opts, (result: any) => {
      const file = result.value;
      const originalSlice = file.getSliceAsync.bind(file);
      file.getSliceAsync = (index: number, cb: (result: any) => void) => { sliceReads += 1; originalSlice(index, cb); };
      callback(result);
    });

  const measurement = await measureWorkbookExport();
  assert.equal(sliceReads, 0, "срезы не читались");
  assert.equal(measurement.complete, false);
  assert.match(measurement.reason ?? "", /предела замера/);
});

test("the slice size stays within the Office limit", async () => {
  fakeOffice({ size: 10, sliceCount: 1, sliceLength: 10 });
  const tooBig = await measureWorkbookExport({ sliceSizeBytes: 99_000_000 });
  assert.equal(tooBig.sliceSizeBytes, MAX_SLICE_BYTES);

  fakeOffice({ size: 10, sliceCount: 1, sliceLength: 10 });
  const tooSmall = await measureWorkbookExport({ sliceSizeBytes: 1 });
  assert.equal(tooSmall.sliceSizeBytes, 1024);
});
