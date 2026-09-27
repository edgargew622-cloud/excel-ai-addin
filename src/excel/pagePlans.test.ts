import test from "node:test";
import assert from "node:assert/strict";
import { executePageLayoutPlan, parsePageRequest, POINTS_PER_CM, preparePageLayoutPlan } from "./pagePlans";
import { PLANNED_TOOLS } from "./plans";
import { clear as clearUndo, setUndoMonitorReady, undoLast } from "./undo";

/**
 * Лист с параметрами страницы, как их отдал Excel (Office 2021) в замере
 * 28 сентября 2026 года: поля в пунктах (50,4 и 54 по умолчанию), при
 * вписывании scale — null, 0 страниц читается как null, пустая строка снимает
 * область печати и строки заголовков; адреса — с именем листа, строки — «1:1».
 */
function pageSheet(options: { ignore?: string[] } = {}) {
  const ignored = (name: string) => options.ignore?.includes(name) === true;
  const state: any = {
    orientation: "Portrait",
    leftMargin: 50.4, rightMargin: 50.4, topMargin: 54, bottomMargin: 54,
    zoom: { scale: 100, horizontalFitToPages: null, verticalFitToPages: null },
    printArea: null as string | null,
    titleRows: null as string | null,
    texts: { leftHeader: "", centerHeader: "", rightHeader: "", leftFooter: "", centerFooter: "", rightFooter: "" }
  };
  const layout: any = {
    get orientation() { return state.orientation; },
    set orientation(value: string) { if (!ignored("orientation")) state.orientation = value; },
    get leftMargin() { return state.leftMargin; },
    set leftMargin(value: number) { state.leftMargin = ignored("left") ? 50.4 : value; },
    get rightMargin() { return state.rightMargin; },
    set rightMargin(value: number) { state.rightMargin = value; },
    get topMargin() { return state.topMargin; },
    set topMargin(value: number) { state.topMargin = value; },
    get bottomMargin() { return state.bottomMargin; },
    set bottomMargin(value: number) { state.bottomMargin = value; },
    get zoom() { return { ...state.zoom }; },
    set zoom(value: any) {
      if (value.scale !== undefined) state.zoom = { scale: value.scale, horizontalFitToPages: null, verticalFitToPages: null };
      else state.zoom = { scale: null, horizontalFitToPages: value.horizontalFitToPages || null, verticalFitToPages: value.verticalFitToPages || null };
    },
    load: () => undefined,
    setPrintArea: (address: string) => { state.printArea = address ? `Отчёт!${address}` : null; },
    setPrintTitleRows: (rows: string) => { state.titleRows = rows ? `Отчёт!${rows.replace(/\$/g, "")}` : null; },
    getPrintAreaOrNullObject: () => ({ get isNullObject() { return state.printArea === null; }, get address() { return state.printArea; }, load: () => undefined }),
    getPrintTitleRowsOrNullObject: () => ({ get isNullObject() { return state.titleRows === null; }, get address() { return state.titleRows; }, load: () => undefined }),
    headersFooters: {
      defaultForAllPages: new Proxy(state.texts, {
        get: (target, key) => (key === "load" ? () => undefined : target[key as string]),
        set: (target, key, value) => { target[key as string] = ignored("header") && key === "centerHeader" ? "" : value; return true; }
      })
    }
  };
  const sheet: any = { id: "sheet-1", name: "Отчёт", load: () => undefined, protection: { protected: false, load: () => undefined }, pageLayout: layout };
  (globalThis as any).Office = { context: { document: { url: "C:/page.xlsx" }, requirements: { isSetSupported: () => true } } };
  (globalThis as any).Excel = {
    run: async (fn: any) => fn({ workbook: { worksheets: { getActiveWorksheet: () => sheet, getItem: () => sheet } }, sync: async () => undefined })
  };
  return state;
}

test("set_page_layout goes through the plan registry", () => {
  assert.ok(PLANNED_TOOLS.includes("set_page_layout"));
});

test("landscape, fit to one page wide, a print area, title rows and a footer are set and read back", async () => {
  const state = pageSheet();
  const plan = await preparePageLayoutPlan({
    sheet: "Отчёт", orientation: "landscape", fitToPagesWide: 1, printArea: "a1:f40", printTitleRows: "1", footer: { center: "Стр. &P из &N" }
  });
  assert.deepEqual(plan.preview.slice(0, 4), [
    "Ориентация: книжная → альбомная",
    "Масштаб: 100% → вписать: 1 стр. в ширину × авто в высоту",
    "Область печати: весь лист → A1:F40",
    "Повтор строк на каждой странице: нет → 1:1"
  ]);
  assert.ok(plan.preview.some((line) => /&P — номер страницы/.test(line)));
  const result = await executePageLayoutPlan(plan) as any;
  assert.equal(result.executionState, "verified");
  assert.equal(state.orientation, "Landscape");
  assert.deepEqual(state.zoom, { scale: null, horizontalFitToPages: 1, verticalFitToPages: null });
  assert.equal(state.printArea, "Отчёт!A1:F40");
  assert.equal(state.titleRows, "Отчёт!1:1");
  assert.equal(state.texts.centerFooter, "Стр. &P из &N");
});

test("margins are given in centimetres and stored in points", async () => {
  const state = pageSheet();
  await executePageLayoutPlan(await preparePageLayoutPlan({ sheet: "Отчёт", marginsCm: { left: 1, right: 1.5 } }));
  assert.ok(Math.abs(state.leftMargin - POINTS_PER_CM) < 0.01);
  assert.ok(Math.abs(state.rightMargin - 1.5 * POINTS_PER_CM) < 0.01);
  assert.equal(state.topMargin, 54, "незапрошенное поле не тронуто");
});

test("a value Excel did not take is named, and undo brings back only what was changed", async () => {
  pageSheet({ ignore: ["left"] });
  await assert.rejects(async () => executePageLayoutPlan(await preparePageLayoutPlan({ sheet: "Отчёт", marginsCm: { left: 1 } })), (error: any) => {
    assert.equal(error.executionState, "applied");
    assert.match(error.message, /поле left 1,78 см вместо 1,00/);
    return true;
  });

  const state = pageSheet();
  setUndoMonitorReady(true);
  try {
    await executePageLayoutPlan(await preparePageLayoutPlan({ sheet: "Отчёт", orientation: "landscape", printArea: "A1:C10" }));
    state.texts.leftHeader = "ручная правка другого поля";
    await undoLast();
    assert.equal(state.orientation, "Portrait");
    assert.equal(state.printArea, null);
    assert.equal(state.texts.leftHeader, "ручная правка другого поля", "отмена не трогает то, что операция не меняла");
  } finally {
    clearUndo();
    setUndoMonitorReady(false);
  }
});

test("page settings changed after the preview stop the operation; the same settings are refused up front", async () => {
  const state = pageSheet();
  const plan = await preparePageLayoutPlan({ sheet: "Отчёт", orientation: "landscape" });
  state.leftMargin = 20;
  await assert.rejects(() => executePageLayoutPlan(plan), (error: any) => error.executionState === "failed_before_write");
  pageSheet();
  await assert.rejects(() => preparePageLayoutPlan({ sheet: "Отчёт", orientation: "portrait" }), /Менять нечего/);
});

test("requests the panel cannot carry out are refused before the card", () => {
  assert.throws(() => parsePageRequest({ scale: 90, fitToPagesWide: 1 }), /либо масштабирует, либо вписывает/);
  assert.throws(() => parsePageRequest({ scale: 5 }), /от 10 до 400/);
  assert.throws(() => parsePageRequest({ printTitleRows: "A1" }), /номера строк/);
  assert.throws(() => parsePageRequest({ marginsCm: { left: -1 } }), /от 0 до 20 см/);
  assert.throws(() => parsePageRequest({}), /Не указано, что менять/);
  assert.deepEqual(parsePageRequest({ printArea: "none" }), { printArea: null });
});
