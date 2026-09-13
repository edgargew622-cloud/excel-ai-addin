/**
 * Замер фактических возможностей установленного Excel.
 *
 * Решения плана опираются на измерение, а не на предположение о сборке Office.
 * Чистая часть (таблица требований и разрешение потолка) отделена от Office,
 * поэтому тестируется без Excel.
 */

/** Версии набора ExcelApi, которые проверяем. Достаточно с запасом вперёд. */
export const PROBED_VERSIONS = [
  "1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10",
  "1.11", "1.12", "1.13", "1.14", "1.15", "1.16", "1.17", "1.18", "1.19", "1.20"
] as const;

export interface FeatureRequirement {
  id: string;
  /** Что это даёт пользователю или плану. */
  label: string;
  /** Минимальная версия ExcelApi. */
  requires: string;
  /** Релиз плана, которому нужен ответ. */
  forRelease?: string;
  /** true — версия взята предположительно и требует подтверждения по документации. */
  versionUnconfirmed?: boolean;
}

/**
 * Возможности, от которых зависят решения плана. Список намеренно короткий:
 * проверяем то, что действительно нужно функциям надстройки.
 */
export const FEATURES: FeatureRequirement[] = [
  { id: "worksheet_copy", label: "Копия листа", requires: "1.7" },
  { id: "pivot_tables", label: "Сводные таблицы", requires: "1.8" },
  { id: "autofilter", label: "Автофильтр", requires: "1.9" },
  {
    id: "structural_monitor",
    label: "Монитор структурных изменений (защита custom undo)",
    requires: "1.9"
  },
  {
    id: "cell_properties",
    label: "Пакетное чтение свойств ячеек (точный снимок формата)",
    requires: "1.9",
    forRelease: "5"
  },
  { id: "comments", label: "Современные комментарии", requires: "1.10", forRelease: "3" },
  {
    id: "workbook_save",
    label: "Сохранение книги из кода (точка восстановления)",
    requires: "1.11",
    forRelease: "5"
  },
  {
    id: "direct_precedents",
    label: "Прямые влияющие ячейки",
    requires: "1.14",
    forRelease: "5",
    versionUnconfirmed: true
  },
  {
    id: "direct_dependents",
    label: "Прямые зависимые ячейки (точная область сравнения ошибок)",
    requires: "1.15",
    forRelease: "5",
    versionUnconfirmed: true
  },
  { id: "notes", label: "Старые примечания Excel", requires: "1.18" }
];

export interface FeatureStatus extends FeatureRequirement {
  available: boolean;
  /** Причина отсутствия, понятная человеку. Пусто, когда возможность есть. */
  reason: string;
}

export interface CapabilityReport {
  /** Наибольшая поддержанная версия ExcelApi или null, если не поддержана ни одна. */
  ceiling: string | null;
  supported: string[];
  features: FeatureStatus[];
  host: string;
  platform: string;
  officeVersion: string;
}

/** Сравнение версий вида "1.10" по числам, а не по строкам: "1.10" > "1.9". */
export function compareVersions(a: string, b: string): number {
  const pa = a.split(".").map(Number);
  const pb = b.split(".").map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const d = (pa[i] ?? 0) - (pb[i] ?? 0);
    if (d !== 0) return d;
  }
  return 0;
}

/**
 * Наборы требований не обязаны поддерживаться непрерывно, поэтому потолком
 * считаем наибольшую поддержанную версию, а не первый разрыв.
 */
export function resolveCeiling(supported: string[]): string | null {
  if (!supported.length) return null;
  return supported.reduce((max, v) => (compareVersions(v, max) > 0 ? v : max));
}

export function resolveFeatures(
  supported: string[],
  features: FeatureRequirement[] = FEATURES
): FeatureStatus[] {
  const set = new Set(supported);
  const ceiling = resolveCeiling(supported);
  return features.map((f) => {
    const available = set.has(f.requires);
    let reason = "";
    if (!available) {
      reason = ceiling
        ? `Нужен ExcelApi ${f.requires}; на этой сборке доступно до ${ceiling}.`
        : `Нужен ExcelApi ${f.requires}; ни один проверенный набор не поддержан.`;
    } else if (f.versionUnconfirmed) {
      reason = `Набор ${f.requires} поддержан, но привязка возможности к этой версии требует подтверждения по документации.`;
    }
    return { ...f, available, reason };
  });
}

/** Изолирует замер от Office, чтобы чистую часть можно было тестировать. */
export type SetProbe = (set: string, version: string) => boolean;

export function probeVersions(
  isSetSupported: SetProbe,
  versions: readonly string[] = PROBED_VERSIONS
): string[] {
  const supported: string[] = [];
  for (const v of versions) {
    let ok = false;
    try {
      ok = isSetSupported("ExcelApi", v);
    } catch {
      ok = false;
    }
    if (ok) supported.push(v);
  }
  return supported;
}

/** Замер на живом Office. Вызывается один раз при старте панели. */
export function measureCapabilities(): CapabilityReport {
  const probe: SetProbe = (set, version) => {
    try {
      return typeof Office !== "undefined" && Office.context.requirements.isSetSupported(set, version);
    } catch {
      return false;
    }
  };

  const supported = probeVersions(probe);
  const diagnostics = (() => {
    try {
      return Office?.context?.diagnostics;
    } catch {
      return undefined;
    }
  })();

  return {
    ceiling: resolveCeiling(supported),
    supported,
    features: resolveFeatures(supported),
    host: String(diagnostics?.host ?? "неизвестно"),
    platform: String(diagnostics?.platform ?? "неизвестно"),
    officeVersion: String(diagnostics?.version ?? "неизвестно")
  };
}

/**
 * Одна строка для технического журнала. Содержимого ячеек и ключей здесь нет
 * и быть не может: только версии и признаки поддержки.
 */
export function formatCapabilityLog(report: CapabilityReport): string {
  const missing = report.features.filter((f) => !f.available).map((f) => f.id);
  return [
    `host=${report.host}`,
    `platform=${report.platform}`,
    `office=${report.officeVersion}`,
    `excelApiCeiling=${report.ceiling ?? "нет"}`,
    `missing=${missing.length ? missing.join(",") : "нет"}`
  ].join(" ");
}
