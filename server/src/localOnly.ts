/**
 * Граница доступа локального сервера.
 *
 * До объединения портов эти проверки жили в dev-плагине Vite и для собранной
 * панели не действовали вовсе. Теперь они в рабочем сервере, поэтому вынесены
 * отдельно и покрыты тестами: ошибка здесь открывает доступ к биллингу.
 */

/** Петлевые адреса во всех формах, в которых их отдаёт Node. */
export function isLoopbackAddress(address: string | undefined | null): boolean {
  if (!address) return false;
  const value = address.trim().toLowerCase();
  if (!value) return false;

  // IPv6 может прийти с зоной: ::1%lo0
  const withoutZone = value.split("%")[0];

  if (withoutZone === "::1") return true;
  // IPv4, отображённый в IPv6: ::ffff:127.0.0.1
  const mapped = withoutZone.startsWith("::ffff:") ? withoutZone.slice("::ffff:".length) : withoutZone;

  // Вся сеть 127.0.0.0/8 петлевая, не только 127.0.0.1.
  const octets = mapped.split(".");
  if (octets.length !== 4) return false;
  const numbers = octets.map((o) => (/^\d{1,3}$/.test(o) ? Number(o) : NaN));
  if (numbers.some((n) => Number.isNaN(n) || n > 255)) return false;
  return numbers[0] === 127;
}

/**
 * Панель и API теперь одного происхождения. Свой запрос либо не присылает
 * Origin (обычная навигация, GET статики), либо присылает собственный.
 * Всё остальное — сторонний сайт, обратившийся к локальному серверу.
 */
export function isAllowedOrigin(origin: string | undefined | null, port: number): boolean {
  if (origin === undefined || origin === null || origin === "") return true;
  const value = origin.trim().toLowerCase();
  // Некоторые webview присылают строку "null" для непрозрачного origin.
  if (value === "null") return false;

  let parsed: URL;
  try {
    parsed = new URL(value);
  } catch {
    return false;
  }
  if (parsed.protocol !== "https:") return false;
  if (parsed.port !== String(port)) return false;
  // URL отдаёт IPv6-хост в скобках: [::1]. Для сравнения с адресом петли их
  // нужно снять, иначе собственный origin из IPv6-резолва localhost отвергается.
  const hostname = parsed.hostname.replace(/^\[|\]$/g, "");
  return hostname === "localhost" || isLoopbackAddress(hostname);
}
