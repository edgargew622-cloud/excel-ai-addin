/**
 * Понятный текст вместо «fetch failed» (08.10.2026: xAI не отвечал из сети
 * пользователя — «Connect Timeout» трижды, а панель показала голое
 * «fetch failed», и было непонятно, что делать).
 */
const CONNECT_CODES = new Set(["UND_ERR_CONNECT_TIMEOUT", "ETIMEDOUT", "ECONNREFUSED", "ENOTFOUND", "EAI_AGAIN", "ECONNRESET", "UND_ERR_SOCKET"]);

export function networkErrorText(error: any, providerLabel: string): string | null {
  const code = error?.cause?.code ?? error?.code;
  if (!CONNECT_CODES.has(code) && !/^fetch failed$/i.test(String(error?.message ?? ""))) return null;
  const why = code === "ENOTFOUND" || code === "EAI_AGAIN"
    ? "адрес сервиса не находится (нет интернета или DNS)"
    : code === "ECONNRESET" || code === "UND_ERR_SOCKET"
      ? "соединение оборвалось"
      : "сервер не ответил за 10 секунд, три попытки";
  return `Не удалось связаться с ${providerLabel}: ${why}. Проверьте интернет; если с ним всё в порядке — сервис может быть недоступен из вашей сети. ` +
    "Выберите другого провайдера в списке сверху (например DeepSeek) и повторите просьбу.";
}
