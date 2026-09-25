/**
 * Токен панели: без него локальный API не отвечает (этап 8, 8.0.1).
 *
 * Аудит 24 сентября 2026 года (SEC-03): сервер отличал панель от прочих
 * клиентов только по петлевому адресу и заголовкам браузера. Другой
 * пользователь того же компьютера мог тратить ключи и менять сохранённые.
 *
 * Токен — случайные 32 байта в server/panel-token внутри папки надстройки,
 * доступ к которой установщик оставляет только владельцу. Панель получает
 * его из адреса, записанного в манифест каталога Excel, — каталог лежит в той
 * же закрытой папке. От программы, запущенной тем же пользователем Windows,
 * токен не защищает и не может защитить: она так же может расшифровать ключи
 * DPAPI. Защищает от других пользователей компьютера и от случайных клиентов.
 *
 * Тот же формат пишет scripts/panel-token.ps1: кто первым, тот и создаёт.
 */

import { randomBytes, timingSafeEqual } from "node:crypto";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import type { NextFunction, Request, Response } from "express";

export const PANEL_TOKEN_HEADER = "x-panel-token";

export function loadOrCreatePanelToken(file: string): string {
  if (!existsSync(file)) {
    const token = randomBytes(32).toString("base64url");
    // wx: если файл успел создать сценарий установки, не перезаписываем его.
    try { writeFileSync(file, token, { encoding: "utf8", flag: "wx", mode: 0o600 }); } catch { /* уже создан */ }
  }
  const token = readFileSync(file, "utf8").trim();
  if (!/^[A-Za-z0-9_-]{32,}$/.test(token)) throw new Error(`Файл токена панели повреждён: ${file}. Удалите его и перезапустите сервер.`);
  return token;
}

export function tokenMatches(expected: string, presented: unknown): boolean {
  if (typeof presented !== "string") return false;
  const a = Buffer.from(expected);
  const b = Buffer.from(presented);
  return a.length === b.length && timingSafeEqual(a, b);
}

/**
 * Проверка стоит до разбора тела запроса: чужой клиент не должен тратить
 * даже разбор JSON. Открыт только /api/health — по нему установщик и
 * автозапуск узнают, что сервер жив; секретов в нём нет.
 */
export function requirePanelToken(token: string) {
  return (req: Request, res: Response, next: NextFunction) => {
    if (req.path === "/health") return next();
    if (tokenMatches(token, req.headers[PANEL_TOKEN_HEADER])) return next();
    res.status(401).json({ error: { message: "Запрос без токена панели. Откройте панель из Excel; если она открыта — переустановите надстройку." } });
  };
}
