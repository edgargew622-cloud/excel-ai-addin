/**
 * Шифрование ключей средствами Windows (DPAPI, область CurrentUser).
 *
 * Node не умеет DPAPI сам, а нативный модуль усложнил бы установку, поэтому
 * вызывается встроенный Windows PowerShell 5.1. Данные передаются через
 * переменную окружения дочернего процесса: не в командной строке, которую
 * видят все процессы, и не через stdin, который powershell.exe с
 * -EncodedCommand читает ненадёжно. Переменная окружения видна только той же
 * учётной записи — а она и так может расшифровать DPAPI.
 */

import { execFile } from "node:child_process";
import { join } from "node:path";
import type { Protector } from "./keyStore.js";

export const DPAPI_INPUT_VARIABLE = "EXCEL_AI_DPAPI_INPUT";
const ENTROPY = "excel-ai-addin/provider-keys/v1";

export function dpapiScript(mode: "protect" | "unprotect"): string {
  const method = mode === "protect" ? "Protect" : "Unprotect";
  return [
    "$ErrorActionPreference = 'Stop'",
    "Add-Type -AssemblyName System.Security",
    `$data = [Convert]::FromBase64String($env:${DPAPI_INPUT_VARIABLE})`,
    `$entropy = [Text.Encoding]::UTF8.GetBytes('${ENTROPY}')`,
    `$out = [Security.Cryptography.ProtectedData]::${method}($data, $entropy, [Security.Cryptography.DataProtectionScope]::CurrentUser)`,
    "[Console]::Out.Write([Convert]::ToBase64String($out))"
  ].join("; ");
}

export function dpapiArgs(mode: "protect" | "unprotect"): string[] {
  return [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-EncodedCommand",
    Buffer.from(dpapiScript(mode), "utf16le").toString("base64")
  ];
}

/** Полный путь: не зависим от PATH, в который мог попасть чужой powershell.exe. */
function powershellPath(): string {
  return join(process.env.SystemRoot ?? "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
}

function run(mode: "protect" | "unprotect", input: Buffer): Promise<Buffer> {
  return new Promise((resolvePromise, reject) => {
    execFile(
      powershellPath(),
      dpapiArgs(mode),
      {
        env: { ...process.env, [DPAPI_INPUT_VARIABLE]: input.toString("base64") },
        windowsHide: true,
        timeout: 30_000,
        maxBuffer: 1024 * 1024
      },
      (error, stdout) => {
        // stderr PowerShell может процитировать входные данные, поэтому наружу
        // уходит только общий текст без подробностей.
        if (error) return reject(new Error(`DPAPI: ${mode === "protect" ? "шифрование" : "расшифровка"} не удалось`));
        const text = String(stdout).trim();
        if (!/^[A-Za-z0-9+/]+={0,2}$/.test(text)) return reject(new Error("DPAPI: неожиданный ответ PowerShell"));
        resolvePromise(Buffer.from(text, "base64"));
      }
    );
  });
}

export function windowsDpapi(): Protector {
  return {
    available: true,
    protect: (plain) => run("protect", plain),
    unprotect: (blob) => run("unprotect", blob)
  };
}

export function unavailableProtector(): Protector {
  const refuse = () => Promise.reject(new Error("защищённое хранилище недоступно на этой системе"));
  return { available: false, protect: refuse, unprotect: refuse };
}

export function systemProtector(): Protector {
  return process.platform === "win32" ? windowsDpapi() : unavailableProtector();
}
