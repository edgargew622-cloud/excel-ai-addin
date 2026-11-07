/** Browser speech recognition; no provider keys, recording files or automatic send. */
export interface SpeechResult {
  readonly isFinal: boolean;
  readonly 0: { readonly transcript: string };
}

export interface SpeechEngine {
  lang: string;
  continuous: boolean;
  interimResults: boolean;
  onstart: (() => void) | null;
  onend: (() => void) | null;
  onerror: ((event: { error: string }) => void) | null;
  onresult: ((event: { resultIndex: number; results: ArrayLike<SpeechResult> }) => void) | null;
  start(): void;
  stop(): void;
  abort(): void;
}

export type SpeechPhase = "idle" | "starting" | "listening" | "stopping";
export type SpeechFactory = () => SpeechEngine | null;
export const SPEECH_FALLBACK = "В этой версии Excel или браузера встроенное распознавание недоступно. Выберите поле ввода и нажмите Win + H в Windows; на Mac используйте системную диктовку.";

export function browserSpeechEngine(): SpeechEngine | null {
  const host = globalThis as unknown as {
    SpeechRecognition?: new () => SpeechEngine;
    webkitSpeechRecognition?: new () => SpeechEngine;
  };
  const Engine = host.SpeechRecognition ?? host.webkitSpeechRecognition;
  return Engine ? new Engine() : null;
}

export function speechError(error: string): string {
  switch (error) {
    case "not-allowed": case "service-not-allowed":
      return "Доступ к микрофону или распознаванию не разрешён. Проверьте разрешения Excel/браузера либо используйте системную диктовку.";
    case "audio-capture": return "Микрофон не найден или занят другой программой.";
    case "no-speech": return "Речь не распознана. Нажмите микрофон и попробуйте ещё раз.";
    case "network": return "Служба распознавания недоступна. Проверьте интернет; в Windows можно использовать Win + H в поле ввода.";
    case "language-not-supported": return "Русский язык недоступен в службе распознавания. Используйте системную диктовку.";
    default: return "Не удалось распознать речь. Попробуйте ещё раз или используйте системную диктовку.";
  }
}

export function appendDictation(draft: string, transcript: string): string {
  const text = transcript.trim();
  if (!text) return draft;
  return draft + (draft && !/\s$/.test(draft) ? " " : "") + text;
}

/** One session owns its callbacks. Disposed/aborted sessions cannot change a new draft. */
export class SpeechInput {
  private engine: SpeechEngine | null = null;
  private timer: ReturnType<typeof setTimeout> | null = null;
  private phase: SpeechPhase = "idle";

  get active(): boolean { return this.engine !== null; }

  constructor(private readonly events: {
    phase: (value: SpeechPhase) => void;
    transcript: (text: string) => void;
    message: (text: string) => void;
  }, private readonly factory: SpeechFactory = browserSpeechEngine) {}

  private setPhase(value: SpeechPhase): void {
    this.phase = value;
    this.events.phase(value);
  }

  start(): void {
    if (this.engine) return;
    let engine: SpeechEngine | null;
    try { engine = this.factory(); }
    catch { this.events.message(SPEECH_FALLBACK); return; }
    if (!engine) { this.events.message(SPEECH_FALLBACK); return; }
    this.engine = engine;
    engine.lang = "ru-RU";
    engine.continuous = true;
    engine.interimResults = true;
    const committed = new Set<number>();
    let heard = false;
    const current = () => this.engine === engine;
    engine.onstart = () => {
      if (!current()) return;
      // A second click can stop while the browser is still asking for permission.
      if (this.phase !== "starting") { this.cancel(); return; }
      this.setPhase("listening");
      this.events.message("Слушаю… Речь распознаёт служба браузера; аудио может передаваться ей. Текст не отправляется в чат автоматически.");
    };
    engine.onresult = (event) => {
      if (!current()) return;
      const interim: string[] = [];
      for (let i = event.resultIndex; i < event.results.length; i++) {
        const result = event.results[i];
        const text = result[0].transcript.trim();
        if (result.isFinal && !committed.has(i)) {
          committed.add(i);
          if (text) { heard = true; this.events.transcript(text); }
        } else if (!result.isFinal && text) interim.push(text);
      }
      if (interim.length) this.events.message(`Слышу: ${interim.join(" ")}`);
    };
    engine.onerror = ({ error }) => {
      if (!current()) return;
      this.events.message(error === "aborted" ? "Диктовка остановлена." : speechError(error));
      this.detach(engine);
      try { engine.abort(); } catch { /* already stopped */ }
    };
    engine.onend = () => {
      if (!current()) return;
      this.detach(engine);
      this.events.message(heard ? "Диктовка завершена. Проверьте текст и нажмите «Отправить»." : speechError("no-speech"));
    };
    this.setPhase("starting");
    this.events.message("Подключение микрофона… Речь распознаёт служба браузера; аудио может передаваться ей.");
    // Limit both permission/startup waits and an accidentally open microphone.
    this.timer = setTimeout(() => {
      if (!current()) return;
      this.cancel();
      this.events.message("Диктовка остановлена по времени. Для продолжения снова нажмите микрофон.");
    }, 120_000);
    try { engine.start(); }
    catch {
      this.detach(engine);
      try { engine.abort(); } catch { /* start failed */ }
      this.events.message(speechError("start"));
    }
  }

  stop(): void {
    const engine = this.engine;
    if (!engine || this.phase === "stopping") return;
    if (this.phase === "starting") { this.cancel(); this.events.message("Диктовка остановлена."); return; }
    this.setPhase("stopping");
    this.events.message("Завершаю распознавание…");
    if (this.timer) clearTimeout(this.timer);
    this.timer = setTimeout(() => {
      if (this.engine !== engine) return;
      this.cancel();
      this.events.message("Диктовка завершена. Проверьте текст перед отправкой.");
    }, 5_000);
    try { engine.stop(); }
    catch { this.cancel(); this.events.message(speechError("stop")); }
  }

  cancel(): void {
    const engine = this.engine;
    if (!engine) return;
    this.detach(engine);
    try { engine.abort(); } catch { /* already stopped */ }
  }

  private detach(engine: SpeechEngine): void {
    if (this.engine !== engine) return;
    this.engine = null;
    engine.onstart = engine.onend = engine.onerror = engine.onresult = null;
    if (this.timer) clearTimeout(this.timer);
    this.timer = null;
    this.setPhase("idle");
  }
}
