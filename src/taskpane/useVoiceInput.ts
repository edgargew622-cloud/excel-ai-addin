import { useEffect, useRef, useState, type Dispatch, type SetStateAction } from "react";
import { appendDictation, SpeechInput, type SpeechPhase } from "./speechInput";

export function useVoiceInput(setDraft: Dispatch<SetStateAction<string>>, enabled = true) {
  const [phase, setPhase] = useState<SpeechPhase>("idle");
  const [message, setMessage] = useState("");
  const controller = useRef<SpeechInput | null>(null);
  if (!controller.current) controller.current = new SpeechInput({
    phase: setPhase,
    transcript: (text) => setDraft((draft) => appendDictation(draft, text)),
    message: setMessage
  });
  useEffect(() => () => controller.current?.cancel(), []);
  useEffect(() => {
    if (!enabled) { controller.current?.cancel(); setMessage(""); }
  }, [enabled]);
  useEffect(() => {
    const cancel = () => controller.current?.cancel();
    window.addEventListener("pagehide", cancel);
    return () => window.removeEventListener("pagehide", cancel);
  }, []);
  return {
    active: phase !== "idle",
    stopping: phase === "stopping",
    message,
    isActive: () => controller.current?.active ?? false,
    toggle: () => phase === "idle" ? controller.current?.start() : controller.current?.stop(),
    cancel: () => { controller.current?.cancel(); setMessage(""); }
  };
}
