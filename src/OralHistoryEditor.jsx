import { useState, useRef, useCallback, useEffect, useMemo } from "react";

/* ── Utilities ──────────────────────────────────────────────── */
const uid = () => Math.random().toString(36).slice(2, 10);

function parseTime(str) {
  const m = str.match(/(\d{2}):(\d{2}):(\d{2})[,.](\d{3})/);
  if (!m) return 0;
  return +m[1] * 3600 + +m[2] * 60 + +m[3] + +m[4] / 1000;
}

function fmtTime(s) {
  if (s == null || isNaN(s)) return "00:00.00";
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const sec = (s % 60).toFixed(1);
  return `${h > 0 ? h + ":" : ""}${String(m).padStart(2, "0")}:${sec.padStart(4, "0")}`;
}

function parseSRT(text) {
  return text.trim().split(/\n\n+/).map((block) => {
    const lines = block.split("\n");
    const tm = lines.find((l) => l.includes("-->"));
    if (!tm) return null;
    const [s, e] = tm.split("-->").map((t) => parseTime(t.trim()));
    const txt = lines.slice(lines.indexOf(tm) + 1).join(" ").replace(/<[^>]+>/g, "").trim();
    if (!txt) return null;
    return { id: uid(), text: txt, originalText: txt, start: s, end: e, originalIndex: 0 };
  }).filter(Boolean).map((seg, i) => ({ ...seg, originalIndex: i }));
}

function parseVTT(text) {
  const body = text.replace(/^WEBVTT[^\n]*\n/, "").replace(/^NOTE[^\n]*\n(?:[^\n]+\n)*/gm, "");
  return parseSRT(body);
}

function detectAndParse(text) {
  return text.trimStart().startsWith("WEBVTT") ? parseVTT(text) : parseSRT(text);
}

/* ── Audio chunking for streaming transcription ── */
const MAX_CHUNK_BYTES = 24 * 1024 * 1024; // 24 MB — stay under OpenAI's 25 MB limit

/** Calculate max chunk duration (seconds) that keeps WAV under MAX_CHUNK_BYTES when encoded as mono 16-bit PCM. */
function maxChunkDuration(sampleRate) {
  const bytesPerSample = 2; // 16-bit mono
  return Math.floor((MAX_CHUNK_BYTES - 44) / (sampleRate * bytesPerSample));
}

/** Downmix multi-channel Float32Arrays to a single mono Float32Array. */
function downmixToMono(channelData) {
  if (channelData.length === 1) return channelData[0];
  const len = channelData[0].length;
  const mono = new Float32Array(len);
  const n = channelData.length;
  for (let i = 0; i < len; i++) {
    let sum = 0;
    for (let ch = 0; ch < n; ch++) sum += channelData[ch][i];
    mono[i] = sum / n;
  }
  return mono;
}

/**
 * Probe the audio file by decoding a small initial slice to get sample rate,
 * channel count, and estimate total duration without loading the whole file.
 */
async function probeAudioFile(file, onProgress) {
  onProgress?.("Probing audio file for metadata…");
  // Read up to 1 MB to decode enough header + samples for metadata
  const probeSize = Math.min(file.size, 1024 * 1024);
  const probeBuf = await file.slice(0, probeSize).arrayBuffer();
  const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  try {
    onProgress?.("Decoding audio header…");
    let decoded;
    try {
      decoded = await audioCtx.decodeAudioData(probeBuf);
    } catch (e) {
      throw new Error(`Could not decode "${file.name}" — the file may be corrupted or in an unsupported format (try MP3, WAV, M4A, OGG, or FLAC)`);
    }
    const sampleRate = decoded.sampleRate;
    const numChannels = decoded.numberOfChannels;
    // Estimate total duration from file size using decoded bitrate
    const probeDuration = decoded.duration;
    const bytesPerSecond = probeSize / probeDuration;
    const estimatedDuration = file.size / bytesPerSecond;
    return { sampleRate, numChannels, estimatedDuration };
  } finally {
    audioCtx.close();
  }
}

async function sliceAudioFile(file, chunkDuration, onProgress) {
  // For small files (< 50 MB), use the simple in-memory approach
  if (file.size < 50 * 1024 * 1024) {
    return sliceAudioFileInMemory(file, chunkDuration, onProgress);
  }

  // For large files, slice by byte ranges to avoid loading everything into memory
  const { sampleRate, numChannels, estimatedDuration } = await probeAudioFile(file, onProgress);
  // Use dynamic chunk duration based on sample rate to stay under 25 MB
  const safeChunkDuration = chunkDuration || maxChunkDuration(sampleRate);
  const totalChunks = Math.ceil(estimatedDuration / safeChunkDuration);
  const bytesPerChunk = Math.floor(file.size / totalChunks);
  const chunks = [];
  const fileSizeMB = (file.size / (1024 * 1024)).toFixed(0);

  onProgress?.(`Splitting ${fileSizeMB} MB file into ${totalChunks} chunks (mono, under 24 MB each)…`);

  for (let i = 0; i < totalChunks; i++) {
    onProgress?.(`Decoding chunk ${i + 1} of ${totalChunks}…`);
    const byteStart = i * bytesPerChunk;
    const byteEnd = i === totalChunks - 1 ? file.size : (i + 1) * bytesPerChunk;
    const blob = file.slice(byteStart, byteEnd);
    const chunkStart = i * safeChunkDuration;

    // Decode this slice to re-encode as a clean WAV chunk
    const sliceBuf = await blob.arrayBuffer();
    const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    try {
      let decoded;
      try {
        decoded = await audioCtx.decodeAudioData(sliceBuf);
      } catch (e) {
        throw new Error(`Failed to decode chunk ${i + 1} of ${totalChunks} — the audio file may be corrupted at byte offset ${byteStart}`);
      }
      const channelData = [];
      for (let ch = 0; ch < decoded.numberOfChannels; ch++) {
        channelData.push(decoded.getChannelData(ch));
      }
      // Downmix to mono to reduce WAV size
      const mono = downmixToMono(channelData);
      onProgress?.(`Encoding chunk ${i + 1} of ${totalChunks} as mono WAV…`);
      const wavBlob = encodeWAV([mono], decoded.sampleRate, 1);
      const chunkFile = new File([wavBlob], `chunk_${i}.wav`, { type: "audio/wav" });
      chunks.push({ file: chunkFile, startOffset: chunkStart, index: i });
    } finally {
      audioCtx.close();
    }
  }

  return chunks;
}

async function sliceAudioFileInMemory(file, chunkDuration, onProgress) {
  const fileSizeMB = (file.size / (1024 * 1024)).toFixed(1);
  onProgress?.(`Reading ${fileSizeMB} MB audio file into memory…`);
  const arrayBuf = await file.arrayBuffer();
  const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  try {
    onProgress?.(`Decoding ${fileSizeMB} MB audio file…`);
    let decoded;
    try {
      decoded = await audioCtx.decodeAudioData(arrayBuf);
    } catch (e) {
      throw new Error(`Could not decode "${file.name}" — the file may be corrupted or in an unsupported format (try MP3, WAV, M4A, OGG, or FLAC)`);
    }
    const sampleRate = decoded.sampleRate;
    const totalSamples = decoded.length;
    // Use dynamic chunk duration based on sample rate to stay under 25 MB
    const safeChunkDuration = chunkDuration || maxChunkDuration(sampleRate);
    const chunkSamples = Math.floor(safeChunkDuration * sampleRate);
    const totalChunks = Math.ceil(totalSamples / chunkSamples);
    const chunks = [];

    onProgress?.(`Splitting into ${totalChunks} chunks (${Math.round(decoded.duration / 60)} min audio, ${sampleRate} Hz, mono)…`);

    for (let offset = 0; offset < totalSamples; offset += chunkSamples) {
      const chunkIndex = chunks.length;
      onProgress?.(`Encoding chunk ${chunkIndex + 1} of ${totalChunks} as mono WAV…`);
      const length = Math.min(chunkSamples, totalSamples - offset);
      const chunkStart = offset / sampleRate;
      const channelData = [];
      for (let ch = 0; ch < decoded.numberOfChannels; ch++) {
        const full = decoded.getChannelData(ch);
        channelData.push(full.slice(offset, offset + length));
      }
      // Downmix to mono to reduce WAV size
      const mono = downmixToMono(channelData);
      const wavBlob = encodeWAV([mono], sampleRate, 1);
      const chunkFile = new File([wavBlob], `chunk_${chunkIndex}.wav`, { type: "audio/wav" });
      chunks.push({ file: chunkFile, startOffset: chunkStart, index: chunkIndex });
    }

    return chunks;
  } finally {
    audioCtx.close();
  }
}

function encodeWAV(channelData, sampleRate, numChannels) {
  const numSamples = channelData[0].length;
  const bytesPerSample = 2; // 16-bit
  const dataSize = numSamples * numChannels * bytesPerSample;
  const buffer = new ArrayBuffer(44 + dataSize);
  const view = new DataView(buffer);

  const writeStr = (off, str) => { for (let i = 0; i < str.length; i++) view.setUint8(off + i, str.charCodeAt(i)); };
  writeStr(0, "RIFF");
  view.setUint32(4, 36 + dataSize, true);
  writeStr(8, "WAVE");
  writeStr(12, "fmt ");
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true); // PCM
  view.setUint16(22, numChannels, true);
  view.setUint32(24, sampleRate, true);
  view.setUint32(28, sampleRate * numChannels * bytesPerSample, true);
  view.setUint16(32, numChannels * bytesPerSample, true);
  view.setUint16(34, 16, true);
  writeStr(36, "data");
  view.setUint32(40, dataSize, true);

  let pos = 44;
  for (let i = 0; i < numSamples; i++) {
    for (let ch = 0; ch < numChannels; ch++) {
      const sample = Math.max(-1, Math.min(1, channelData[ch][i]));
      view.setInt16(pos, sample < 0 ? sample * 0x8000 : sample * 0x7FFF, true);
      pos += 2;
    }
  }
  return new Blob([buffer], { type: "audio/wav" });
}

/* ── Hardcoded API webhook base URL ── */
const WEBHOOK_BASE = "https://n8n.intelechia.com/webhook/oral-history";

/* ── Speaker Colors ── */
const SPEAKER_PALETTE = [
  { color: "#c9a55a", bg: "rgba(201,165,90,0.06)" },
  { color: "#5a8fc2", bg: "rgba(90,143,194,0.06)" },
  { color: "#c26a5a", bg: "rgba(194,106,90,0.06)" },
  { color: "#5ac28a", bg: "rgba(90,194,138,0.06)" },
  { color: "#b05ac2", bg: "rgba(176,90,194,0.06)" },
  { color: "#c2995a", bg: "rgba(194,153,90,0.06)" },
  { color: "#5ac2c2", bg: "rgba(90,194,194,0.06)" },
  { color: "#c25a8f", bg: "rgba(194,90,143,0.06)" },
];

/* ── Styles ─────────────────────────────────────────────────── */
const fonts = `@import url('https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=DM+Mono:wght@300;400;500&family=Newsreader:ital,opsz,wght@0,6..72,300;0,6..72,400;0,6..72,500;1,6..72,300;1,6..72,400&display=swap');`;

const C = {
  bg: "#faf8f5", surface: "#ffffff", raised: "#f2efe9", border: "#e0dbd3",
  text: "#1a1815", textDim: "#7a756c", accent: "#8b6914", accentDim: "rgba(139,105,20,0.08)",
  accentHover: "#a47d1a", danger: "#b34040",
  selected: "rgba(139,105,20,0.06)", selectedBorder: "rgba(139,105,20,0.3)",
  toolbar: "#f7f5f0",
};

/* 
  Data model:
  - "blocks" are the paragraph-level editing units
  - Each block: { id, speaker (id or null), segments: [...], wasMoved, wasEdited }
  - segments carry: { id, text, originalText, start, end, originalIndex }
  - Display merges segment texts into flowing paragraph
*/

/* ── Merge adjacent blocks with the same effective speaker into paragraphs ── */
function mergeAdjacentBlocks(blockList) {
  if (blockList.length <= 1) return blockList;
  const eff = blockList.map((b, i) => {
    if (b.speaker) return b.speaker;
    for (let j = i - 1; j >= 0; j--) { if (blockList[j].speaker) return blockList[j].speaker; }
    return null;
  });
  const bEnd = (b) => Math.max(...b.segments.map((s) => s.end));
  const bStart = (b) => Math.min(...b.segments.map((s) => s.start));
  const bText = (b) => b.segments.map((s) => s.text).join(" ");
  const result = [{ ...blockList[0], segments: [...blockList[0].segments] }];
  let resultEff = [eff[0]];
  for (let i = 1; i < blockList.length; i++) {
    // Never merge header blocks
    if (blockList[i].type === "header" || result[result.length - 1].type === "header") {
      result.push({ ...blockList[i], segments: [...blockList[i].segments] });
      resultEff.push(eff[i]);
      continue;
    }
    const prevEff = resultEff[resultEff.length - 1];
    const currEff = eff[i];
    const gap = bStart(blockList[i]) - bEnd(result[result.length - 1]);
    let sameGroup;
    if (prevEff !== currEff) {
      sameGroup = false;
    } else if (prevEff != null) {
      // Same attributed speaker — always merge
      sameGroup = true;
    } else {
      // Both unattributed — merge into paragraphs using sentence + pause heuristic.
      // Only break when the previous block ends a sentence AND there's a pause > 2s,
      // or when there's a very long pause (> 15s) regardless of punctuation.
      const prevText = bText(result[result.length - 1]).trim();
      const endsSentence = /[.!?]["''")\]]?\s*$/.test(prevText);
      sameGroup = gap < 15 && !(endsSentence && gap > 2);
    }
    if (sameGroup) {
      const prev = result[result.length - 1];
      prev.segments = [...prev.segments, ...blockList[i].segments];
      prev.wasMoved = prev.wasMoved || blockList[i].wasMoved;
      prev.wasEdited = prev.wasEdited || blockList[i].wasEdited;
    } else {
      result.push({ ...blockList[i], segments: [...blockList[i].segments] });
      resultEff.push(currEff);
    }
  }
  return result;
}

export default function OralHistoryEditor() {
  const [blocks, setBlocks] = useState([]);
  const [audioUrl, setAudioUrl] = useState(null);
  const [audioName, setAudioName] = useState("");
  const [subName, setSubName] = useState("");
  const [selected, setSelected] = useState(new Set());
  const [clipboard, setClipboard] = useState([]);
  const [editingId, setEditingId] = useState(null);
  const [editText, setEditText] = useState("");
  const [playing, setPlaying] = useState(null);
  const [undoStack, setUndoStack] = useState([]);
  const [dragOver, setDragOver] = useState(null);
  const [dragging, setDragging] = useState(null);

  // Speakers
  const [speakers, setSpeakers] = useState([]);
  const [addingSpeaker, setAddingSpeaker] = useState(false);
  const [newSpeakerName, setNewSpeakerName] = useState("");
  const [editingSpeakerId, setEditingSpeakerId] = useState(null);
  const [editingSpeakerName, setEditingSpeakerName] = useState("");
  const [showSpeakers, setShowSpeakers] = useState(true);

  // Export panel
  const [showExport, setShowExport] = useState(false);

  // Cloud sync & project
  const [showSettings, setShowSettings] = useState(false);
  const webhookBase = WEBHOOK_BASE;
  const [driveFileId, setDriveFileId] = useState("");
  const [syncStatus, setSyncStatus] = useState(null); // null | "saving" | "loading" | "saved" | "loaded" | "error"
  const [syncMsg, _setSyncMsg] = useState("");
  const syncMsgRef = useRef("");
  const setSyncMsg = useCallback((v) => { syncMsgRef.current = v; _setSyncMsg(v); }, []);
  const [projectTitle, setProjectTitle] = useState("Untitled Oral History");
  const [openaiApiKey, setOpenaiApiKey] = useState("");
  const [transcribing, setTranscribing] = useState(false);
  const [transcriptionProgress, setTranscriptionProgress] = useState(null); // { completed, total, segments }
  const [autoSaveEnabled, setAutoSaveEnabled] = useState(false);
  const autoSaveTimerRef = useRef(null);

  // Sections
  const [sections, setSections] = useState([]);
  const [showTOC, setShowTOC] = useState(true);
  const [activeSectionId, setActiveSectionId] = useState(null);
  const [addingSection, setAddingSection] = useState(false);
  const [newSectionName, setNewSectionName] = useState("");
  const [editingSectionId, setEditingSectionId] = useState(null);
  const [editingSectionName, setEditingSectionName] = useState("");

  // Original segments pool (for search/trim features)
  const [originalSegments, setOriginalSegments] = useState([]);

  // AI provider
  const [aiProvider, setAiProvider] = useState("anthropic");
  const [anthropicApiKey, setAnthropicApiKey] = useState("");
  const [aiAnalyzing, setAiAnalyzing] = useState(false);
  const [aiSuggestions, setAiSuggestions] = useState(null);
  const [aiSelectedSections, setAiSelectedSections] = useState(new Set());

  // Source tags
  const [sourceTags, setSourceTags] = useState({});

  // Edit history
  const [expandedEdit, setExpandedEdit] = useState(false);

  // Audio sync
  const [playingSegId, setPlayingSegId] = useState(null);
  const [currentTime, setCurrentTime] = useState(0);
  const [audioDuration, setAudioDuration] = useState(0);

  const audioRef = useRef(null);
  const audioFileRef = useRef(null);
  const playTimerRef = useRef(null);
  const newSpkRef = useRef(null);
  const editSpkRef = useRef(null);
  const projectUploadRef = useRef(null);
  const transcribeCancelRef = useRef(false);
  const newSectionRef = useRef(null);
  const editSectionRef = useRef(null);
  const editorContentRef = useRef(null);

  /* ── File handling ── */
  const handleAudioUpload = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    audioFileRef.current = f;
    setAudioName(f.name);
    const url = URL.createObjectURL(f);
    setAudioUrl(url);
    if (audioRef.current) audioRef.current.src = url;
  };

  const transcribeAudio = async (resume = false) => {
    if (!audioFileRef.current) {
      setSyncStatus("error"); setSyncMsg("Upload an audio file first");
      setTimeout(() => setSyncStatus(null), 3000); return;
    }
    if (!openaiApiKey.trim()) {
      setShowSettings(true);
      setSyncStatus("error"); setSyncMsg("Enter your OpenAI API key in Settings");
      setTimeout(() => setSyncStatus(null), 3000); return;
    }
    transcribeCancelRef.current = false;
    setTranscribing(true);
    const startTime = Date.now();
    const elapsed = () => {
      const s = Math.floor((Date.now() - startTime) / 1000);
      return s < 60 ? `${s}s` : `${Math.floor(s / 60)}m ${s % 60}s`;
    };
    setSyncStatus("saving"); setSyncMsg("Preparing audio…");
    const timer = setInterval(() => {
      if (syncMsgRef.current) {
        const base = syncMsgRef.current.replace(/ \(\d+[ms].*?\)$/, "");
        setSyncMsg(`${base} (${elapsed()})`);
      }
    }, 1000);

    try {
      // Split audio into chunks
      const chunks = await sliceAudioFile(audioFileRef.current, null, setSyncMsg);
      const totalChunks = chunks.length;

      // Determine where to resume from
      let startChunk = 0;
      let allSegs = [];
      if (resume && transcriptionProgress && transcriptionProgress.segments.length > 0) {
        startChunk = transcriptionProgress.completed;
        allSegs = [...transcriptionProgress.segments];
        setSyncMsg(`Resuming from chunk ${startChunk + 1} of ${totalChunks}…`);
      } else {
        setTranscriptionProgress({ completed: 0, total: totalChunks, segments: [] });
        setBlocks([]);
        setSelected(new Set());
        setUndoStack([]);
      }

      setSyncMsg(`Transcribing chunk 1 of ${totalChunks}…`);

      for (let i = startChunk; i < totalChunks; i++) {
        if (transcribeCancelRef.current) {
          setSyncMsg(`Paused after ${i} of ${totalChunks} chunks — ${allSegs.length} segments so far`);
          setSyncStatus("loaded");
          clearInterval(timer);
          setTranscribing(false);
          return;
        }

        const chunk = chunks[i];
        setSyncMsg(`Transcribing chunk ${i + 1} of ${totalChunks}…`);

        const formData = new FormData();
        formData.append("file", chunk.file);
        formData.append("model", "whisper-1");
        formData.append("response_format", "verbose_json");
        formData.append("timestamp_granularities[]", "segment");

        const res = await fetch("https://api.openai.com/v1/audio/transcriptions", {
          method: "POST",
          headers: { Authorization: `Bearer ${openaiApiKey.trim()}` },
          body: formData,
        });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.error?.message || `API error ${res.status} on chunk ${i + 1}`);
        }
        const data = await res.json();
        const chunkSegs = (data.segments || []).map((seg, j) => ({
          id: uid(),
          text: seg.text.trim(),
          originalText: seg.text.trim(),
          start: seg.start + chunk.startOffset,
          end: seg.end + chunk.startOffset,
          originalIndex: allSegs.length + j,
        }));

        allSegs = [...allSegs, ...chunkSegs];

        // Update blocks progressively so user sees results stream in
        const newBlocks = allSegs.map((seg) => ({
          id: uid(), speaker: null, segments: [seg], wasMoved: false, wasEdited: false, sectionId: null,
        }));
        setBlocks(mergeAdjacentBlocks(newBlocks));
        setTranscriptionProgress({ completed: i + 1, total: totalChunks, segments: allSegs });
        setSyncMsg(`Transcribed chunk ${i + 1}/${totalChunks} — ${allSegs.length} segments`);
      }

      if (allSegs.length === 0) throw new Error("No segments returned");
      setOriginalSegments(allSegs.map(s => ({...s})));
      setSubName("(transcribed)");
      setTranscriptionProgress(null);
      clearInterval(timer);
      setSyncStatus("loaded"); setSyncMsg(`Transcribed ${allSegs.length} segments in ${elapsed()}`);
      setTimeout(() => setSyncStatus(null), 4000);
    } catch (err) {
      clearInterval(timer);
      // Keep progress so user can resume
      setSyncStatus("error");
      const prog = transcriptionProgress;
      const resumeHint = prog && prog.completed > 0
        ? ` (${prog.completed}/${prog.total} chunks done — click Transcribe to resume)`
        : "";
      setSyncMsg(err.message + resumeHint);
      setTimeout(() => setSyncStatus(null), 8000);
    } finally {
      setTranscribing(false);
    }
  };

  const handleSubUpload = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setSubName(f.name);
    const reader = new FileReader();
    reader.onload = (ev) => {
      // Handle JSON project files as transcript imports
      if (f.name.endsWith(".json")) {
        try {
          const proj = JSON.parse(ev.target.result);
          if (proj.blocks && Array.isArray(proj.blocks)) {
            restoreProject(proj);
            return;
          }
        } catch (err) { /* not valid JSON, fall through to subtitle parsing */ }
      }
      const segments = detectAndParse(ev.target.result);
      setOriginalSegments(segments.map(s => ({...s})));
      const newBlocks = segments.map((seg) => ({
        id: uid(),
        speaker: null,
        segments: [seg],
        wasMoved: false,
        wasEdited: false,
        sectionId: null,
      }));
      setBlocks(mergeAdjacentBlocks(newBlocks));
      setSelected(new Set());
      setUndoStack([]);
    };
    reader.readAsText(f);
  };

  /* ── Load project from JSON file ── */
  const handleProjectUpload = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const proj = JSON.parse(ev.target.result);
        restoreProject(proj);
      } catch (err) {
        setSyncStatus("error");
        setSyncMsg("Invalid project file");
        setTimeout(() => setSyncStatus(null), 3000);
      }
    };
    reader.readAsText(f);
    e.target.value = "";
  };

  const restoreProject = (proj) => {
    if (!proj.blocks || !Array.isArray(proj.blocks)) return;
    if (proj.speakers && Array.isArray(proj.speakers)) {
      setSpeakers(proj.speakers.map((s) => ({
        id: s.id || uid(),
        name: s.name,
        color: s.color || SPEAKER_PALETTE[0].color,
        bg: SPEAKER_PALETTE.find((p) => p.color === s.color)?.bg || SPEAKER_PALETTE[0].bg,
      })));
    }
    if (proj.sections && Array.isArray(proj.sections)) {
      setSections(proj.sections.map(s => ({ id: s.id || uid(), name: s.name, collapsed: false })));
    }
    const spkNameToId = {};
    (proj.speakers || []).forEach((s) => { spkNameToId[s.name] = s.id || uid(); });
    const restored = proj.blocks.map((b) => ({
      id: uid(),
      speaker: b.speakerId || (b.speaker ? spkNameToId[b.speaker] : null) || null,
      wasMoved: b.wasMoved || false,
      wasEdited: b.wasEdited || false,
      sectionId: b.sectionId || null,
      ...(b.type === "header" ? { type: "header", text: b.headerText || b.text || "", level: b.level || 2, segments: [] } : {}),
      segments: (b.type === "header") ? [] : (b.segments || []).map((seg) => ({
        id: uid(),
        text: seg.text,
        originalText: seg.originalText || seg.text,
        start: seg.sourceStart,
        end: seg.sourceEnd,
        originalIndex: seg.originalIndex || 0,
      })),
    }));
    setBlocks(mergeAdjacentBlocks(restored));
    if (proj.sourceFile) setAudioName(proj.sourceFile);
    if (proj.title) setProjectTitle(proj.title);
    setSelected(new Set());
    setUndoStack([]);
    setSyncStatus("loaded");
    setSyncMsg("Project loaded");
    setTimeout(() => setSyncStatus(null), 2000);
  };

  /* ── Undo ── */
  const pushUndo = useCallback(() => {
    setUndoStack((prev) => [...prev.slice(-30), { blocks: JSON.parse(JSON.stringify(blocks)), sections: JSON.parse(JSON.stringify(sections)) }]);
  }, [blocks, sections]);

  const undo = useCallback(() => {
    if (undoStack.length === 0) return;
    const snapshot = undoStack[undoStack.length - 1];
    if (snapshot.blocks) {
      setBlocks(snapshot.blocks);
      if (snapshot.sections) setSections(snapshot.sections);
    } else {
      // Backwards compat with old undo entries that were just block arrays
      setBlocks(snapshot);
    }
    setUndoStack((s) => s.slice(0, -1));
    setSelected(new Set());
    setEditingId(null);
  }, [undoStack]);

  /* ── Derived data ── */
  const blockText = useCallback((block) => block.type === "header" ? (block.text || "") : block.segments.map((s) => s.text).join(" "), []);
  const blockStart = useCallback((block) => block.type === "header" ? 0 : Math.min(...block.segments.map((s) => s.start)), []);
  const blockEnd = useCallback((block) => block.type === "header" ? 0 : Math.max(...block.segments.map((s) => s.end)), []);
  const blockDuration = useCallback((block) => block.type === "header" ? 0 : block.segments.reduce((sum, s) => sum + (s.end - s.start), 0), []);

  const totalDuration = useMemo(() => blocks.reduce((sum, b) => sum + blockDuration(b), 0), [blocks, blockDuration]);

  const sectionedBlocks = useMemo(() => {
    const unsectioned = blocks.map((b, i) => ({ block: b, globalIdx: i })).filter(({ block }) => !block.sectionId);
    const grouped = sections.map(section => ({
      ...section,
      blocks: blocks.map((b, i) => ({ block: b, globalIdx: i })).filter(({ block }) => block.sectionId === section.id),
      duration: blocks.filter(b => b.sectionId === section.id).reduce((sum, b) => sum + b.segments.reduce((s, seg) => s + (seg.end - seg.start), 0), 0),
    }));
    return { unsectioned, sections: grouped };
  }, [blocks, sections]);

  const getSpeaker = useCallback((id) => speakers.find((s) => s.id === id) || null, [speakers]);

  // Resolve effective speaker: if a block has no explicit speaker, inherit from the nearest prior block that does
  const getEffectiveSpeakerId = useCallback((blockIdx) => {
    for (let i = blockIdx; i >= 0; i--) {
      if (blocks[i].speaker) return blocks[i].speaker;
    }
    return null;
  }, [blocks]);

  const getEffectiveSpeaker = useCallback((blockIdx) => {
    const sid = getEffectiveSpeakerId(blockIdx);
    return sid ? getSpeaker(sid) : null;
  }, [getEffectiveSpeakerId, getSpeaker]);

  /* ── Speaker management ── */
  const addSpeaker = useCallback((name) => {
    if (!name.trim()) return;
    const pal = SPEAKER_PALETTE[speakers.length % SPEAKER_PALETTE.length];
    setSpeakers((prev) => [...prev, { id: uid(), name: name.trim(), ...pal }]);
    setNewSpeakerName("");
    setAddingSpeaker(false);
  }, [speakers]);

  const removeSpeaker = useCallback((sid) => {
    setSpeakers((prev) => prev.filter((s) => s.id !== sid));
    setBlocks((prev) => prev.map((b) => b.speaker === sid ? { ...b, speaker: null } : b));
  }, []);

  const renameSpeaker = useCallback((sid, name) => {
    if (!name.trim()) return;
    setSpeakers((prev) => prev.map((s) => s.id === sid ? { ...s, name: name.trim() } : s));
    setEditingSpeakerId(null);
  }, []);

  const assignSpeaker = useCallback((sid) => {
    if (selected.size === 0) return;
    pushUndo();
    setBlocks((prev) => prev.map((b) =>
      selected.has(b.id) ? { ...b, speaker: b.speaker === sid ? null : sid } : b
    ));
  }, [selected, pushUndo]);

  /* ── Auto-merge adjacent same-speaker blocks ── */
  const mergeAdjacent = useCallback((blockList) => mergeAdjacentBlocks(blockList), []);

  /* After speaker assignment, auto-merge */
  const assignAndMerge = useCallback((sid) => {
    if (selected.size === 0) return;
    pushUndo();
    const updated = blocks.map((b) =>
      selected.has(b.id) ? { ...b, speaker: b.speaker === sid ? null : sid } : b
    );
    setBlocks(mergeAdjacent(updated));
    setSelected(new Set());
  }, [selected, pushUndo, blocks, mergeAdjacent]);

  /* ── Section management ── */
  const addSection = useCallback((name) => {
    if (!name.trim()) return;
    setSections(prev => [...prev, { id: uid(), name: name.trim(), collapsed: false }]);
    setNewSectionName("");
    setAddingSection(false);
  }, []);

  const removeSection = useCallback((sid) => {
    pushUndo();
    setSections(prev => prev.filter(s => s.id !== sid));
    setBlocks(prev => prev.map(b => b.sectionId === sid ? { ...b, sectionId: null } : b));
  }, [pushUndo]);

  const renameSection = useCallback((sid, name) => {
    if (!name.trim()) return;
    setSections(prev => prev.map(s => s.id === sid ? { ...s, name: name.trim() } : s));
    setEditingSectionId(null);
  }, []);

  const toggleSectionCollapse = useCallback((sid) => {
    setSections(prev => prev.map(s => s.id === sid ? { ...s, collapsed: !s.collapsed } : s));
  }, []);

  /* ── Section assignment ── */
  const assignToSection = useCallback((sectionId) => {
    if (selected.size === 0) return;
    pushUndo();
    setBlocks(prev => prev.map(b => selected.has(b.id) ? { ...b, sectionId } : b));
    setSelected(new Set());
  }, [selected, pushUndo]);

  const unsectionBlocks = useCallback(() => {
    if (selected.size === 0) return;
    pushUndo();
    setBlocks(prev => prev.map(b => selected.has(b.id) ? { ...b, sectionId: null } : b));
    setSelected(new Set());
  }, [selected, pushUndo]);

  /* ── Selection ── */
  const handleSelect = useCallback((id, e) => {
    if (editingId) return;
    if (e.shiftKey && selected.size > 0) {
      const ids = blocks.map((b) => b.id);
      const lastSel = [...selected].pop();
      const a = ids.indexOf(lastSel), b = ids.indexOf(id);
      const [lo, hi] = a < b ? [a, b] : [b, a];
      setSelected(new Set(ids.slice(lo, hi + 1)));
    } else if (e.metaKey || e.ctrlKey) {
      setSelected((prev) => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
    } else {
      setSelected(new Set([id]));
    }
  }, [blocks, selected, editingId]);

  /* ── Clipboard ── */
  const doCopy = useCallback(() => {
    if (selected.size === 0) return;
    setClipboard(blocks.filter((b) => selected.has(b.id)).map((b) => JSON.parse(JSON.stringify(b))));
  }, [blocks, selected]);

  const doCut = useCallback(() => {
    if (selected.size === 0) return;
    pushUndo();
    setClipboard(blocks.filter((b) => selected.has(b.id)).map((b) => JSON.parse(JSON.stringify(b))));
    setBlocks((prev) => prev.filter((b) => !selected.has(b.id)));
    setSelected(new Set());
  }, [blocks, selected, pushUndo]);

  const doPaste = useCallback(() => {
    if (clipboard.length === 0) return;
    pushUndo();
    const pasted = clipboard.map((b) => ({ ...JSON.parse(JSON.stringify(b)), id: uid(), wasMoved: true }));
    const ids = blocks.map((b) => b.id);
    const lastSel = [...selected].pop();
    const insertIdx = lastSel ? ids.indexOf(lastSel) + 1 : blocks.length;
    const next = [...blocks];
    next.splice(insertIdx, 0, ...pasted);
    setBlocks(next);
    setSelected(new Set(pasted.map((b) => b.id)));
  }, [clipboard, blocks, selected, pushUndo]);

  const doDelete = useCallback(() => {
    if (selected.size === 0) return;
    pushUndo();
    setBlocks((prev) => prev.filter((b) => !selected.has(b.id)));
    setSelected(new Set());
  }, [selected, pushUndo]);

  /* ── Merge selected blocks ── */
  const mergeSelected = useCallback(() => {
    if (selected.size < 2) return;
    pushUndo();
    const sel = blocks.filter((b) => selected.has(b.id));
    const merged = {
      id: uid(),
      speaker: sel[0].speaker,
      segments: sel.flatMap((b) => b.segments),
      wasMoved: sel.some((b) => b.wasMoved),
      wasEdited: true,
      sectionId: sel[0].sectionId || null,
    };
    const firstIdx = blocks.findIndex((b) => selected.has(b.id));
    const next = blocks.filter((b) => !selected.has(b.id));
    next.splice(firstIdx, 0, merged);
    setBlocks(next);
    setSelected(new Set([merged.id]));
  }, [blocks, selected, pushUndo]);

  /* ── Split block at cursor ── */
  const splitBlock = useCallback((blockId, charPos) => {
    pushUndo();
    setBlocks((prev) => {
      const idx = prev.findIndex((b) => b.id === blockId);
      if (idx === -1) return prev;
      const block = prev[idx];
      const fullText = block.segments.map((s) => s.text).join(" ");
      if (charPos <= 0 || charPos >= fullText.length) return prev;

      // Find which segment and where to split
      let running = 0;
      let splitSegIdx = 0;
      let splitInSeg = 0;
      for (let i = 0; i < block.segments.length; i++) {
        const segEnd = running + block.segments[i].text.length + (i > 0 ? 1 : 0);
        if (charPos <= segEnd) {
          splitSegIdx = i;
          splitInSeg = charPos - running - (i > 0 ? 1 : 0);
          break;
        }
        running = segEnd;
      }

      const segs1 = block.segments.slice(0, splitSegIdx);
      const segs2 = block.segments.slice(splitSegIdx + 1);
      const splitSeg = block.segments[splitSegIdx];

      if (splitInSeg > 0 && splitInSeg < splitSeg.text.length) {
        const ratio = splitInSeg / splitSeg.text.length;
        const mid = splitSeg.start + (splitSeg.end - splitSeg.start) * ratio;
        segs1.push({ ...splitSeg, id: uid(), text: splitSeg.text.slice(0, splitInSeg).trim(), end: mid });
        segs2.unshift({ ...splitSeg, id: uid(), text: splitSeg.text.slice(splitInSeg).trim(), start: mid });
      } else if (splitInSeg <= 0) {
        segs2.unshift(splitSeg);
      } else {
        segs1.push(splitSeg);
      }

      if (segs1.length === 0 || segs2.length === 0) return prev;

      const b1 = { id: uid(), speaker: block.speaker, segments: segs1, wasMoved: block.wasMoved, wasEdited: true, sectionId: block.sectionId || null };
      const b2 = { id: uid(), speaker: block.speaker, segments: segs2, wasMoved: block.wasMoved, wasEdited: true, sectionId: block.sectionId || null };
      const next = [...prev];
      next.splice(idx, 1, b1, b2);
      return next;
    });
  }, [pushUndo]);

  /* ── Split by sentence ── */
  const splitBySentence = useCallback((blockId) => {
    pushUndo();
    setBlocks(prev => {
      const idx = prev.findIndex(b => b.id === blockId);
      if (idx === -1) return prev;
      const block = prev[idx];
      const fullText = block.segments.map(s => s.text).join(" ");
      // Find sentence boundaries
      const sentences = [];
      let last = 0;
      const regex = /[.!?]+[\s"')\]]*(?=[A-Z])/g;
      let match;
      while ((match = regex.exec(fullText)) !== null) {
        sentences.push(fullText.slice(last, match.index + match[0].length).trim());
        last = match.index + match[0].length;
      }
      if (last < fullText.length) sentences.push(fullText.slice(last).trim());
      if (sentences.length <= 1) return prev;

      // Create new blocks for each sentence
      const totalDur = block.segments.reduce((s, seg) => s + (seg.end - seg.start), 0);
      const startTime = Math.min(...block.segments.map(s => s.start));
      const newBlocks = sentences.filter(s => s.length > 0).map((sentence, i) => {
        const charRatio = sentence.length / fullText.length;
        const segStart = startTime + (totalDur * (sentences.slice(0, i).join("").length / fullText.length));
        const segEnd = segStart + totalDur * charRatio;
        return {
          id: uid(), speaker: block.speaker, sectionId: block.sectionId || null,
          wasMoved: block.wasMoved, wasEdited: true,
          segments: [{ id: uid(), text: sentence, originalText: sentence, start: segStart, end: segEnd, originalIndex: 0 }],
        };
      });
      const next = [...prev];
      next.splice(idx, 1, ...newBlocks);
      return next;
    });
  }, [pushUndo]);

  /* ── Split at largest pause ── */
  const splitAtPause = useCallback((blockId) => {
    pushUndo();
    setBlocks(prev => {
      const idx = prev.findIndex(b => b.id === blockId);
      if (idx === -1) return prev;
      const block = prev[idx];
      if (block.segments.length < 2) return prev;
      const sorted = [...block.segments].sort((a, b) => a.start - b.start);
      let maxGap = 0, maxGapIdx = 0;
      for (let i = 1; i < sorted.length; i++) {
        const gap = sorted[i].start - sorted[i - 1].end;
        if (gap > maxGap) { maxGap = gap; maxGapIdx = i; }
      }
      if (maxGap <= 0) return prev;
      const segs1 = sorted.slice(0, maxGapIdx);
      const segs2 = sorted.slice(maxGapIdx);
      const b1 = { id: uid(), speaker: block.speaker, segments: segs1, wasMoved: block.wasMoved, wasEdited: true, sectionId: block.sectionId || null };
      const b2 = { id: uid(), speaker: block.speaker, segments: segs2, wasMoved: block.wasMoved, wasEdited: true, sectionId: block.sectionId || null };
      const next = [...prev];
      next.splice(idx, 1, b1, b2);
      return next;
    });
  }, [pushUndo]);

  /* ── Decompose block to individual segments ── */
  const decomposeBlock = useCallback((blockId) => {
    pushUndo();
    setBlocks(prev => {
      const idx = prev.findIndex(b => b.id === blockId);
      if (idx === -1) return prev;
      const block = prev[idx];
      if (block.segments.length <= 1) return prev;
      const newBlocks = block.segments.map(seg => ({
        id: uid(), speaker: block.speaker, segments: [{ ...seg, id: uid() }],
        wasMoved: block.wasMoved, wasEdited: false, sectionId: block.sectionId || null,
      }));
      const next = [...prev];
      next.splice(idx, 1, ...newBlocks);
      return next;
    });
  }, [pushUndo]);

  /* ── Header blocks ── */
  const addHeaderBlock = useCallback(() => {
    pushUndo();
    const header = {
      id: uid(), type: "header", text: "New Section", level: 2,
      speaker: null, sectionId: null, segments: [],
      wasMoved: false, wasEdited: false,
    };
    // Insert before first selected block, or at beginning
    const ids = blocks.map(b => b.id);
    const firstSel = [...selected].find(id => ids.includes(id));
    const insertIdx = firstSel ? ids.indexOf(firstSel) : 0;
    const next = [...blocks];
    next.splice(insertIdx, 0, header);
    setBlocks(next);
    setSelected(new Set([header.id]));
    setEditingId(header.id);
    setEditText("New Section");
  }, [blocks, selected, pushUndo]);

  /* ── Inline text edit ── */
  const startEdit = (block) => { setEditingId(block.id); setEditText(block.type === "header" ? (block.text || "") : blockText(block)); };

  const commitEdit = useCallback(() => {
    if (!editingId) return;
    pushUndo();
    // Check if editing a header block
    const editingBlock = blocks.find(b => b.id === editingId);
    if (editingBlock?.type === "header") {
      setBlocks(prev => prev.map(b => b.id !== editingId ? b : { ...b, text: editText }));
      setEditingId(null);
      setEditText("");
      return;
    }
    setBlocks((prev) => prev.map((b) => {
      if (b.id !== editingId) return b;
      // Redistribute edited text across segments proportionally by duration
      const totalChars = editText.length;
      const totalDur = b.segments.reduce((s, seg) => s + (seg.end - seg.start), 0);
      let charPos = 0;
      const newSegs = b.segments.map((seg, i) => {
        const segDurRatio = totalDur > 0 ? (seg.end - seg.start) / totalDur : 1 / b.segments.length;
        const chars = i === b.segments.length - 1
          ? totalChars - charPos
          : Math.round(totalChars * segDurRatio);
        const newText = editText.slice(charPos, charPos + chars).trim();
        charPos += chars;
        return { ...seg, text: newText || seg.text };
      }).filter((s) => s.text.length > 0);

      const origText = b.segments.map((s) => s.originalText).join(" ");
      return { ...b, segments: newSegs.length > 0 ? newSegs : b.segments, wasEdited: editText !== origText };
    }));
    setEditingId(null);
    setEditText("");
  }, [editingId, editText, pushUndo]);

  const cancelEdit = () => { setEditingId(null); setEditText(""); };

  /* ── Drag reorder ── */
  const handleDrop = useCallback((targetId) => {
    if (!dragging || dragging === targetId) { setDragging(null); setDragOver(null); return; }
    pushUndo();
    const from = blocks.findIndex((b) => b.id === dragging);
    const to = blocks.findIndex((b) => b.id === targetId);
    const next = [...blocks];
    const [moved] = next.splice(from, 1);
    moved.wasMoved = true;
    // Cross-section drag: adopt target block's sectionId
    const target = next.find(b => b.id === targetId);
    if (target) moved.sectionId = target.sectionId || null;
    next.splice(to, 0, moved);
    setBlocks(next);
    setDragging(null);
    setDragOver(null);
  }, [dragging, blocks, pushUndo]);

  /* ── Playback ── */
  const playBlock = useCallback((block) => {
    if (!audioRef.current || !audioUrl) return;
    clearInterval(playTimerRef.current);
    const segs = [...block.segments].sort((a, b) => a.start - b.start);
    let idx = 0;
    const playNext = () => {
      if (idx >= segs.length) { setPlaying(null); return; }
      const seg = segs[idx];
      const a = audioRef.current;
      a.currentTime = seg.start;
      a.play();
      setPlaying(block.id);
      playTimerRef.current = setInterval(() => {
        if (a.currentTime >= seg.end) {
          clearInterval(playTimerRef.current);
          idx++;
          playNext();
        }
      }, 50);
    };
    playNext();
  }, [audioUrl]);

  const playAll = useCallback(() => {
    if (!audioRef.current || !audioUrl || blocks.length === 0) return;
    let bIdx = 0;
    const playNextBlock = () => {
      if (bIdx >= blocks.length) { setPlaying(null); return; }
      const block = blocks[bIdx];
      const segs = [...block.segments].sort((a, b) => a.start - b.start);
      let sIdx = 0;
      const playNextSeg = () => {
        if (sIdx >= segs.length) { bIdx++; playNextBlock(); return; }
        const seg = segs[sIdx];
        audioRef.current.currentTime = seg.start;
        audioRef.current.play();
        setPlaying(block.id);
        clearInterval(playTimerRef.current);
        playTimerRef.current = setInterval(() => {
          if (audioRef.current.currentTime >= seg.end) {
            clearInterval(playTimerRef.current);
            sIdx++;
            playNextSeg();
          }
        }, 50);
      };
      playNextSeg();
    };
    playNextBlock();
  }, [audioUrl, blocks]);

  const stopPlayback = () => {
    clearInterval(playTimerRef.current);
    audioRef.current?.pause();
    setPlaying(null);
  };

  // Audio time tracking for scrub bar
  useEffect(() => {
    const audio = audioRef.current;
    if (!audio) return;
    const onTimeUpdate = () => setCurrentTime(audio.currentTime);
    const onLoadedMetadata = () => setAudioDuration(audio.duration || 0);
    const onDurationChange = () => setAudioDuration(audio.duration || 0);
    audio.addEventListener("timeupdate", onTimeUpdate);
    audio.addEventListener("loadedmetadata", onLoadedMetadata);
    audio.addEventListener("durationchange", onDurationChange);
    return () => {
      audio.removeEventListener("timeupdate", onTimeUpdate);
      audio.removeEventListener("loadedmetadata", onLoadedMetadata);
      audio.removeEventListener("durationchange", onDurationChange);
    };
  }, []);

  /* ── Keyboard ── */
  useEffect(() => {
    const handler = (e) => {
      if (addingSpeaker || editingSpeakerId) return;
      if (editingId && e.key === "Escape") { cancelEdit(); return; }
      if (editingId) return;

      const mod = e.metaKey || e.ctrlKey;
      if (mod && e.key === "c") { e.preventDefault(); doCopy(); }
      if (mod && e.key === "x") { e.preventDefault(); doCut(); }
      if (mod && e.key === "v") { e.preventDefault(); doPaste(); }
      if (mod && e.key === "z") { e.preventDefault(); undo(); }
      if (mod && e.key === "a") { e.preventDefault(); setSelected(new Set(blocks.map((b) => b.id))); }
      if ((e.key === "Delete" || e.key === "Backspace") && selected.size > 0) { e.preventDefault(); doDelete(); }

      if (selected.size > 0 && !mod && e.key >= "1" && e.key <= "9") {
        const idx = parseInt(e.key) - 1;
        if (idx < speakers.length) { e.preventDefault(); assignAndMerge(speakers[idx].id); }
      }
      if (selected.size > 0 && !mod && e.key === "0") {
        e.preventDefault();
        pushUndo();
        setBlocks((prev) => prev.map((b) => selected.has(b.id) ? { ...b, speaker: null } : b));
      }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [editingId, doCopy, doCut, doPaste, doDelete, undo, selected, speakers, assignAndMerge, addingSpeaker, editingSpeakerId, blocks, pushUndo]);

  useEffect(() => { if (addingSpeaker && newSpkRef.current) newSpkRef.current.focus(); }, [addingSpeaker]);
  useEffect(() => { if (editingSpeakerId && editSpkRef.current) editSpkRef.current.focus(); }, [editingSpeakerId]);
  useEffect(() => { if (addingSection && newSectionRef.current) newSectionRef.current.focus(); }, [addingSection]);
  useEffect(() => { if (editingSectionId && editSectionRef.current) editSectionRef.current.focus(); }, [editingSectionId]);

  /* ── Full project state for save ── */
  const buildProjectState = useCallback(() => {
    const speakerMap = {};
    speakers.forEach((s) => { speakerMap[s.id] = s.name; });
    return {
      version: 2,
      title: projectTitle,
      sourceFile: audioName || "source.wav",
      totalDuration: +(blocks.reduce((sum, b) => sum + b.segments.reduce((s2, seg) => s2 + (seg.end - seg.start), 0), 0)).toFixed(3),
      speakers: speakers.map((s) => ({ id: s.id, name: s.name, color: s.color })),
      sections: sections.map((s, i) => ({ id: s.id, name: s.name, order: i + 1 })),
      blocks: blocks.map((b, i) => {
        let effId = b.speaker;
        if (!effId) { for (let j = i - 1; j >= 0; j--) { if (blocks[j].speaker) { effId = blocks[j].speaker; break; } } }
        return {
          order: i + 1,
          speaker: effId ? (speakerMap[effId] || null) : null,
          speakerId: b.speaker,
          explicitSpeaker: !!b.speaker,
          sectionId: b.sectionId || null,
          type: b.type || "paragraph",
          ...(b.type === "header" ? { headerText: b.text, level: b.level || 2 } : {}),
          editHistory: b.editHistory || [],
          text: b.segments.map((s) => s.text).join(" "),
          wasEdited: b.wasEdited,
          wasMoved: b.wasMoved,
          segments: b.segments.map((seg) => ({
            text: seg.text, originalText: seg.originalText,
            sourceStart: +seg.start.toFixed(3), sourceEnd: +seg.end.toFixed(3),
            duration: +((seg.end - seg.start).toFixed(3)),
            originalIndex: seg.originalIndex,
          })),
        };
      }),
    };
  }, [blocks, speakers, sections, audioName, projectTitle]);

  /* ── Cloud sync via n8n webhooks ── */
  const cloudSave = useCallback(async () => {
    if (!webhookBase) { setSyncStatus("error"); setSyncMsg("Set webhook URL in settings"); setTimeout(() => setSyncStatus(null), 3000); return; }
    setSyncStatus("saving");
    setSyncMsg("Saving to Google Drive…");
    const proj = buildProjectState();
    try {
      const isUpdate = !!driveFileId;
      const url = isUpdate
        ? `${webhookBase.replace(/\/$/, "")}-update`
        : webhookBase;
      const body = {
        fileName: `${projectTitle.replace(/[^a-zA-Z0-9_ -]/g, "")}.json`,
        content: JSON.stringify(proj, null, 2),
        ...(isUpdate ? { fileId: driveFileId } : {}),
      };
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const data = await res.json();
      if (data.id) setDriveFileId(data.id);
      setSyncStatus("saved");
      setSyncMsg("Saved to Google Drive");
      setTimeout(() => setSyncStatus(null), 3000);
    } catch (err) {
      setSyncStatus("error");
      setSyncMsg("Save failed: " + err.message);
      setTimeout(() => setSyncStatus(null), 4000);
    }
  }, [webhookBase, driveFileId, buildProjectState, projectTitle]);

  const cloudLoad = useCallback(async () => {
    if (!webhookBase || !driveFileId) { setSyncStatus("error"); setSyncMsg("Set webhook URL and file ID in settings"); setTimeout(() => setSyncStatus(null), 3000); return; }
    setSyncStatus("loading");
    setSyncMsg("Loading from Google Drive…");
    try {
      const url = `${webhookBase.replace(/\/$/, "")}-load?fileId=${driveFileId}`;
      const res = await fetch(url);
      const data = await res.json();
      if (data.content) {
        const proj = typeof data.content === "string" ? JSON.parse(data.content) : data.content;
        restoreProject(proj);
      } else {
        restoreProject(data);
      }
      setSyncStatus("loaded");
      setSyncMsg("Loaded from Google Drive");
      setTimeout(() => setSyncStatus(null), 3000);
    } catch (err) {
      setSyncStatus("error");
      setSyncMsg("Load failed: " + err.message);
      setTimeout(() => setSyncStatus(null), 4000);
    }
  }, [webhookBase, driveFileId]);

  /* ── AI Analysis ── */
  const callAI = useCallback(async (systemPrompt, userContent) => {
    if (aiProvider === "anthropic") {
      if (!anthropicApiKey.trim()) throw new Error("Enter your Anthropic API key in Settings");
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": anthropicApiKey.trim(),
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: "claude-haiku-4-5-20251001",
          max_tokens: 4096,
          messages: [{ role: "user", content: systemPrompt + "\n\n" + userContent }],
        }),
      });
      if (!res.ok) { const err = await res.json().catch(() => ({})); throw new Error(err.error?.message || `Anthropic API error ${res.status}`); }
      const data = await res.json();
      const text = data.content?.[0]?.text || "";
      const jsonMatch = text.match(/\{[\s\S]*\}/);
      if (!jsonMatch) throw new Error("AI response did not contain valid JSON");
      return JSON.parse(jsonMatch[0]);
    } else {
      if (!openaiApiKey.trim()) throw new Error("Enter your OpenAI API key in Settings");
      const res = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${openaiApiKey.trim()}` },
        body: JSON.stringify({
          model: "gpt-4o-mini",
          messages: [{ role: "system", content: systemPrompt }, { role: "user", content: userContent }],
          response_format: { type: "json_object" },
          temperature: 0.3,
        }),
      });
      if (!res.ok) { const err = await res.json().catch(() => ({})); throw new Error(err.error?.message || `OpenAI API error ${res.status}`); }
      const data = await res.json();
      return JSON.parse(data.choices[0].message.content);
    }
  }, [aiProvider, anthropicApiKey, openaiApiKey]);

  const analyzeTranscript = useCallback(async () => {
    if (blocks.length === 0) return;
    setAiAnalyzing(true);
    setSyncStatus("saving");
    setSyncMsg("Analyzing transcript with AI...");
    try {
      // Build condensed transcript
      const condensed = blocks.map((b, i) => {
        const spk = speakers.find(s => s.id === b.speaker)?.name || "";
        const text = b.segments.map(s => s.text).join(" ").slice(0, 80);
        return `[${i}] ${spk ? spk + ": " : ""}${text}...`;
      }).join("\n");

      const systemPrompt = `You are analyzing an oral history transcript to identify natural topic sections and tags. Return a JSON object with:
- "sections": array of { "name": string, "startBlock": number, "endBlock": number } identifying natural topic/chapter boundaries
- "tags": array of { "startBlock": number, "endBlock": number, "tags": string[] } identifying key topics discussed in block ranges

Keep section names concise (3-6 words). Identify 3-10 sections. Tag with 1-3 relevant topic keywords per range.
Return ONLY valid JSON, no other text.`;

      const result = await callAI(systemPrompt, condensed);
      setAiSuggestions(result);
      setSyncStatus("loaded");
      setSyncMsg("AI analysis complete — review suggestions");
      setTimeout(() => setSyncStatus(null), 3000);
    } catch (err) {
      setSyncStatus("error");
      setSyncMsg("AI analysis failed: " + err.message);
      setTimeout(() => setSyncStatus(null), 5000);
    } finally {
      setAiAnalyzing(false);
    }
  }, [blocks, speakers, callAI]);

  const applyAiSuggestions = useCallback((selectedSections) => {
    if (!aiSuggestions?.sections) return;
    pushUndo();
    const newSections = [];
    const blockUpdates = {};
    selectedSections.forEach(idx => {
      const suggestion = aiSuggestions.sections[idx];
      if (!suggestion) return;
      const secId = uid();
      newSections.push({ id: secId, name: suggestion.name, collapsed: false });
      for (let i = suggestion.startBlock; i <= Math.min(suggestion.endBlock, blocks.length - 1); i++) {
        blockUpdates[blocks[i].id] = secId;
      }
    });
    setSections(prev => [...prev, ...newSections]);
    setBlocks(prev => prev.map(b => blockUpdates[b.id] ? { ...b, sectionId: blockUpdates[b.id] } : b));
    setAiSuggestions(null);
  }, [aiSuggestions, blocks, pushUndo]);

  /* ── Periodic auto-save ── */
  const AUTO_SAVE_INTERVAL = 3 * 60 * 1000; // 3 minutes
  useEffect(() => {
    if (autoSaveTimerRef.current) clearInterval(autoSaveTimerRef.current);
    if (autoSaveEnabled && webhookBase && blocks.length > 0) {
      autoSaveTimerRef.current = setInterval(() => {
        // Don't auto-save while actively transcribing or during another sync
        if (syncStatus === "saving" || syncStatus === "loading") return;
        cloudSave();
      }, AUTO_SAVE_INTERVAL);
    }
    return () => { if (autoSaveTimerRef.current) clearInterval(autoSaveTimerRef.current); };
  }, [autoSaveEnabled, webhookBase, blocks.length, cloudSave, syncStatus]);

  /* ── Export: Markdown ── */
  const generateMarkdown = useCallback((sectionFilter) => {
    let md = `# ${projectTitle}\n\n`;
    if (audioName) md += `*Source: ${audioName}*\n\n---\n\n`;

    const renderBlocks = (blockList, blockIndices) => {
      let out = "";
      let lastSpeakerName = null;
      blockList.forEach((block, i) => {
        if (block.type === "header") {
          const prefix = "#".repeat(block.level || 2);
          md += `${prefix} ${block.text || "Untitled"}\n\n`;
          return;
        }
        const idx = blockIndices ? blockIndices[i] : blocks.indexOf(block);
        const spk = getEffectiveSpeaker(idx);
        const name = spk ? spk.name : "Unknown Speaker";
        const text = blockText(block);
        if (name !== lastSpeakerName) {
          out += `**${name}:** ${text}\n\n`;
        } else {
          out += `${text}\n\n`;
        }
        lastSpeakerName = name;
      });
      return out;
    };

    if (sections.length === 0 || sectionFilter) {
      // Flat or single-section export
      const filteredBlocks = sectionFilter
        ? blocks.filter(b => b.sectionId === sectionFilter)
        : blocks;
      const filteredIndices = sectionFilter
        ? blocks.map((b, i) => b.sectionId === sectionFilter ? i : -1).filter(i => i !== -1)
        : null;
      md += renderBlocks(filteredBlocks, filteredIndices);
    } else {
      // Section-grouped export
      const unsectioned = blocks.filter(b => !b.sectionId);
      if (unsectioned.length > 0) {
        md += renderBlocks(unsectioned, unsectioned.map(b => blocks.indexOf(b)));
        md += "---\n\n";
      }
      sections.forEach(section => {
        const sBlocks = blocks.filter(b => b.sectionId === section.id);
        if (sBlocks.length > 0) {
          md += `## ${section.name}\n\n`;
          md += renderBlocks(sBlocks, sBlocks.map(b => blocks.indexOf(b)));
        }
      });
    }
    return md;
  }, [blocks, sections, getEffectiveSpeaker, blockText, audioName, projectTitle]);

  const downloadMarkdown = useCallback((sectionFilter) => {
    const md = generateMarkdown(sectionFilter);
    const blob = new Blob([md], { type: "text/markdown" });
    const url = URL.createObjectURL(blob);
    let fn = projectTitle.replace(/[^a-zA-Z0-9_ -]/g, "").replace(/\s+/g, "-").toLowerCase() || "oral-history";
    if (sectionFilter) {
      const sec = sections.find(s => s.id === sectionFilter);
      if (sec) fn += "-" + sec.name.replace(/[^a-zA-Z0-9_ -]/g, "").replace(/\s+/g, "-").toLowerCase();
    }
    const a = document.createElement("a"); a.href = url; a.download = `${fn}.md`; a.click();
    URL.revokeObjectURL(url);
  }, [generateMarkdown, projectTitle, sections]);

  /* ── Export: Project JSON ── */
  const generateProject = useCallback(() => {
    const proj = buildProjectState();
    proj.ffmpeg = blocks.flatMap((b) => b.segments).map((s, i) =>
      `ffmpeg -ss ${s.start.toFixed(3)} -to ${s.end.toFixed(3)} -i "$INPUT" -c copy /tmp/seg_${String(i).padStart(4, "0")}.wav`
    ).join("\n") + `\nprintf "file '%s'\\n" /tmp/seg_*.wav > /tmp/list.txt\nffmpeg -f concat -safe 0 -i /tmp/list.txt -c copy output.wav`;
    return proj;
  }, [buildProjectState, blocks]);

  const downloadProject = useCallback(() => {
    const blob = new Blob([JSON.stringify(generateProject(), null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const fn = projectTitle.replace(/[^a-zA-Z0-9_ -]/g, "").replace(/\s+/g, "-").toLowerCase() || "oral-history-project";
    const a = document.createElement("a"); a.href = url; a.download = `${fn}.json`; a.click();
    URL.revokeObjectURL(url);
  }, [generateProject, projectTitle]);

  /* ── Render ── */
  const hasContent = blocks.length > 0;

  return (
    <div style={{ fontFamily: "'Newsreader', Georgia, serif", background: C.bg, color: C.text, minHeight: "100vh", display: "flex", flexDirection: "column" }}>
      <style>{fonts + `
        html, body, #root { height: 100%; margin: 0; }
        *, *::before, *::after { box-sizing: border-box; margin: 0; }
        ::selection { background: rgba(139,105,20,0.15); }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 3px; }
        textarea:focus, input:focus { outline: none; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }
        @keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:0.4; } }
      `}</style>
      <audio ref={audioRef} preload="auto" />

      {/* ── Header ── */}
      <header style={{ padding: "24px 36px", borderBottom: `1px solid ${C.border}`, background: C.surface }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
          <div>
            <h1 style={{ fontFamily: "'Instrument Serif', serif", fontSize: 28, fontWeight: 400, letterSpacing: "-0.02em", color: C.text }}>
              Oral History Editor
            </h1>
            <p style={{ fontFamily: "'DM Mono', monospace", fontSize: 11, color: C.textDim, marginTop: 2, letterSpacing: "0.05em" }}>
              NARRATIVE AUDIO EDITOR WITH PROVENANCE TRACKING
            </p>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
            <UploadBtn onChange={handleAudioUpload} accept="audio/*" label={audioName ? `◆ ${audioName.slice(0, 18)}` : "↑ Audio"} />
            <UploadBtn onChange={handleSubUpload} accept=".srt,.vtt,.txt,.json" label={subName ? `◆ ${subName.slice(0, 18)}` : "↑ Subtitles"} />
            {audioUrl && !transcribing && (
              <SmBtn onClick={() => transcribeAudio(!!transcriptionProgress)} accent>
                {transcriptionProgress ? "⟳ Resume" : "⟳ Transcribe"}
              </SmBtn>
            )}
            {transcribing && (
              <SmBtn onClick={() => { transcribeCancelRef.current = true; }}>
                ■ Pause
              </SmBtn>
            )}
            <div style={{ width: 1, height: 20, background: C.border, margin: "0 2px" }} />
            <UploadBtn onChange={handleProjectUpload} accept=".json" label="↑ Load Project" />
            {hasContent && <SmBtn onClick={downloadProject} accent>↓ Save Project</SmBtn>}
            {hasContent && webhookBase && (
              <SmBtn onClick={cloudSave} accent>
                {syncStatus === "saving" ? "…" : "☁ Save"}
              </SmBtn>
            )}
            {hasContent && webhookBase && (
              <SmBtn onClick={() => setAutoSaveEnabled(!autoSaveEnabled)}
                accent={autoSaveEnabled}>
                {autoSaveEnabled ? "☁ Auto ✓" : "☁ Auto"}
              </SmBtn>
            )}
            <button onClick={() => setShowSettings(!showSettings)} style={{
              fontFamily: "'DM Mono', monospace", fontSize: 13, background: "none",
              border: `1px solid ${showSettings ? C.accent : C.border}`, borderRadius: 4,
              color: showSettings ? C.accent : C.textDim, cursor: "pointer", padding: "5px 8px", lineHeight: 1,
            }} title="Settings">⚙</button>
          </div>
        </div>
        {/* Sync status */}
        {syncStatus && (
          <div style={{ marginTop: 8, display: "inline-flex", alignItems: "center", gap: 8 }}>
            <div style={{
              fontFamily: "'DM Mono', monospace", fontSize: 11, padding: "4px 10px",
              borderRadius: 4, display: "inline-block",
              background: syncStatus === "error" ? "rgba(179,64,64,0.08)" : C.accentDim,
              color: syncStatus === "error" ? C.danger : C.accent,
              ...(transcribing ? { animation: "pulse 2s ease-in-out infinite" } : {}),
            }}>
              {syncMsg}
            </div>
            {transcribing && transcriptionProgress && transcriptionProgress.total > 1 && (
              <div style={{
                width: 120, height: 6, background: C.border, borderRadius: 3, overflow: "hidden",
              }}>
                <div style={{
                  width: `${(transcriptionProgress.completed / transcriptionProgress.total) * 100}%`,
                  height: "100%", background: C.accent, borderRadius: 3,
                  transition: "width 0.3s ease",
                }} />
              </div>
            )}
          </div>
        )}
      </header>

      {/* ── Settings Panel ── */}
      {showSettings && (
        <div style={{
          padding: "14px 36px", background: C.raised, borderBottom: `1px solid ${C.border}`,
          display: "flex", gap: 16, alignItems: "flex-end", flexWrap: "wrap",
        }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
            <label style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim }}>Project Title</label>
            <input value={projectTitle} onChange={(e) => setProjectTitle(e.target.value)}
              style={{ fontFamily: "'DM Mono', monospace", fontSize: 12, padding: "4px 8px", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, width: 200 }} />
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
            <label style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim }}>OpenAI API Key</label>
            <input value={openaiApiKey} onChange={(e) => setOpenaiApiKey(e.target.value)}
              type="password" placeholder="sk-..."
              style={{ fontFamily: "'DM Mono', monospace", fontSize: 12, padding: "4px 8px", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, width: 260 }} />
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
            <label style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim }}>Anthropic API Key</label>
            <input value={anthropicApiKey} onChange={(e) => setAnthropicApiKey(e.target.value)}
              type="password" placeholder="sk-ant-..."
              style={{ fontFamily: "'DM Mono', monospace", fontSize: 12, padding: "4px 8px", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, width: 260 }} />
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
            <label style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim }}>AI Provider</label>
            <div style={{ display: "flex", gap: 4 }}>
              <SmBtn onClick={() => setAiProvider("anthropic")} accent={aiProvider === "anthropic"}>Anthropic</SmBtn>
              <SmBtn onClick={() => setAiProvider("openai")} accent={aiProvider === "openai"}>OpenAI</SmBtn>
            </div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
            <label style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim }}>Google Drive File ID</label>
            <div style={{ display: "flex", gap: 6 }}>
              <input value={driveFileId} onChange={(e) => setDriveFileId(e.target.value)}
                placeholder="(auto-set on first save)"
                style={{ fontFamily: "'DM Mono', monospace", fontSize: 12, padding: "4px 8px", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, width: 240 }} />
              {driveFileId && webhookBase && (
                <SmBtn onClick={cloudLoad}>{syncStatus === "loading" ? "…" : "☁ Load"}</SmBtn>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ── Toolbar ── */}
      {hasContent && (
        <div style={{
          padding: "7px 36px", background: C.toolbar, borderBottom: `1px solid ${C.border}`,
          display: "flex", gap: 4, alignItems: "center", flexWrap: "wrap",
          fontFamily: "'DM Mono', monospace", fontSize: 11,
        }}>
          <TBtn onClick={doCut} disabled={!selected.size} label="Cut" k="⌘X" />
          <TBtn onClick={doCopy} disabled={!selected.size} label="Copy" k="⌘C" />
          <TBtn onClick={doPaste} disabled={!clipboard.length} label="Paste" k="⌘V" />
          <Sep />
          <TBtn onClick={doDelete} disabled={!selected.size} label="Delete" />
          <TBtn onClick={mergeSelected} disabled={selected.size < 2} label="Merge" />
          <TBtn onClick={addHeaderBlock} disabled={!hasContent} label="+ Header" />
          <Sep />
          <TBtn onClick={undo} disabled={!undoStack.length} label="Undo" k="⌘Z" />
          <Sep />
          <TBtn onClick={playAll} disabled={!audioUrl} label="▶ Play" accent />
          <TBtn onClick={stopPlayback} label="■ Stop" />
          <Sep />
          <TBtn onClick={() => setShowSpeakers(!showSpeakers)} label={showSpeakers ? "◂ Speakers" : "▸ Speakers"} />
          <TBtn onClick={() => setShowTOC(!showTOC)} label={showTOC ? "◂ Sections" : "▸ Sections"} />
          <div style={{ flex: 1 }} />
          <span style={{ color: C.textDim, fontSize: 10 }}>
            {blocks.length} ¶ · {fmtTime(totalDuration)}
          </span>
          <Sep />
          <TBtn onClick={() => setShowExport(!showExport)} label="Export" accent />
        </div>
      )}

      <div style={{ flex: 1, display: "flex", overflow: "hidden" }}>

        {/* ── Left Sidebar: Sections + Speakers ── */}
        {hasContent && (showSpeakers || showTOC) && (
          <div style={{
            width: 210, borderRight: `1px solid ${C.border}`, background: C.surface,
            display: "flex", flexDirection: "column", flexShrink: 0,
          }}>
            {/* ── Sections / TOC Panel ── */}
            {showTOC && (
              <div style={{ display: "flex", flexDirection: "column", borderBottom: showSpeakers ? `1px solid ${C.border}` : "none", maxHeight: showSpeakers ? "50%" : "100%", overflow: "hidden" }}>
                <div style={{ padding: "14px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <span style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim, letterSpacing: "0.06em", textTransform: "uppercase" }}>Sections</span>
                  <div style={{ display: "flex", gap: 2, alignItems: "center" }}>
                    <button onClick={analyzeTranscript} disabled={aiAnalyzing || blocks.length === 0}
                      style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, background: "none", border: "none", color: aiAnalyzing ? C.textDim : C.accent, cursor: aiAnalyzing ? "default" : "pointer", padding: "0 4px", lineHeight: 1 }}
                      title="AI auto-section">
                      {aiAnalyzing ? "…" : "AI"}
                    </button>
                    <button onClick={() => setAddingSection(true)} style={{ fontFamily: "'DM Mono', monospace", fontSize: 15, background: "none", border: "none", color: C.accent, cursor: "pointer", padding: "0 4px", lineHeight: 1 }}>+</button>
                  </div>
                </div>
                <div style={{ flex: 1, overflowY: "auto", padding: "6px 0" }}>
                  {/* Unsectioned count */}
                  {sectionedBlocks.unsectioned.length > 0 && (
                    <div
                      onClick={() => { if (selected.size > 0) unsectionBlocks(); }}
                      style={{
                        padding: "5px 16px", display: "flex", alignItems: "center", gap: 7,
                        cursor: selected.size > 0 ? "pointer" : "default",
                        borderLeft: "3px solid transparent",
                        background: "transparent", transition: "background 0.1s",
                      }}
                      onMouseEnter={(e) => { if (selected.size > 0) e.currentTarget.style.background = C.raised; }}
                      onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; }}
                    >
                      <span style={{ flex: 1, fontFamily: "'DM Mono', monospace", fontSize: 11, color: C.textDim, fontStyle: "italic" }}>Unsectioned</span>
                      <span style={{ fontFamily: "'DM Mono', monospace", fontSize: 9, color: C.textDim }}>{sectionedBlocks.unsectioned.length}</span>
                    </div>
                  )}

                  {sectionedBlocks.sections.map((section) => {
                    const isEd = editingSectionId === section.id;
                    return (
                      <div key={section.id}
                        onClick={() => {
                          if (selected.size > 0 && !isEd) {
                            assignToSection(section.id);
                          } else if (!isEd) {
                            document.getElementById('section-' + section.id)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
                          }
                        }}
                        style={{
                          padding: "5px 16px", display: "flex", alignItems: "center", gap: 7,
                          cursor: "pointer",
                          borderLeft: `3px solid ${C.accent}`,
                          background: activeSectionId === section.id ? C.accentDim : "transparent",
                          transition: "background 0.1s",
                        }}
                        onMouseEnter={(e) => { e.currentTarget.style.background = C.raised; }}
                        onMouseLeave={(e) => { e.currentTarget.style.background = activeSectionId === section.id ? C.accentDim : "transparent"; }}
                      >
                        <button onClick={(e) => { e.stopPropagation(); toggleSectionCollapse(section.id); }}
                          style={{ background: "none", border: "none", color: C.textDim, cursor: "pointer", fontSize: 10, padding: 0, lineHeight: 1 }}>
                          {section.collapsed ? "▸" : "▾"}
                        </button>
                        {isEd ? (
                          <input ref={editSectionRef} value={editingSectionName}
                            onChange={(e) => setEditingSectionName(e.target.value)}
                            onKeyDown={(e) => { if (e.key === "Enter") renameSection(section.id, editingSectionName); if (e.key === "Escape") setEditingSectionId(null); }}
                            onBlur={() => editingSectionName.trim() ? renameSection(section.id, editingSectionName) : setEditingSectionId(null)}
                            onClick={(e) => e.stopPropagation()}
                            style={{ flex: 1, background: C.raised, border: `1px solid ${C.accent}`, borderRadius: 3, color: C.text, fontFamily: "'DM Mono', monospace", fontSize: 11, padding: "2px 6px", minWidth: 0 }}
                          />
                        ) : (
                          <span onDoubleClick={(e) => { e.stopPropagation(); setEditingSectionId(section.id); setEditingSectionName(section.name); }}
                            style={{ flex: 1, fontFamily: "'DM Mono', monospace", fontSize: 11, color: C.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                            {section.name}
                          </span>
                        )}
                        <span style={{ fontFamily: "'DM Mono', monospace", fontSize: 9, color: C.textDim, minWidth: 14, textAlign: "right" }}>{section.blocks.length}</span>
                        {!isEd && (
                          <button onClick={(e) => { e.stopPropagation(); removeSection(section.id); }}
                            style={{ background: "none", border: "none", color: C.border, cursor: "pointer", fontSize: 12, padding: 0, lineHeight: 1, opacity: 0.4 }}
                            onMouseEnter={(e) => { e.currentTarget.style.opacity = "1"; e.currentTarget.style.color = C.danger; }}
                            onMouseLeave={(e) => { e.currentTarget.style.opacity = "0.4"; e.currentTarget.style.color = C.border; }}
                          >×</button>
                        )}
                      </div>
                    );
                  })}

                  {addingSection && (
                    <div style={{ padding: "5px 16px", display: "flex", alignItems: "center", gap: 7, borderLeft: `3px solid ${C.accent}` }}>
                      <input ref={newSectionRef} value={newSectionName} onChange={(e) => setNewSectionName(e.target.value)}
                        onKeyDown={(e) => { if (e.key === "Enter") addSection(newSectionName); if (e.key === "Escape") { setAddingSection(false); setNewSectionName(""); } }}
                        onBlur={() => newSectionName.trim() ? addSection(newSectionName) : (setAddingSection(false), setNewSectionName(""))}
                        placeholder="Section name..." style={{ flex: 1, background: C.raised, border: `1px solid ${C.accent}`, borderRadius: 3, color: C.text, fontFamily: "'DM Mono', monospace", fontSize: 11, padding: "2px 6px", minWidth: 0 }}
                      />
                    </div>
                  )}

                  {sections.length === 0 && !addingSection && (
                    <div style={{ padding: "16px 16px", textAlign: "center" }}>
                      <p style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim, lineHeight: 1.6 }}>Organize paragraphs<br />into sections</p>
                      <button onClick={() => setAddingSection(true)}
                        style={{ marginTop: 8, fontFamily: "'DM Mono', monospace", fontSize: 10, padding: "4px 10px", background: C.accentDim, border: `1px solid ${C.accent}`, borderRadius: 4, color: C.accent, cursor: "pointer" }}>
                        + Add section
                      </button>
                    </div>
                  )}
                </div>

                {selected.size > 0 && sections.length > 0 && (
                  <div style={{ padding: "6px 16px", borderTop: `1px solid ${C.border}`, background: C.raised }}>
                    <p style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.accent, lineHeight: 1.4 }}>
                      {selected.size} selected — click section to assign
                    </p>
                  </div>
                )}
              </div>
            )}

            {/* ── Speaker Panel ── */}
            {showSpeakers && (
              <div style={{ display: "flex", flexDirection: "column", flex: 1, overflow: "hidden" }}>
                <div style={{ padding: "14px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <span style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim, letterSpacing: "0.06em", textTransform: "uppercase" }}>Speakers</span>
                  <button onClick={() => setAddingSpeaker(true)} style={{ fontFamily: "'DM Mono', monospace", fontSize: 15, background: "none", border: "none", color: C.accent, cursor: "pointer", padding: "0 4px", lineHeight: 1 }}>+</button>
                </div>

                <div style={{ flex: 1, overflowY: "auto", padding: "6px 0" }}>
                  {speakers.map((spk, idx) => {
                    const count = blocks.filter((_, i) => getEffectiveSpeakerId(i) === spk.id).length;
                    const isEd = editingSpeakerId === spk.id;
                    return (
                      <div key={spk.id}
                        onClick={() => { if (selected.size > 0 && !isEd) assignAndMerge(spk.id); else if (selected.size === 0 && !isEd) { setEditingSpeakerId(spk.id); setEditingSpeakerName(spk.name); } }}
                        style={{
                          padding: "5px 16px", display: "flex", alignItems: "center", gap: 7,
                          cursor: selected.size > 0 ? "pointer" : "default",
                          borderLeft: `3px solid ${spk.color}`,
                          background: "transparent", transition: "background 0.1s",
                        }}
                        onMouseEnter={(e) => { if (selected.size > 0) e.currentTarget.style.background = C.raised; }}
                        onMouseLeave={(e) => { e.currentTarget.style.background = "transparent"; }}
                      >
                        {isEd ? (
                          <input ref={editSpkRef} value={editingSpeakerName}
                            onChange={(e) => setEditingSpeakerName(e.target.value)}
                            onKeyDown={(e) => { if (e.key === "Enter") renameSpeaker(spk.id, editingSpeakerName); if (e.key === "Escape") setEditingSpeakerId(null); }}
                            onBlur={() => editingSpeakerName.trim() ? renameSpeaker(spk.id, editingSpeakerName) : setEditingSpeakerId(null)}
                            style={{ flex: 1, background: C.raised, border: `1px solid ${spk.color}`, borderRadius: 3, color: C.text, fontFamily: "'DM Mono', monospace", fontSize: 12, padding: "2px 6px", minWidth: 0 }}
                          />
                        ) : (
                          <span onDoubleClick={(e) => { e.stopPropagation(); setEditingSpeakerId(spk.id); setEditingSpeakerName(spk.name); }}
                            style={{ flex: 1, fontFamily: "'DM Mono', monospace", fontSize: 12, color: C.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                            {spk.name}
                          </span>
                        )}
                        <span style={{ fontFamily: "'DM Mono', monospace", fontSize: 9, color: C.border, minWidth: 12, textAlign: "center" }}>{idx < 9 ? idx + 1 : ""}</span>
                        <span style={{ fontFamily: "'DM Mono', monospace", fontSize: 9, color: C.textDim, minWidth: 14, textAlign: "right" }}>{count}</span>
                        {!isEd && (
                          <button onClick={(e) => { e.stopPropagation(); removeSpeaker(spk.id); }}
                            style={{ background: "none", border: "none", color: C.border, cursor: "pointer", fontSize: 12, padding: 0, lineHeight: 1, opacity: 0.4 }}
                            onMouseEnter={(e) => { e.currentTarget.style.opacity = "1"; e.currentTarget.style.color = C.danger; }}
                            onMouseLeave={(e) => { e.currentTarget.style.opacity = "0.4"; e.currentTarget.style.color = C.border; }}
                          >×</button>
                        )}
                      </div>
                    );
                  })}

                  {addingSpeaker && (
                    <div style={{ padding: "5px 16px", display: "flex", alignItems: "center", gap: 7, borderLeft: `3px solid ${SPEAKER_PALETTE[speakers.length % SPEAKER_PALETTE.length].color}` }}>
                      <input ref={newSpkRef} value={newSpeakerName} onChange={(e) => setNewSpeakerName(e.target.value)}
                        onKeyDown={(e) => { if (e.key === "Enter") addSpeaker(newSpeakerName); if (e.key === "Escape") { setAddingSpeaker(false); setNewSpeakerName(""); } }}
                        onBlur={() => newSpeakerName.trim() ? addSpeaker(newSpeakerName) : (setAddingSpeaker(false), setNewSpeakerName(""))}
                        placeholder="Name..." style={{ flex: 1, background: C.raised, border: `1px solid ${SPEAKER_PALETTE[speakers.length % SPEAKER_PALETTE.length].color}`, borderRadius: 3, color: C.text, fontFamily: "'DM Mono', monospace", fontSize: 12, padding: "2px 6px", minWidth: 0 }}
                      />
                    </div>
                  )}

                  {speakers.length === 0 && !addingSpeaker && (
                    <div style={{ padding: "24px 16px", textAlign: "center" }}>
                      <p style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim, lineHeight: 1.6 }}>Add speakers to tag<br />each paragraph</p>
                      <button onClick={() => setAddingSpeaker(true)}
                        style={{ marginTop: 8, fontFamily: "'DM Mono', monospace", fontSize: 10, padding: "4px 10px", background: C.accentDim, border: `1px solid ${C.accent}`, borderRadius: 4, color: C.accent, cursor: "pointer" }}>
                        + Add speaker
                      </button>
                    </div>
                  )}
                </div>

                {selected.size > 0 && speakers.length > 0 && (
                  <div style={{ padding: "8px 16px", borderTop: `1px solid ${C.border}`, background: C.raised }}>
                    <p style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.accent, lineHeight: 1.4 }}>
                      {selected.size} selected — press 1–{Math.min(speakers.length, 9)} or click
                    </p>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* ── Main Content: Paragraph View ── */}
        <div ref={editorContentRef} style={{ flex: 1, overflowY: "auto", padding: "32px 48px 100px", maxWidth: 800, margin: "0 auto", width: "100%" }}>
          {!hasContent && (
            <div style={{ textAlign: "center", padding: "80px 20px" }}>
              <p style={{ fontFamily: "'Instrument Serif', serif", fontSize: 32, color: C.text, marginBottom: 12 }}>Begin an oral history</p>
              <p style={{ fontSize: 16, lineHeight: 1.7, color: C.textDim, maxWidth: 420, margin: "0 auto" }}>
                Upload an audio recording and its subtitle file. The transcript becomes an editable narrative — reorder, attribute speakers, fix text — with full audio provenance.
              </p>
              <div style={{ marginTop: 36, display: "flex", gap: 16, justifyContent: "center", flexWrap: "wrap" }}>
                {[["♫", "Audio", "WAV, MP3, M4A"], ["¶", "Subtitles", "SRT or VTT"], ["✎", "Edit", "Text & structure"], ["↓", "Export", "MD + project JSON"]].map(([icon, t, d]) => (
                  <div key={t} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "18px 22px", width: 120, textAlign: "center" }}>
                    <div style={{ fontSize: 22, marginBottom: 6 }}>{icon}</div>
                    <div style={{ fontFamily: "'DM Mono', monospace", fontSize: 11, color: C.text }}>{t}</div>
                    <div style={{ fontFamily: "'DM Mono', monospace", fontSize: 9, color: C.textDim, marginTop: 2 }}>{d}</div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ── Section-grouped block rendering ── */}
          {hasContent && (() => {
            const renderBlock = (block, idx) => {
              const isSelected = selected.has(block.id);
              const isEditing = editingId === block.id;
              const isPlaying = playing === block.id;
              const isDragTarget = dragOver === block.id && dragging !== block.id;
              const spk = getEffectiveSpeaker(idx);
              const isInherited = !block.speaker && spk;
              const text = blockText(block);

              // Header block rendering
              if (block.type === "header") {
                const hSize = { 1: 28, 2: 22, 3: 18 }[block.level || 2] || 22;
                return (
                  <div
                    key={block.id}
                    onClick={(e) => handleSelect(block.id, e)}
                    draggable={!isEditing}
                    onDragStart={() => setDragging(block.id)}
                    onDragOver={(e) => { e.preventDefault(); setDragOver(block.id); }}
                    onDrop={() => handleDrop(block.id)}
                    onDragEnd={() => { setDragging(null); setDragOver(null); }}
                    style={{
                      marginBottom: 8, padding: "8px 20px", borderRadius: 6,
                      background: isSelected ? C.selected : "transparent",
                      border: `1px solid ${isDragTarget ? C.accent : isSelected ? C.selectedBorder : "transparent"}`,
                      cursor: "pointer", transition: "all 0.12s ease",
                    }}
                  >
                    {isEditing ? (
                      <div>
                        <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                          {[1, 2, 3].map(lvl => (
                            <SmBtn key={lvl} onClick={(e) => { e.stopPropagation(); setBlocks(prev => prev.map(b => b.id === block.id ? { ...b, level: lvl } : b)); }}
                              accent={block.level === lvl}>H{lvl}</SmBtn>
                          ))}
                        </div>
                        <input autoFocus value={editText} onChange={(e) => setEditText(e.target.value)}
                          onKeyDown={(e) => { if (e.key === "Enter") commitEdit(); if (e.key === "Escape") cancelEdit(); }}
                          style={{
                            width: "100%", background: C.surface, border: `1px solid ${C.accent}`, borderRadius: 4,
                            color: C.text, fontFamily: "'Instrument Serif', serif", fontSize: hSize, fontWeight: 400,
                            padding: "8px 12px",
                          }}
                        />
                        <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                          <SmBtn onClick={commitEdit} accent>Save</SmBtn>
                          <SmBtn onClick={cancelEdit}>Cancel</SmBtn>
                        </div>
                      </div>
                    ) : (
                      <h2 style={{ fontFamily: "'Instrument Serif', serif", fontSize: hSize, fontWeight: 400, color: C.text, margin: 0 }}>
                        {block.text || "Untitled"}
                      </h2>
                    )}
                    {!isEditing && isSelected && (
                      <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
                        <SmBtn onClick={(e) => { e.stopPropagation(); startEdit(block); }}>✎ Edit</SmBtn>
                        {[1, 2, 3].map(lvl => (
                          <SmBtn key={lvl} onClick={(e) => { e.stopPropagation(); setBlocks(prev => prev.map(b => b.id === block.id ? { ...b, level: lvl } : b)); }}
                            accent={block.level === lvl}>H{lvl}</SmBtn>
                        ))}
                      </div>
                    )}
                  </div>
                );
              }

              return (
                <div
                  key={block.id}
                  onClick={(e) => handleSelect(block.id, e)}
                  draggable={!isEditing}
                  onDragStart={() => setDragging(block.id)}
                  onDragOver={(e) => { e.preventDefault(); setDragOver(block.id); }}
                  onDrop={() => handleDrop(block.id)}
                  onDragEnd={() => { setDragging(null); setDragOver(null); }}
                  style={{
                    marginBottom: spk ? 16 : 4,
                    padding: "10px 20px",
                    borderRadius: 6,
                    background: isSelected ? C.selected : "transparent",
                    border: `1px solid ${isDragTarget ? C.accent : isSelected ? C.selectedBorder : "transparent"}`,
                    borderLeft: `3px solid ${spk ? spk.color : "transparent"}`,
                    cursor: "pointer",
                    transition: "all 0.12s ease",
                    animation: "fadeIn 0.2s ease",
                    position: "relative",
                  }}
                >
                  {/* Speaker name + timestamp */}
                  <div style={{ display: "flex", alignItems: "baseline", gap: 12, marginBottom: 6 }}>
                    {spk ? (
                      <span style={{
                        fontFamily: "'DM Mono', monospace", fontSize: 12, fontWeight: 500,
                        color: spk.color, letterSpacing: "0.02em",
                        opacity: isInherited ? 0.45 : 1,
                      }}>
                        {spk.name}{isInherited ? " (cont.)" : ""}
                      </span>
                    ) : null}
                    <span style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim }}>
                      {fmtTime(blockStart(block))} — {fmtTime(blockEnd(block))}
                    </span>
                    {isPlaying && <span style={{ color: C.accent, animation: "pulse 1s infinite", fontSize: 12 }}>▶</span>}
                    {(block.wasEdited || block.wasMoved) && (
                      <span style={{ fontFamily: "'DM Mono', monospace", fontSize: 9, color: C.textDim, opacity: 0.6 }}>
                        {[block.wasEdited && "edited", block.wasMoved && "moved"].filter(Boolean).join(" · ")}
                      </span>
                    )}
                  </div>

                  {/* Paragraph text */}
                  {isEditing ? (
                    <div>
                      <textarea
                        autoFocus value={editText} onChange={(e) => { setEditText(e.target.value); e.target.style.height = "auto"; e.target.style.height = e.target.scrollHeight + "px"; }}
                        onKeyDown={(e) => { if (e.key === "Escape") cancelEdit(); }}
                        ref={(el) => { if (el) { el.style.height = "auto"; el.style.height = el.scrollHeight + "px"; } }}
                        style={{
                          width: "100%", background: C.surface, border: `1px solid ${C.accent}`, borderRadius: 4,
                          color: C.text, fontFamily: "'Newsreader', Georgia, serif", fontSize: 17, lineHeight: 1.75,
                          padding: "12px 14px", resize: "vertical", minHeight: Math.min(Math.max(200, text.length / 2), 500),
                          maxHeight: "60vh", overflow: "auto",
                        }}
                      />
                      <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
                        <SmBtn onClick={commitEdit} accent>Save</SmBtn>
                        <SmBtn onClick={() => {
                          const ta = document.querySelector("textarea:focus");
                          if (ta && ta.selectionStart > 0 && ta.selectionStart < editText.length) {
                            cancelEdit();
                            splitBlock(block.id, ta.selectionStart);
                          }
                        }}>Split at cursor</SmBtn>
                        <SmBtn onClick={cancelEdit}>Cancel</SmBtn>
                      </div>
                    </div>
                  ) : (
                    <p style={{ fontSize: 17, lineHeight: 1.75, margin: 0, color: C.text }}>
                      {block.segments.map((seg, si) => (
                        <span key={seg.id || si}
                          onClick={(e) => {
                            if (audioUrl && audioRef.current) {
                              e.stopPropagation();
                              audioRef.current.currentTime = seg.start;
                              audioRef.current.play();
                              setPlaying(block.id);
                              setPlayingSegId(seg.id);
                            }
                          }}
                          style={{
                            cursor: audioUrl ? "pointer" : "default",
                            background: playingSegId === seg.id ? C.accentDim : "transparent",
                            borderRadius: 2, padding: "0 1px",
                            transition: "background 0.15s",
                          }}
                          title={audioUrl ? `${fmtTime(seg.start)} — ${fmtTime(seg.end)}` : undefined}
                        >
                          {seg.text}{si < block.segments.length - 1 ? " " : ""}
                        </span>
                      ))}
                    </p>
                  )}

                  {/* Hover actions */}
                  {!isEditing && isSelected && (
                    <div style={{ display: "flex", gap: 6, marginTop: 8, flexWrap: "wrap" }}>
                      {audioUrl && <SmBtn onClick={(e) => { e.stopPropagation(); playBlock(block); }}>▶ Play</SmBtn>}
                      <SmBtn onClick={(e) => { e.stopPropagation(); startEdit(block); }}>✎ Edit</SmBtn>
                      <SmBtn onClick={(e) => { e.stopPropagation(); pushUndo(); splitBlock(block.id, Math.floor(text.length / 2)); }}>Split ½</SmBtn>
                      <SmBtn onClick={(e) => { e.stopPropagation(); splitBySentence(block.id); }}>Split sentences</SmBtn>
                      {block.segments.length > 1 && <SmBtn onClick={(e) => { e.stopPropagation(); splitAtPause(block.id); }}>Split at pause</SmBtn>}
                      {block.segments.length > 1 && <SmBtn onClick={(e) => { e.stopPropagation(); decomposeBlock(block.id); }}>Decompose</SmBtn>}
                      {block.wasEdited && <SmBtn onClick={(e) => {
                        e.stopPropagation(); pushUndo();
                        setBlocks(prev => prev.map(b => b.id !== block.id ? b : {
                          ...b, wasEdited: false, segments: b.segments.map(s => ({ ...s, text: s.originalText })),
                        }));
                      }}>Revert</SmBtn>}
                    </div>
                  )}
                </div>
              );
            };

            // If no sections exist, render all blocks flat (original behavior)
            if (sections.length === 0) {
              return blocks.map((block, idx) => renderBlock(block, idx));
            }

            // Section-grouped rendering
            return (
              <>
                {/* Unsectioned blocks */}
                {sectionedBlocks.unsectioned.length > 0 && (
                  <div style={{ marginBottom: 24 }}>
                    <div style={{
                      fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim,
                      letterSpacing: "0.06em", textTransform: "uppercase", padding: "8px 0",
                      borderBottom: `1px solid ${C.border}`, marginBottom: 12,
                    }}>
                      Unsectioned
                    </div>
                    {sectionedBlocks.unsectioned.map(({ block, globalIdx }) => renderBlock(block, globalIdx))}
                  </div>
                )}

                {/* Sectioned blocks */}
                {sectionedBlocks.sections.map((section) => (
                  <div key={section.id} id={`section-${section.id}`} style={{ marginBottom: 28 }}>
                    <div
                      onClick={() => toggleSectionCollapse(section.id)}
                      style={{
                        display: "flex", alignItems: "center", gap: 10, padding: "10px 0",
                        borderBottom: `2px solid ${C.accent}`, marginBottom: 14, cursor: "pointer",
                      }}
                    >
                      <span style={{ fontSize: 12, color: C.textDim }}>{section.collapsed ? "▸" : "▾"}</span>
                      <h2 style={{
                        fontFamily: "'Instrument Serif', serif", fontSize: 22, fontWeight: 400,
                        color: C.text, margin: 0, flex: 1,
                      }}>
                        {section.name}
                      </h2>
                      <span style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim }}>
                        {section.blocks.length} ¶ · {fmtTime(section.duration)}
                      </span>
                    </div>
                    {!section.collapsed && section.blocks.map(({ block, globalIdx }) => renderBlock(block, globalIdx))}
                    {section.collapsed && section.blocks.length > 0 && (
                      <div style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim, padding: "8px 20px", fontStyle: "italic" }}>
                        {section.blocks.length} paragraphs hidden
                      </div>
                    )}
                  </div>
                ))}
              </>
            );
          })()}
        </div>

        {/* ── Scrub Bar ── */}
        {audioUrl && hasContent && (
          <div style={{
            position: "fixed", bottom: 40, left: "50%", transform: "translateX(-50%)",
            background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8,
            padding: "8px 16px", display: "flex", alignItems: "center", gap: 12,
            boxShadow: "0 2px 12px rgba(0,0,0,0.1)", zIndex: 50, minWidth: 400,
            fontFamily: "'DM Mono', monospace", fontSize: 10,
          }}>
            <span style={{ color: C.textDim, minWidth: 50 }}>{fmtTime(currentTime)}</span>
            <div
              onClick={(e) => {
                const rect = e.currentTarget.getBoundingClientRect();
                const ratio = (e.clientX - rect.left) / rect.width;
                const newTime = ratio * audioDuration;
                if (audioRef.current) { audioRef.current.currentTime = newTime; setCurrentTime(newTime); }
              }}
              style={{
                flex: 1, height: 6, background: C.raised, borderRadius: 3, cursor: "pointer", position: "relative",
              }}
            >
              <div style={{
                width: `${audioDuration > 0 ? (currentTime / audioDuration) * 100 : 0}%`,
                height: "100%", background: C.accent, borderRadius: 3, transition: "width 0.1s",
              }} />
            </div>
            <span style={{ color: C.textDim, minWidth: 50, textAlign: "right" }}>{fmtTime(audioDuration)}</span>
            <button onClick={() => { if (audioRef.current) { if (audioRef.current.paused) audioRef.current.play(); else audioRef.current.pause(); } }}
              style={{ background: "none", border: "none", color: C.accent, cursor: "pointer", fontSize: 14, padding: "0 4px" }}>
              {playing ? "■" : "▶"}
            </button>
          </div>
        )}

        {/* ── Export Panel ── */}
        {showExport && hasContent && (
          <div style={{ width: 340, borderLeft: `1px solid ${C.border}`, background: C.surface, overflowY: "auto", display: "flex", flexDirection: "column", flexShrink: 0 }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}` }}>
              <h3 style={{ fontFamily: "'Instrument Serif', serif", fontSize: 18, fontWeight: 400 }}>Export</h3>
              <p style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim, marginTop: 4 }}>
                {blocks.length} paragraphs · {fmtTime(totalDuration)} · {speakers.length} speakers
              </p>
              <div style={{ display: "flex", gap: 8, marginTop: 12, flexWrap: "wrap" }}>
                <SmBtn onClick={() => downloadMarkdown()} accent>↓ Markdown (.md)</SmBtn>
                <SmBtn onClick={downloadProject} accent>↓ Project (.json)</SmBtn>
              </div>
              {sections.length > 0 && (
                <div style={{ marginTop: 12 }}>
                  <div style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim, marginBottom: 6, letterSpacing: "0.04em" }}>
                    PER-SECTION EXPORT
                  </div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {sections.map(section => {
                      const count = blocks.filter(b => b.sectionId === section.id).length;
                      return (
                        <SmBtn key={section.id} onClick={() => downloadMarkdown(section.id)} disabled={count === 0}>
                          ↓ {section.name} ({count})
                        </SmBtn>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>

            {/* MD Preview */}
            <div style={{ padding: "14px 20px", flex: 1 }}>
              <div style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim, marginBottom: 8, letterSpacing: "0.04em" }}>
                MARKDOWN PREVIEW
              </div>
              <div style={{
                fontFamily: "'Newsreader', Georgia, serif", fontSize: 13, lineHeight: 1.7,
                color: C.textDim, background: C.raised, padding: 14, borderRadius: 6,
                overflow: "auto", maxHeight: 400, border: `1px solid ${C.border}`,
              }}>
              {(() => {
                let lastSpkName = null;
                return blocks.slice(0, 6).map((block, i) => {
                  const spk = getEffectiveSpeaker(i);
                  const name = spk?.name || "Unknown Speaker";
                  const showName = name !== lastSpkName;
                  lastSpkName = name;
                  return (
                    <div key={i} style={{ marginBottom: 14 }}>
                      {showName && (
                        <strong style={{ color: spk?.color || C.textDim }}>
                          {name}:
                        </strong>
                      )}{showName ? " " : ""}
                      <span>{blockText(block).slice(0, 200)}{blockText(block).length > 200 ? "…" : ""}</span>
                    </div>
                  );
                });
              })()}
                {blocks.length > 6 && (
                  <p style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.border }}>
                    … and {blocks.length - 6} more paragraphs
                  </p>
                )}
              </div>

              <div style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim, marginBottom: 8, marginTop: 20, letterSpacing: "0.04em" }}>
                PROJECT JSON PREVIEW
              </div>
              <pre style={{
                fontFamily: "'DM Mono', monospace", fontSize: 10, lineHeight: 1.5,
                color: C.textDim, background: C.raised, padding: 12, borderRadius: 6,
                overflow: "auto", maxHeight: 250, whiteSpace: "pre-wrap", wordBreak: "break-all",
                border: `1px solid ${C.border}`,
              }}>
                {JSON.stringify({ ...generateProject(), blocks: generateProject().blocks.slice(0, 3), ffmpeg: "..." }, null, 2)}
                {blocks.length > 3 && `\n  ... +${blocks.length - 3} more blocks`}
              </pre>
            </div>
          </div>
        )}
      </div>

      {/* ── AI Suggestions Modal ── */}
      {aiSuggestions?.sections && (
        <div style={{
          position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100,
        }} onClick={() => setAiSuggestions(null)}>
          <div onClick={(e) => e.stopPropagation()} style={{
            background: C.surface, borderRadius: 12, padding: "24px 28px", maxWidth: 500, width: "90%",
            maxHeight: "80vh", overflowY: "auto", border: `1px solid ${C.border}`,
          }}>
            <h3 style={{ fontFamily: "'Instrument Serif', serif", fontSize: 22, fontWeight: 400, marginBottom: 16 }}>AI Suggested Sections</h3>
            {aiSuggestions.sections.map((sec, i) => (
              <label key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: `1px solid ${C.border}`, cursor: "pointer" }}>
                <input type="checkbox" checked={aiSelectedSections.has(i)}
                  onChange={() => setAiSelectedSections(prev => { const n = new Set(prev); n.has(i) ? n.delete(i) : n.add(i); return n; })} />
                <div style={{ flex: 1 }}>
                  <div style={{ fontFamily: "'DM Mono', monospace", fontSize: 12, color: C.text }}>{sec.name}</div>
                  <div style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim }}>
                    Blocks {sec.startBlock}–{sec.endBlock} ({sec.endBlock - sec.startBlock + 1} paragraphs)
                  </div>
                </div>
              </label>
            ))}
            <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
              <SmBtn onClick={() => setAiSelectedSections(new Set(aiSuggestions.sections.map((_, i) => i)))}>Select All</SmBtn>
              <SmBtn onClick={() => setAiSelectedSections(new Set())}>Deselect All</SmBtn>
              <div style={{ flex: 1 }} />
              <SmBtn onClick={() => setAiSuggestions(null)}>Cancel</SmBtn>
              <SmBtn onClick={() => { applyAiSuggestions([...aiSelectedSections]); setAiSelectedSections(new Set()); }} accent disabled={aiSelectedSections.size === 0}>Apply Selected</SmBtn>
            </div>
          </div>
        </div>
      )}

      {/* ── Footer ── */}
      {hasContent && (
        <div style={{
          padding: "7px 36px", borderTop: `1px solid ${C.border}`, background: C.surface,
          fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.border, display: "flex", gap: 14, flexWrap: "wrap",
        }}>
          <span>Click select</span><span>Shift range</span><span>⌘ multi</span>
          <span>✎ edit text</span><span>Drag reorder</span>
          <span>1–9 assign speaker</span><span>0 clear</span><span>⌘Z undo</span>
        </div>
      )}
    </div>
  );
}

/* ── Sub-components ── */
function UploadBtn({ onChange, accept, label }) {
  return (
    <label style={{
      fontFamily: "'DM Mono', monospace", fontSize: 11, padding: "7px 14px",
      background: C.raised, border: `1px solid ${C.border}`, borderRadius: 5,
      color: C.text, letterSpacing: "0.02em", cursor: "pointer",
    }}>
      {label}
      <input type="file" accept={accept} onChange={onChange} style={{ display: "none" }} />
    </label>
  );
}

function TBtn({ onClick, disabled, label, k, accent }) {
  return (
    <button onClick={onClick} disabled={disabled} style={{
      fontFamily: "'DM Mono', monospace", fontSize: 11, padding: "3px 9px",
      background: "transparent", border: "none", borderRadius: 4,
      color: disabled ? C.border : accent ? C.accent : C.textDim,
      cursor: disabled ? "default" : "pointer", display: "flex", gap: 3, alignItems: "center",
      opacity: disabled ? 0.35 : 1, transition: "color 0.15s",
    }}>
      {label}{k && <span style={{ fontSize: 9, opacity: 0.5 }}>{k}</span>}
    </button>
  );
}

function SmBtn({ onClick, children, accent, disabled }) {
  return (
    <button onClick={onClick} disabled={disabled} style={{
      fontFamily: "'DM Mono', monospace", fontSize: 10, padding: "3px 10px",
      background: accent ? C.accentDim : C.raised,
      border: `1px solid ${accent ? C.accent : C.border}`,
      borderRadius: 4, color: accent ? C.accent : C.textDim,
      cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.5 : 1,
    }}>
      {children}
    </button>
  );
}

function Sep() {
  return <div style={{ width: 1, height: 14, background: C.border, margin: "0 3px" }} />;
}
