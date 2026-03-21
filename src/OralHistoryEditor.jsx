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
  const [webhookBase, setWebhookBase] = useState("https://n8n.intelechia.com/webhook/oral-history");
  const [driveFileId, setDriveFileId] = useState("");
  const [syncStatus, setSyncStatus] = useState(null); // null | "saving" | "loading" | "saved" | "loaded" | "error"
  const [syncMsg, _setSyncMsg] = useState("");
  const syncMsgRef = useRef("");
  const setSyncMsg = useCallback((v) => { syncMsgRef.current = v; _setSyncMsg(v); }, []);
  const [projectTitle, setProjectTitle] = useState("Untitled Oral History");
  const [openaiApiKey, setOpenaiApiKey] = useState("");
  const [transcribing, setTranscribing] = useState(false);

  const audioRef = useRef(null);
  const audioFileRef = useRef(null);
  const playTimerRef = useRef(null);
  const newSpkRef = useRef(null);
  const editSpkRef = useRef(null);
  const projectUploadRef = useRef(null);

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

  const transcribeAudio = async () => {
    if (!audioFileRef.current) {
      setSyncStatus("error"); setSyncMsg("Upload an audio file first");
      setTimeout(() => setSyncStatus(null), 3000); return;
    }
    if (!openaiApiKey.trim()) {
      setShowSettings(true);
      setSyncStatus("error"); setSyncMsg("Enter your OpenAI API key in Settings");
      setTimeout(() => setSyncStatus(null), 3000); return;
    }
    setTranscribing(true);
    const startTime = Date.now();
    const elapsed = () => {
      const s = Math.floor((Date.now() - startTime) / 1000);
      return s < 60 ? `${s}s` : `${Math.floor(s / 60)}m ${s % 60}s`;
    };
    setSyncStatus("saving"); setSyncMsg("Preparing audio for upload…");
    // Show elapsed time while transcribing
    const timer = setInterval(() => {
      if (syncMsgRef.current) {
        const base = syncMsgRef.current.replace(/ \(\d+[ms].*?\)$/, "");
        setSyncMsg(`${base} (${elapsed()})`);
      }
    }, 1000);
    try {
      const formData = new FormData();
      formData.append("file", audioFileRef.current);
      formData.append("model", "whisper-1");
      formData.append("response_format", "verbose_json");
      formData.append("timestamp_granularities[]", "segment");
      const fileSizeMB = (audioFileRef.current.size / (1024 * 1024)).toFixed(1);
      setSyncMsg(`Uploading ${fileSizeMB} MB to OpenAI…`);
      const res = await fetch("https://api.openai.com/v1/audio/transcriptions", {
        method: "POST",
        headers: { Authorization: `Bearer ${openaiApiKey.trim()}` },
        body: formData,
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error?.message || `API error ${res.status}`);
      }
      setSyncMsg(`Transcription complete, processing results… (${elapsed()})`);
      const data = await res.json();
      const segs = (data.segments || []).map((seg, i) => ({
        id: uid(), text: seg.text.trim(), originalText: seg.text.trim(),
        start: seg.start, end: seg.end, originalIndex: i,
      }));
      if (segs.length === 0) throw new Error("No segments returned");
      const newBlocks = segs.map((seg) => ({
        id: uid(), speaker: null, segments: [seg], wasMoved: false, wasEdited: false,
      }));
      setBlocks(newBlocks);
      setSelected(new Set());
      setUndoStack([]);
      setSubName("(transcribed)");
      clearInterval(timer);
      setSyncStatus("loaded"); setSyncMsg(`Transcribed ${segs.length} segments in ${elapsed()}`);
      setTimeout(() => setSyncStatus(null), 4000);
    } catch (err) {
      clearInterval(timer);
      setSyncStatus("error"); setSyncMsg(err.message);
      setTimeout(() => setSyncStatus(null), 5000);
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
      const segments = detectAndParse(ev.target.result);
      // Each segment becomes its own block initially (unassigned speaker)
      const newBlocks = segments.map((seg) => ({
        id: uid(),
        speaker: null,
        segments: [seg],
        wasMoved: false,
        wasEdited: false,
      }));
      setBlocks(newBlocks);
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
    const spkNameToId = {};
    (proj.speakers || []).forEach((s) => { spkNameToId[s.name] = s.id || uid(); });
    const restored = proj.blocks.map((b) => ({
      id: uid(),
      speaker: b.speakerId || (b.speaker ? spkNameToId[b.speaker] : null) || null,
      wasMoved: b.wasMoved || false,
      wasEdited: b.wasEdited || false,
      segments: (b.segments || []).map((seg) => ({
        id: uid(),
        text: seg.text,
        originalText: seg.originalText || seg.text,
        start: seg.sourceStart,
        end: seg.sourceEnd,
        originalIndex: seg.originalIndex || 0,
      })),
    }));
    setBlocks(restored);
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
    setUndoStack((prev) => [...prev.slice(-30), JSON.parse(JSON.stringify(blocks))]);
  }, [blocks]);

  const undo = useCallback(() => {
    if (undoStack.length === 0) return;
    setBlocks(undoStack[undoStack.length - 1]);
    setUndoStack((s) => s.slice(0, -1));
    setSelected(new Set());
    setEditingId(null);
  }, [undoStack]);

  /* ── Derived data ── */
  const blockText = useCallback((block) => block.segments.map((s) => s.text).join(" "), []);
  const blockStart = useCallback((block) => Math.min(...block.segments.map((s) => s.start)), []);
  const blockEnd = useCallback((block) => Math.max(...block.segments.map((s) => s.end)), []);
  const blockDuration = useCallback((block) => block.segments.reduce((sum, s) => sum + (s.end - s.start), 0), []);

  const totalDuration = useMemo(() => blocks.reduce((sum, b) => sum + blockDuration(b), 0), [blocks, blockDuration]);

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
  const mergeAdjacent = useCallback((blockList) => {
    if (blockList.length <= 1) return blockList;
    // Compute effective speakers for the list
    const eff = blockList.map((b, i) => {
      if (b.speaker) return b.speaker;
      for (let j = i - 1; j >= 0; j--) { if (blockList[j].speaker) return blockList[j].speaker; }
      return null;
    });
    const result = [{ ...blockList[0], segments: [...blockList[0].segments] }];
    let resultEff = [eff[0]];
    for (let i = 1; i < blockList.length; i++) {
      const prevEff = resultEff[resultEff.length - 1];
      const currEff = eff[i];
      if (prevEff && prevEff === currEff) {
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
  }, []);

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

      const b1 = { id: uid(), speaker: block.speaker, segments: segs1, wasMoved: block.wasMoved, wasEdited: true };
      const b2 = { id: uid(), speaker: block.speaker, segments: segs2, wasMoved: block.wasMoved, wasEdited: true };
      const next = [...prev];
      next.splice(idx, 1, b1, b2);
      return next;
    });
  }, [pushUndo]);

  /* ── Inline text edit ── */
  const startEdit = (block) => { setEditingId(block.id); setEditText(blockText(block)); };

  const commitEdit = useCallback(() => {
    if (!editingId) return;
    pushUndo();
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

  /* ── Full project state for save ── */
  const buildProjectState = useCallback(() => {
    const speakerMap = {};
    speakers.forEach((s) => { speakerMap[s.id] = s.name; });
    return {
      version: 1,
      title: projectTitle,
      sourceFile: audioName || "source.wav",
      totalDuration: +(blocks.reduce((sum, b) => sum + b.segments.reduce((s2, seg) => s2 + (seg.end - seg.start), 0), 0)).toFixed(3),
      speakers: speakers.map((s) => ({ id: s.id, name: s.name, color: s.color })),
      blocks: blocks.map((b, i) => {
        let effId = b.speaker;
        if (!effId) { for (let j = i - 1; j >= 0; j--) { if (blocks[j].speaker) { effId = blocks[j].speaker; break; } } }
        return {
          order: i + 1,
          speaker: effId ? (speakerMap[effId] || null) : null,
          speakerId: b.speaker,
          explicitSpeaker: !!b.speaker,
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
  }, [blocks, speakers, audioName, projectTitle]);

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

  /* ── Export: Markdown ── */
  const generateMarkdown = useCallback(() => {
    let md = `# ${projectTitle}\n\n`;
    if (audioName) md += `*Source: ${audioName}*\n\n---\n\n`;
    let lastSpeakerName = null;
    blocks.forEach((block, idx) => {
      const spk = getEffectiveSpeaker(idx);
      const name = spk ? spk.name : "Unknown Speaker";
      const text = blockText(block);
      if (name !== lastSpeakerName) {
        md += `**${name}:** ${text}\n\n`;
      } else {
        md += `${text}\n\n`;
      }
      lastSpeakerName = name;
    });
    return md;
  }, [blocks, getEffectiveSpeaker, blockText, audioName, projectTitle]);

  const downloadMarkdown = useCallback(() => {
    const blob = new Blob([generateMarkdown()], { type: "text/markdown" });
    const url = URL.createObjectURL(blob);
    const fn = projectTitle.replace(/[^a-zA-Z0-9_ -]/g, "").replace(/\s+/g, "-").toLowerCase() || "oral-history";
    const a = document.createElement("a"); a.href = url; a.download = `${fn}.md`; a.click();
    URL.revokeObjectURL(url);
  }, [generateMarkdown, projectTitle]);

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
            <UploadBtn onChange={handleSubUpload} accept=".srt,.vtt,.txt" label={subName ? `◆ ${subName.slice(0, 18)}` : "↑ Subtitles"} />
            {audioUrl && <SmBtn onClick={transcribeAudio} accent disabled={transcribing}>
              {transcribing ? "…Transcribing" : "⟳ Transcribe"}
            </SmBtn>}
            <div style={{ width: 1, height: 20, background: C.border, margin: "0 2px" }} />
            <UploadBtn onChange={handleProjectUpload} accept=".json" label="↑ Load Project" />
            {hasContent && <SmBtn onClick={downloadProject} accent>↓ Save Project</SmBtn>}
            {hasContent && webhookBase && (
              <SmBtn onClick={cloudSave} accent>
                {syncStatus === "saving" ? "…" : "☁ Save"}
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
          <div style={{
            marginTop: 8, fontFamily: "'DM Mono', monospace", fontSize: 11, padding: "4px 10px",
            borderRadius: 4, display: "inline-block",
            background: syncStatus === "error" ? "rgba(179,64,64,0.08)" : C.accentDim,
            color: syncStatus === "error" ? C.danger : C.accent,
            ...(transcribing ? { animation: "pulse 2s ease-in-out infinite" } : {}),
          }}>
            {syncMsg}
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
            <label style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim }}>n8n Webhook Base URL</label>
            <input value={webhookBase} onChange={(e) => setWebhookBase(e.target.value)}
              placeholder="https://your-n8n.app/webhook/..."
              style={{ fontFamily: "'DM Mono', monospace", fontSize: 12, padding: "4px 8px", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, width: 320 }} />
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
            <label style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim }}>OpenAI API Key</label>
            <input value={openaiApiKey} onChange={(e) => setOpenaiApiKey(e.target.value)}
              type="password" placeholder="sk-..."
              style={{ fontFamily: "'DM Mono', monospace", fontSize: 12, padding: "4px 8px", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, width: 260 }} />
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
          <Sep />
          <TBtn onClick={undo} disabled={!undoStack.length} label="Undo" k="⌘Z" />
          <Sep />
          <TBtn onClick={playAll} disabled={!audioUrl} label="▶ Play" accent />
          <TBtn onClick={stopPlayback} label="■ Stop" />
          <Sep />
          <TBtn onClick={() => setShowSpeakers(!showSpeakers)} label={showSpeakers ? "◂ Speakers" : "▸ Speakers"} />
          <div style={{ flex: 1 }} />
          <span style={{ color: C.textDim, fontSize: 10 }}>
            {blocks.length} ¶ · {fmtTime(totalDuration)}
          </span>
          <Sep />
          <TBtn onClick={() => setShowExport(!showExport)} label="Export" accent />
        </div>
      )}

      <div style={{ flex: 1, display: "flex", overflow: "hidden" }}>

        {/* ── Speaker Panel ── */}
        {hasContent && showSpeakers && (
          <div style={{
            width: 210, borderRight: `1px solid ${C.border}`, background: C.surface,
            display: "flex", flexDirection: "column", flexShrink: 0,
          }}>
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
                    onClick={() => { if (selected.size > 0 && !isEd) assignAndMerge(spk.id); }}
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

        {/* ── Main Content: Paragraph View ── */}
        <div style={{ flex: 1, overflowY: "auto", padding: "32px 48px 100px", maxWidth: 800, margin: "0 auto", width: "100%" }}>
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

          {blocks.map((block, idx) => {
            const isSelected = selected.has(block.id);
            const isEditing = editingId === block.id;
            const isPlaying = playing === block.id;
            const isDragTarget = dragOver === block.id && dragging !== block.id;
            const spk = getEffectiveSpeaker(idx);
            const isInherited = !block.speaker && spk;
            const text = blockText(block);
            const duration = blockDuration(block);

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
                  marginBottom: 6,
                  padding: "14px 20px",
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
                  ) : (
                    <span style={{ fontFamily: "'DM Mono', monospace", fontSize: 11, color: C.border, fontStyle: "italic" }}>
                      unattributed
                    </span>
                  )}
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
                      autoFocus value={editText} onChange={(e) => setEditText(e.target.value)}
                      onKeyDown={(e) => { if (e.key === "Escape") cancelEdit(); }}
                      style={{
                        width: "100%", background: C.surface, border: `1px solid ${C.accent}`, borderRadius: 4,
                        color: C.text, fontFamily: "'Newsreader', Georgia, serif", fontSize: 17, lineHeight: 1.75,
                        padding: "12px 14px", resize: "vertical", minHeight: 80,
                      }}
                    />
                    <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                      <SmBtn onClick={commitEdit} accent>Save</SmBtn>
                      <SmBtn onClick={() => {
                        // Split at cursor position
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
                    {text}
                  </p>
                )}

                {/* Hover actions */}
                {!isEditing && isSelected && (
                  <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
                    {audioUrl && <SmBtn onClick={(e) => { e.stopPropagation(); playBlock(block); }}>▶ Play</SmBtn>}
                    <SmBtn onClick={(e) => { e.stopPropagation(); startEdit(block); }}>✎ Edit</SmBtn>
                    <SmBtn onClick={(e) => { e.stopPropagation(); pushUndo(); splitBlock(block.id, Math.floor(text.length / 2)); }}>Split ½</SmBtn>
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {/* ── Export Panel ── */}
        {showExport && hasContent && (
          <div style={{ width: 340, borderLeft: `1px solid ${C.border}`, background: C.surface, overflowY: "auto", display: "flex", flexDirection: "column", flexShrink: 0 }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}` }}>
              <h3 style={{ fontFamily: "'Instrument Serif', serif", fontSize: 18, fontWeight: 400 }}>Export</h3>
              <p style={{ fontFamily: "'DM Mono', monospace", fontSize: 10, color: C.textDim, marginTop: 4 }}>
                {blocks.length} paragraphs · {fmtTime(totalDuration)} · {speakers.length} speakers
              </p>
              <div style={{ display: "flex", gap: 8, marginTop: 12, flexWrap: "wrap" }}>
                <SmBtn onClick={downloadMarkdown} accent>↓ Markdown (.md)</SmBtn>
                <SmBtn onClick={downloadProject} accent>↓ Project (.json)</SmBtn>
              </div>
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
