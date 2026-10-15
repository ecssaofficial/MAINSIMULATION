```typescript
import type { MissionSession } from './types';

/**
 * Stage 1 and Stage 2 penalties are counts of mistakes.
 * Stage 1: 3 seconds per penalty.
 * Stage 2: 15 seconds per penalty.
 * Stage 3: manually selected overtime penalty in seconds (0, 15, 30, or 45).
 */
export type StagePenalties = {
  stage1: number;
  stage2: number;
  stage3: number;
};

const penaltyStorageKey = (id: string) => `mosaic-penalties-${id}`;

function clampPenaltyCount(value: unknown, max: number): number {
  const count = Number(value ?? 0);
  return Number.isInteger(count) && count >= 0 && count <= max ? count : 0;
}

export function getPenalties(id: string): StagePenalties {
  try {
    const stored = JSON.parse(
      localStorage.getItem(penaltyStorageKey(id)) || '{}',
    );

    return {
      // Stage 1 supports 0–20 penalties.
      stage1: clampPenaltyCount(stored.stage1, 20),

      // Preserve the existing Stage 2 dropdown range of 0–4.
      stage2: clampPenaltyCount(stored.stage2, 4),

      // Stage 3 remains a selected penalty measured in seconds.
      stage3: [0, 15, 30, 45].includes(Number(stored.stage3))
        ? Number(stored.stage3)
        : 0,
    };
  } catch {
    return { stage1: 0, stage2: 0, stage3: 0 };
  }
}

export function savePenalties(
  id: string,
  penalties: StagePenalties,
): void {
  localStorage.setItem(
    penaltyStorageKey(id),
    JSON.stringify(penalties),
  );
}

export function clearPenalties(id: string): void {
  localStorage.removeItem(penaltyStorageKey(id));
}

export function secondsToClock(seconds: number): string {
  const safeSeconds = Math.max(0, Math.floor(seconds));
  const hours = Math.floor(safeSeconds / 3600);
  const minutes = Math.floor((safeSeconds % 3600) / 60);
  const remainingSeconds = safeSeconds % 60;

  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(remainingSeconds).padStart(2, '0')}`;
}

function durationBetween(
  start: string | null,
  end: string | null,
): number {
  if (!start || !end) return 0;

  const startMs = new Date(start).getTime();
  const endMs = new Date(end).getTime();

  if (!Number.isFinite(startMs) || !Number.isFinite(endMs)) {
    return 0;
  }

  return Math.max(0, Math.floor((endMs - startMs) / 1000));
}

/**
 * Overall raw elapsed time from mission start until completion (or now).
 */
export function elapsedSeconds(session: MissionSession): number {
  if (!session.started_at) return 0;

  const endTime = session.completed_at
    ? new Date(session.completed_at).getTime()
    : Date.now();

  const startTime = new Date(session.started_at).getTime();

  if (!Number.isFinite(startTime) || !Number.isFinite(endTime)) {
    return 0;
  }

  return Math.max(0, Math.floor((endTime - startTime) / 1000));
}

/**
 * Build the complete result payload for Google Apps Script.
 *
 * Stage 1 raw duration:
 * Mission start -> Stage 1 completion.
 *
 * Stage 2 raw duration:
 * Stage 1 completion -> Stage 2 completion.
 *
 * Stage 3 raw duration:
 * Stage 2 completion -> Mission completion.
 *
 * Each stage's adjusted time includes its own penalty time.
 * Overall adjusted time is overall raw time plus all penalty seconds.
 * The overall clock itself is never reset between stages.
 */
export function buildResult(
  session: MissionSession,
  penalties: StagePenalties,
) {
  const rawSeconds = elapsedSeconds(session);

  // Stage 1: mission start -> Stage 1 completion.
  const stage1DurationSeconds = durationBetween(
    session.started_at,
    session.stage_1_completed_at,
  );

  // Stage 2: Stage 1 completion -> Stage 2 completion.
  const stage2DurationSeconds = durationBetween(
    session.stage_1_completed_at,
    session.stage_2_completed_at,
  );

  // Stage 3: Stage 2 completion -> mission completion.
  const stage3DurationSeconds = durationBetween(
    session.stage_2_completed_at,
    session.completed_at,
  );

  const stage1PenaltySeconds = penalties.stage1 * 3;
  const stage2PenaltySeconds = penalties.stage2 * 15;
  const stage3PenaltySeconds = penalties.stage3;

  const totalPenaltySeconds =
    stage1PenaltySeconds +
    stage2PenaltySeconds +
    stage3PenaltySeconds;

  const stage3OvertimeSeconds = Math.max(
    0,
    stage3DurationSeconds - 180,
  );

  const adjustedSeconds = rawSeconds + totalPenaltySeconds;

  return {
    sessionId: session.id,
    teamName: session.team_name,
    teamCode: session.team_code,

    // ISO timestamps are sent to Apps Script.
    // Code.gs must format start/end timestamps as 12-hour time with AM/PM.
    startedAt: session.started_at,
    stage1CompletedAt: session.stage_1_completed_at,
    stage2CompletedAt: session.stage_2_completed_at,
    completedAt: session.completed_at,

    // Overall timing.
    rawSeconds,
    rawTime: secondsToClock(rawSeconds),

    // Stage 1 timing and penalties.
    stage1DurationSeconds,
    stage1Duration: secondsToClock(stage1DurationSeconds),
    stage1Penalties: penalties.stage1,
    stage1PenaltySeconds,
    stage1AdjustedSeconds:
      stage1DurationSeconds + stage1PenaltySeconds,
    stage1AdjustedTime: secondsToClock(
      stage1DurationSeconds + stage1PenaltySeconds,
    ),

    // Stage 2 timing and penalties.
    stage2DurationSeconds,
    stage2Duration: secondsToClock(stage2DurationSeconds),
    stage2Penalties: penalties.stage2,
    stage2PenaltySeconds,
    stage2AdjustedSeconds:
      stage2DurationSeconds + stage2PenaltySeconds,
    stage2AdjustedTime: secondsToClock(
      stage2DurationSeconds + stage2PenaltySeconds,
    ),

    // Stage 3 timing and penalties.
    stage3DurationSeconds,
    stage3Duration: secondsToClock(stage3DurationSeconds),
    stage3OvertimeSeconds,
    stage3Overtime: secondsToClock(stage3OvertimeSeconds),
    stage3Penalties: penalties.stage3,
    stage3PenaltySeconds,

    // Overall penalty and adjusted time.
    totalPenaltyCount: penalties.stage1 + penalties.stage2,
    totalPenaltySeconds,
    penaltyTime: secondsToClock(totalPenaltySeconds),
    adjustedSeconds,
    adjustedTime: secondsToClock(adjustedSeconds),

    displayedResult:
      `${secondsToClock(rawSeconds)} + ` +
      `(${secondsToClock(totalPenaltySeconds)}) = ` +
      `${secondsToClock(adjustedSeconds)}`,
  };
}

export async function sendResultToSheet(
  payload: ReturnType<typeof buildResult>,
) {
  const endpoint = (
    import.meta.env.VITE_GOOGLE_APPS_SCRIPT_URL as
      | string
      | undefined
  )?.trim();

  if (!endpoint) {
    return { configured: false as const };
  }

  await fetch(endpoint, {
    method: 'POST',
    mode: 'no-cors',
    headers: {
      'Content-Type': 'text/plain;charset=utf-8',
    },
    body: JSON.stringify(payload),
  });

  // With no-cors, the browser cannot read the Apps Script response.
  // This confirms only that the request was attempted, not that a row was saved.
  return { configured: true as const };
}
```
