/**
 * Forms が出力する日時文字列の解析と、アプリで使う日付の整形。
 * moment の役割はこの3関数だけなので、依存を持たない最小実装で置き換える。
 */

// 「M/D/Y H:mm[:ss]」（年は2桁か4桁）と「YYYY/M/D H:mm[:ss]」の2系統を受け付ける。
// 区切りは / と - のどちらでもよいが、後方参照で日付内の混在は許さない。
const FORMS_DATE_TIME =
  /^(?:(\d{4})([\/-])(\d{1,2})\2(\d{1,2})|(\d{1,2})([\/-])(\d{1,2})\6(\d{4}|\d{2})) (\d{1,2}):(\d{2})(?::(\d{2}))?$/;

/** moment と同じ規則: 68 以下は 2000 年代、69 以上は 1900 年代とみなす。 */
function expandYear(raw: string): number {
  const value = Number(raw);
  if (raw.length === 4) {
    return value;
  }
  return value < 69 ? 2000 + value : 1900 + value;
}

const pad = (value: number) => String(value).padStart(2, '0');

/** 解析できない場合は null を返す（moment の strict 解析 + isValid() 相当）。 */
export function parseFormsDateTime(value: string): Date | null {
  const matched = FORMS_DATE_TIME.exec(String(value ?? ''));
  if (!matched) {
    return null;
  }

  const [year, month, day] = matched[1] !== undefined
    ? [Number(matched[1]), Number(matched[3]), Number(matched[4])]
    : [expandYear(matched[8]), Number(matched[5]), Number(matched[7])];
  const hour = Number(matched[9]);
  const minute = Number(matched[10]);
  const second = matched[11] === undefined ? 0 : Number(matched[11]);

  const parsed = new Date(year, month - 1, day, hour, minute, second, 0);
  // 02/30 や 24:00 のように繰り上がる値は、strict 解析と同様に無効として扱う。
  const exact = parsed.getFullYear() === year
    && parsed.getMonth() === month - 1
    && parsed.getDate() === day
    && parsed.getHours() === hour
    && parsed.getMinutes() === minute
    && parsed.getSeconds() === second;

  return exact ? parsed : null;
}

/** 出席簿の列見出しに使う YY-MM-DD 形式。 */
export function formatShortDate(date: Date): string {
  return `${pad(date.getFullYear() % 100)}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

/** 出力ファイル名に使う YYYYMMDD-HHmmss 形式。 */
export function formatTimestamp(date: Date): string {
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}`
    + `-${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
}
