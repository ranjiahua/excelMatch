import type { WorkbookData } from "../types";
import { parseExcelFile } from "../excel";

interface FileUploaderProps {
  label: string;
  onParsed: (data: WorkbookData, file: File) => void;
  onError: (msg: string) => void;
  onLoading: (loading: boolean) => void;
}

export default function FileUploader({
  label,
  onParsed,
  onError,
  onLoading,
}: FileUploaderProps) {
  async function handleFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    onError("");
    onLoading(true);
    try {
      const data = await parseExcelFile(file);
      onParsed(data, file);
    } catch (err) {
      onError(
        err instanceof Error ? err.message : "解析 Excel 文件失败"
      );
    } finally {
      onLoading(false);
      e.target.value = "";
    }
  }

  return (
    <div className="file-uploader">
      <label className="file-label">{label}</label>
      <input type="file" accept=".xlsx,.xls" onChange={handleFile} />
    </div>
  );
}
