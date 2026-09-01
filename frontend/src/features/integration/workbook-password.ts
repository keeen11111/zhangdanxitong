import { UploadFailure } from "@/lib/api";

export function isEncryptedWorkbookFailure(failure: UploadFailure) {
  return /加密|打开密码/.test(failure.message);
}

export function promptForWorkbookPassword() {
  const password = window.prompt("检测到加密的 Excel 文件，请输入打开密码：");
  return password === null ? undefined : password;
}
