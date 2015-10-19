
#include <string>
#include <vector>
#include <algorithm>
#include <iostream>

#include <Windows.h>
#include <tlhelp32.h>

#include <boost/filesystem.hpp>
#include <boost/filesystem/fstream.hpp>
#include <boost/format.hpp>
#include <boost/algorithm/string.hpp>
#include <boost/date_time.hpp>


// 試合終了後8人のスコアが表示された瞬間から録画停止するまでの待ち時間(単位ミリ秒)
const unsigned int WAIT_STOP_MILLIS = 20 * 1000;

// 録画停止から録画ファイルリネームを試みるまでの待ち時間(単位ミリ秒)
const unsigned int WAIT_RENAME_MILLIS = 10 * 1000;

// 録画開始/停止のホットキーを押下してから話すまでの待ち時間(単位ミリ秒)
const unsigned int WAIT_KEY_UP_MILLIS = 16;

// ファイル名の出力フォーマット
// %year% 西暦の四桁表記、2015年は2015
// %month% 月の2桁表記、1月は01
// %date% 日の2桁表記、1日は01
// %hour% 時の2桁表記、1時は01、23時は23
// %minutes% 分の2桁表記、1分は01
// %second% 秒の2桁表記、1秒は01
// %stage% ステージ名の日本語表記
// %rule% ルールの日本語表記
// %won% 勝敗、win/lose/unknownの3種のいずれか
const std::wstring FILENAME_FORMAT = L"%year%%month%%date%_%hour%%minutes%_%stage%_%rule%_%won%.avi";

// 録画開始終了などを指示するのに使うホットキー
// アルファベットは大文字で一文字、特殊キーはWinUser.hのVK_から始まる定数を使用できる
// 複数指定で同時押し
// ex. 「Ctrl+z」は「VK_CONTROL, 'Z'」、「F7」は「VK_F7」
const unsigned int HOT_KEYS[] = {
  VK_CONTROL, 'Z',
};


unsigned int GetScanCode(const unsigned int virtualKeyCode) {
  return ::MapVirtualKeyExW(virtualKeyCode, MAPVK_VK_TO_VSC, NULL);
}
void PressKey(const unsigned int virtualKeyCode) {
  ::keybd_event(virtualKeyCode, GetScanCode(virtualKeyCode), 0, NULL);
}
void UpKey(const unsigned int virtualKeyCode) {
  ::keybd_event(virtualKeyCode, GetScanCode(virtualKeyCode), KEYEVENTF_KEYUP, NULL);
}
void SendHotKey() {
  for (const unsigned int key : HOT_KEYS) {
    PressKey(key);
  }
  Sleep(WAIT_KEY_UP_MILLIS);
  for (const unsigned int key : HOT_KEYS) {
    UpKey(key);
  }
}

std::wstring GetEnv(const wchar_t * const name) {
  wchar_t buf[32768];
  const unsigned int length = ::GetEnvironmentVariableW(name, buf, _countof(buf));
  if (length == 0) {
    return{};
  }
  return{ &buf[0], &buf[length] };
}
std::wstring CreateFilename() {
  std::wstring result = FILENAME_FORMAT;
  const auto src = boost::posix_time::second_clock::local_time();
  const tm now = boost::posix_time::to_tm(src);
  boost::algorithm::replace_all(result, L"%year%", (boost::wformat(L"%04d") % (now.tm_year + 1900)).str());
  boost::algorithm::replace_all(result, L"%month%", (boost::wformat(L"%02d") % (now.tm_mon + 1)).str());
  boost::algorithm::replace_all(result, L"%date%", (boost::wformat(L"%02d") % now.tm_mday).str());
  boost::algorithm::replace_all(result, L"%hour%", (boost::wformat(L"%02d") % now.tm_hour).str());
  boost::algorithm::replace_all(result, L"%minutes%", (boost::wformat(L"%02d") % now.tm_min).str());
  boost::algorithm::replace_all(result, L"%second%", (boost::wformat(L"%02d") % now.tm_sec).str());
  boost::algorithm::replace_all(result, L"%stage%", GetEnv(L"IKALOG_STAGE"));
  boost::algorithm::replace_all(result, L"%rule%", GetEnv(L"IKALOG_RULE"));
  boost::algorithm::replace_all(result, L"%won%", GetEnv(L"IKALOG_WON"));
  return result;
}
boost::filesystem::path GetDestDir() {
  return GetEnv(L"IKALOG_MP4_DESTDIR");
}
boost::filesystem::path GetDestFilename() {
  return CreateFilename();
}
boost::filesystem::path GetDestPath() {
  boost::filesystem::path path = GetDestDir() / GetDestFilename();
  path.replace_extension(".avi");
  return path;
}

std::vector<std::wstring> GetVisibleWindowTitleList() {
  std::vector<std::wstring> result;
  ::EnumWindows([](const HWND handle, const LPARAM resultPtr) -> BOOL {
    if (!::IsWindowVisible(handle)) {
      return TRUE;
    }
    wchar_t buf[4096];
    const unsigned int length = ::GetWindowTextW(handle, buf, _countof(buf));
    reinterpret_cast<std::vector<std::wstring>*>(resultPtr)->push_back({ &buf[0], &buf[length] });
    return TRUE;
  }, reinterpret_cast<LPARAM>(&result));
  return result;
}
std::wstring GetAmaRecTvTitle() {
  const auto list = GetVisibleWindowTitleList();
  const auto it = std::find_if(list.begin(), list.end(), [](const std::wstring &title) {
    if (title.length() < 8) {
      return false;
    }
    return title.substr(0, 8) == L"AmaRecTV";
  });
  if (it == list.end()) {
    return{};
  }
  return *it;
}
boost::filesystem::path GetSrcDir() {
  return GetDestDir();
}
boost::filesystem::path GetSrcPath() {
  const std::wstring title = GetAmaRecTvTitle();
  if (title.empty() || title.length() < 15 || title.substr(title.length() - 4, 4) != L".avi") {
    return{};
  }
  return GetSrcDir() / title.substr(10);
}

bool IsFileExclusiveLock(const boost::filesystem::path &path) {
  if (path.empty() || !boost::filesystem::is_regular_file(path)) {
    return false;
  }
  const HANDLE handle = ::CreateFileW(path.wstring().c_str(), GENERIC_WRITE, FILE_SHARE_WRITE | FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
  if (handle == INVALID_HANDLE_VALUE) {
    return true;
  }
  ::CloseHandle(handle);
  return false;
}
bool IsRecording() {
  return IsFileExclusiveLock(GetSrcPath());
}

void WriteDebugLog() {
  boost::filesystem::wofstream ofs(L"debug.log", std::ios::app);
  ofs << boost::wformat(L"----\n");
  ofs << boost::wformat(L"commandline: %s\n") % ::GetCommandLineW();
  const wchar_t * const envNameList[] = {
    L"IKALOG_MP4_DESTDIR", L"IKALOG_MP4_DESTNAME", L"IKALOG_STAGE", L"IKALOG_RULE", L"IKALOG_WON"
  };
  for (const wchar_t * const envName : envNameList) {
    ofs << boost::wformat(L"%s: %s\n") % envName % GetEnv(envName);
  }
  ofs << boost::wformat(L"window title list:\n");
  for (const std::wstring &title : GetVisibleWindowTitleList()) {
    ofs << boost::wformat(L"\t%s\n") % title;
  }
  ofs << boost::wformat(L"AmaRecTV title: %s\n") % GetAmaRecTvTitle();
  ofs << boost::wformat(L"src path: %s\n") % GetSrcPath();
  ofs << boost::wformat(L"dest path: %s\n") % GetDestPath();
  ofs << boost::wformat(L"is recording?: %s\n") % (IsRecording() ? L"true" : L"false");
  ofs << boost::wformat(L"----\n");
  ofs.close();
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  std::locale loc = std::locale("japanese").combine<std::numpunct<char> >(std::locale::classic()).combine<std::numpunct<wchar_t> >(std::locale::classic());
  std::locale::global(loc);
  std::wcout.imbue(loc);
  std::cout.imbue(loc);

  //WriteDebugLog();

  int argc = 0;
  const wchar_t * const * const argv = ::CommandLineToArgvW(::GetCommandLineW(), &argc);
  if (argv == nullptr || argc != 2) {
    return 1;
  }
  const std::wstring op = argv[1];

  bool start;
  if (op == L"start") {
    start = true;
  } else if (op == L"stop") {
    start = false;
  } else {
    return 1;
  }

  if (start) {
    if (IsRecording()) {
      return 1;
    }
    SendHotKey();
    return 0;
  }
  if (!IsRecording()) {
    return 1;
  }
  ::Sleep(WAIT_STOP_MILLIS);
  SendHotKey();
  ::Sleep(WAIT_RENAME_MILLIS);
  boost::filesystem::rename(GetSrcPath(), GetDestPath());
  return 0;
}
