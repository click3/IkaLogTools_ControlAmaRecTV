
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


// �����I����8�l�̃X�R�A���\�����ꂽ�u�Ԃ���^���~����܂ł̑҂�����(�P�ʃ~���b)
const unsigned int WAIT_STOP_MILLIS = 20 * 1000;

// �^���~����^��t�@�C�����l�[�������݂�܂ł̑҂�����(�P�ʃ~���b)
const unsigned int WAIT_RENAME_MILLIS = 10 * 1000;

// �^��J�n/��~�̃z�b�g�L�[���������Ă���b���܂ł̑҂�����(�P�ʃ~���b)
const unsigned int WAIT_KEY_UP_MILLIS = 16;

// �t�@�C�����̏o�̓t�H�[�}�b�g
// %year% ����̎l���\�L�A2015�N��2015
// %month% ����2���\�L�A1����01
// %date% ����2���\�L�A1����01
// %hour% ����2���\�L�A1����01�A23����23
// %minutes% ����2���\�L�A1����01
// %second% �b��2���\�L�A1�b��01
// %stage% �X�e�[�W���̓��{��\�L
// %rule% ���[���̓��{��\�L
// %won% ���s�Awin/lose/unknown��3��̂����ꂩ
const std::wstring FILENAME_FORMAT = L"%year%%month%%date%_%hour%%minutes%_%stage%_%rule%_%won%.avi";

// �^��J�n�I���Ȃǂ��w������̂Ɏg���z�b�g�L�[
// �A���t�@�x�b�g�͑啶���ňꕶ���A����L�[��WinUser.h��VK_����n�܂�萔���g�p�ł���
// �����w��œ�������
// ex. �uCtrl+z�v�́uVK_CONTROL, 'Z'�v�A�uF7�v�́uVK_F7�v
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
