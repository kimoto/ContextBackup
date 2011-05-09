#include <Windows.h>
#include <ShlObj.h>
#include <shlwapi.h>
#include <objbase.h>
#include <olectl.h>
#include <new>

#include "Util.h"

#pragma comment(lib, "shlwapi.lib")

#define INITGUID
#include <guiddef.h>
#include <propkeydef.h>

#include <string>
#include <list>
using namespace std;

// {EFB590EF-1ACE-4d83-88C8-F147D9B556AD}
DEFINE_GUID(CLSID_ContextMenuSample, 0xefb590ef, 0x1ace, 0x4d83, 0x88, 0xc8, 0xf1, 0x47, 0xd9, 0xb5, 0x56, 0xad);
const TCHAR g_szClsid[] = L"{EFB590EF-1ACE-4d83-88C8-F147D9B556AD}";
const TCHAR g_szProgid[] = L"ContextBackupExt";
//const TCHAR g_szExt[] = L".bhy";
const TCHAR g_szHandlerName[] = L"ContextBackupExtHandler";

LONG      g_lLocks = 0;			// ロック用変数？
HINSTANCE g_hinstDll = NULL;	// インスタンスハンドル

void doBackup();
void doRestore();
void LockModule(BOOL bLock);
BOOL CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue, LPTSTR lpszData);

BOOL g_isExistBackupFile = FALSE;
std::list<std::wstring> g_fileNameList;

// ロックの管理
// 参照数を増やしたり減らしたり
void LockModule(BOOL bLock)
{
	if (bLock)
		InterlockedIncrement(&g_lLocks);
	else
		InterlockedDecrement(&g_lLocks);
}

// 任意のレジストリキーを作成します
BOOL CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue, LPTSTR lpszData)
{
	HKEY  hKey;
	LONG  lResult;
	DWORD dwSize;

	lResult = RegCreateKeyEx(hKeyRoot, lpszKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
	if (lResult != ERROR_SUCCESS)
		return FALSE;

	if (lpszData != NULL)
		dwSize = (lstrlen(lpszData) + 1) * sizeof(TCHAR);
	else
		dwSize = 0;

	RegSetValueEx(hKey, lpszValue, 0, REG_SZ, (LPBYTE)lpszData, dwSize);
	RegCloseKey(hKey);
	
	return TRUE;
}

// ===================================
//	CContextMenu
// ===================================
// メニューの初期化実行、表示対象となるファイルを伝えるために使用
class CContextMenu : public IContextMenu, public IShellExtInit
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	STDMETHODIMP QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
	STDMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax);
	STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);

	STDMETHODIMP Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject *pdtobj, HKEY hkeyProgID);
	
	CContextMenu();
	~CContextMenu();

private:
	LONG m_cRef;
};

CContextMenu::CContextMenu()
{
	m_cRef = 1;

	LockModule(TRUE); // コンストラクタでロックする
}

CContextMenu::~CContextMenu()
{
	LockModule(FALSE); // デストラクタでロックの解放
}

STDMETHODIMP CContextMenu::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	// 渡されたriidがIUnknownかIContextMenuであれば
	// その両方を実装してるIContextMenuにキャストしますということか
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IContextMenu))
		*ppvObject = static_cast<IContextMenu *>(this);
	else if (IsEqualIID(riid, IID_IShellExtInit)) // ShellExtInit型であったときもその型に
		*ppvObject = static_cast<IShellExtInit *>(this);
	else
		return E_NOINTERFACE; // どちらもなかった場合はエラー値を返却

	// ここは成功した場合だけなので、参照カウンタを増やします
	AddRef();

	return S_OK;
}

// 参照カウンタを増やす
STDMETHODIMP_(ULONG) CContextMenu::AddRef()
{
	// インクリメントするだけの関数、他のスレッドによってぶっこわされないことが保証される
	return InterlockedIncrement(&m_cRef);
}

// 参照カウンタを減らす
STDMETHODIMP_(ULONG) CContextMenu::Release()
{
	// 減らした結果0になったら、もう不要ってことなので
	// 自分自身をdeleteしてデストラクタを呼びます
	if (InterlockedDecrement(&m_cRef) == 0) {	
		delete this;
		return 0;
	}

	return m_cRef;
}

// 右クリしたときに、メニュー表示する必要があったときに呼ばれる関数
STDMETHODIMP CContextMenu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
	MENUITEMINFO mii;
	
	if (uFlags & CMF_DEFAULTONLY) // デフォルトの項目だけであることが望まれるとき
		return MAKE_SCODE(SEVERITY_SUCCESS, FACILITY_NULL, 0);

	mii.cbSize     = sizeof(MENUITEMINFO);
	mii.fMask      = MIIM_ID | MIIM_TYPE;
	mii.fType      = MFT_STRING;
	mii.wID        = idCmdFirst;
	mii.dwTypeData = L"バックアップ";
	InsertMenuItem(hmenu, indexMenu, TRUE, &mii);
	
	mii.wID        = ++idCmdFirst;
	mii.dwTypeData = L"バックアップから復元";
	InsertMenuItem(hmenu, ++indexMenu, TRUE, &mii);
	
	return MAKE_SCODE(SEVERITY_SUCCESS, FACILITY_NULL, 2);
}

STDMETHODIMP CContextMenu::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
	UINT idCmd = LOWORD(pici->lpVerb);

	if (HIWORD(pici->lpVerb) != 0)
		return E_INVALIDARG;

	if (idCmd == 0)
		doBackup();
	else
		doRestore();

	return S_OK;
}

STDMETHODIMP CContextMenu::GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
	if (idCmd == 0) {
		if (uFlags == GCS_HELPTEXTA) // ASCIIで格納する
			lstrcpyA(pszName, "バックアップします");
		else if (uFlags == GCS_HELPTEXTW) // WIDEで格納する
			lstrcpyW((LPWSTR)pszName, L"バックアップします");
		else if (uFlags == GCS_VERBA) // 項目名ASCII
			lstrcpyA(pszName, "A");
		else if (uFlags == GCS_VERBW) // 項目名WIDE
			lstrcpyW((LPWSTR)pszName, L"A");
		else
			;
	}
	else if (idCmd == 1) {
		if (uFlags == GCS_HELPTEXTA)
			lstrcpyA(pszName, "バックアップから復元します");
		else if (uFlags == GCS_HELPTEXTW)
			lstrcpyW((LPWSTR)pszName, L"復元します");
		else if (uFlags == GCS_VERBA)
			lstrcpyA(pszName, "B");
		else if (uFlags == GCS_VERBW)
			lstrcpyW((LPWSTR)pszName, L"B");
		else
			;
	}
	else
		return E_FAIL;

	return S_OK;
}

STDMETHODIMP CContextMenu::Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject *pdtobj, HKEY hkeyProgID)
{
	g_fileNameList.clear();

	// pdtobjから右クリックしたファイル名が取得可能
	STGMEDIUM medium;
	FORMATETC fe = {CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	if( FAILED(pdtobj->GetData(&fe, &medium)) ){
		::ShowLastError();
		return S_OK;
	}
	
	// 選択されてるファイル数を取得
	UINT fileNum = ::DragQueryFile((HDROP)medium.hGlobal, (UINT)-1, NULL, 0);
	for(UINT i=0; i<fileNum; i++){
		TCHAR buf[MAX_PATH]=L"";
		if( FAILED(::DragQueryFile((HDROP)medium.hGlobal, i, buf, MAX_PATH)) ){
			::ShowLastError();
			return S_OK;
		}

		// 対象がファイルだったときのみ処理
		if( !::PathIsDirectory(buf) ){
			//::MessageBox(NULL, buf, L"Confirm", MB_OK);
			std::wstring wbuf = buf;
			g_fileNameList.push_back(buf);
		}
	}

	std::list<std::wstring>::iterator it = g_fileNameList.begin();
	while( it != g_fileNameList.end() ){
		//::MessageBox(NULL, it->c_str(), L"TEST", MB_OK);
		::OutputDebugString(it->c_str());
		it++;
	}

	::ReleaseStgMedium(&medium);
	return S_OK;
}

// ===================================
// COMクラスファクトリの宣言
// ===================================
class CClassFactory : public IClassFactory
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	
	STDMETHODIMP CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject);
	STDMETHODIMP LockServer(BOOL fLock);
};

STDMETHODIMP CClassFactory::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	// Factory or Unknownインターフェースならキャスト
	// それ以外はエラー
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IClassFactory))
		*ppvObject = static_cast<IClassFactory *>(this);
	else
		return E_NOINTERFACE;

	AddRef();
	
	return S_OK;
}

STDMETHODIMP_(ULONG) CClassFactory::AddRef()
{
	LockModule(TRUE); // ロックします
	return 2;
}

STDMETHODIMP_(ULONG) CClassFactory::Release()
{
	LockModule(FALSE); // アンロックします
	return 1;
}

// Factoryからインスタンスを作成します
STDMETHODIMP CClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
{
	CContextMenu *p;
	HRESULT      hr;
	
	*ppvObject = NULL;

	if (pUnkOuter != NULL)
		return CLASS_E_NOAGGREGATION;

	// コンテキストメニュークラスを生成します
	p = new CContextMenu();
	if (p == NULL)
		return E_OUTOFMEMORY; // 失敗したときはエラーを返す

	// コンテキストメニュークラスのQueryInterfaceを実行
	// riidがIUnknownかIContextMenuなら正常にインスタンス生成可能
	hr = p->QueryInterface(riid, ppvObject);
	p->Release();

	return hr; // HRESULT == S_OKなど
}

// サーバーをロックしたり解除したりするUtil関数？
STDMETHODIMP CClassFactory::LockServer(BOOL fLock)
{
	LockModule(fLock);
	return S_OK;
}

//================================
//	DLL Export
//================================
// g_lLocksの値が0なら、つまりどこからもロックされてなかったときは
// S_OKを返す、おそらくS_OKを返せばOSがアンロードしてくれるんだろう
STDAPI DllCanUnloadNow()
{
	return g_lLocks == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	// 静的変数にクラスファクトリを格納、以後も使用するってことか
	static CClassFactory serverFactory;
	HRESULT hr;

	*ppv  = NULL;
	
	// 自分の作るシェル拡張出会った場合のみ、ファクトリを生成して
	// そうでなかったらそんなのありませんって返事をしてる
	if (IsEqualCLSID(rclsid, CLSID_ContextMenuSample))
		hr = serverFactory.QueryInterface(riid, ppv);
	else
		hr = CLASS_E_CLASSNOTAVAILABLE;

	return hr;
}

// regsvr32.dllに呼ばれます
// 主にレジストリに登録する作業なはず
STDAPI DllRegisterServer(void)
{
	TCHAR szModulePath[MAX_PATH];
	TCHAR szKey[256];

	wsprintf(szKey, TEXT("CLSID\\%s"), g_szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, TEXT("ContextBackup ShellExtension")))
		return E_FAIL;

	GetModuleFileName(g_hinstDll, szModulePath, sizeof(szModulePath) / sizeof(TCHAR));
	wsprintf(szKey, TEXT("CLSID\\%s\\InprocServer32"), g_szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, szModulePath))
		return E_FAIL;
	
	wsprintf(szKey, TEXT("CLSID\\%s\\InprocServer32"), g_szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, TEXT("ThreadingModel"), TEXT("Apartment")))
		return E_FAIL;
	
	/*
	wsprintf(szKey, TEXT("%s"), g_szExt);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, (LPTSTR)g_szProgid))
		return E_FAIL;
		*/

	/*
	wsprintf(szKey, TEXT("%s\\shellex\\ContextMenuHandlers\\%s"), g_szProgid, g_szHandlerName);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, (LPTSTR)g_szClsid))
		return E_FAIL;
	*/

	// すべてのファイルを対象とする
	wsprintf(szKey, TEXT("%s\\shellex\\ContextMenuHandlers\\%s"), L"*", g_szHandlerName);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, (LPTSTR)g_szClsid))
		return E_FAIL;
	return S_OK;
}

// regsvr32.dll /u に呼ばれます
STDAPI DllUnregisterServer(void)
{
	TCHAR szKey[256];

	wsprintf(szKey, TEXT("CLSID\\%s"), g_szClsid);
	SHDeleteKey(HKEY_CLASSES_ROOT, szKey);
	/*
	SHDeleteKey(HKEY_CLASSES_ROOT, g_szExt);
	*/

	SHDeleteKey(HKEY_CLASSES_ROOT, g_szProgid);
	
	wsprintf(szKey, TEXT("%s\\shellex\\ContextMenuHandlers\\%s"), L"*", g_szHandlerName);
	SHDeleteKey(HKEY_CLASSES_ROOT, szKey);

	return S_OK;
}

BOOL WINAPI DllMain(HINSTANCE hinstDLL,DWORD fdwReason,LPVOID lpvReserved)
{
	if(fdwReason == DLL_PROCESS_ATTACH){
		g_hinstDll = hinstDLL;
		DisableThreadLibraryCalls(hinstDLL);
    }
    return  TRUE;
}

void doBackup()
{
	std::list<std::wstring>::iterator it = g_fileNameList.begin();
	while( it != g_fileNameList.end() ){
		if( BackupFile(it->c_str(), L".bak") ){
			::MessageBeep(MB_ICONASTERISK);
		}else{
			::MessageBeep(MB_ICONHAND);
			::ErrorMessageBox(L"ファイルのバックアップに失敗しました");
		}
		it++;
	}
}

void doRestore()
{
	std::list<std::wstring>::iterator it = g_fileNameList.begin();
	while( it != g_fileNameList.end() ){
		if( RestoreFile(it->c_str(), L".bak") ){
			::MessageBeep(MB_ICONASTERISK);
		}else{
			::MessageBeep(MB_ICONHAND);
			::ErrorMessageBox(L"ファイルの復元に失敗しました");
		}
		it++;
	}
}
