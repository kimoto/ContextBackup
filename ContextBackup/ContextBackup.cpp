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

LONG      g_lLocks = 0;			// ���b�N�p�ϐ��H
HINSTANCE g_hinstDll = NULL;	// �C���X�^���X�n���h��

void doBackup();
void doRestore();
void LockModule(BOOL bLock);
BOOL CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue, LPTSTR lpszData);

BOOL g_isExistBackupFile = FALSE;
std::list<std::wstring> g_fileNameList;

// ���b�N�̊Ǘ�
// �Q�Ɛ��𑝂₵���茸�炵����
void LockModule(BOOL bLock)
{
	if (bLock)
		InterlockedIncrement(&g_lLocks);
	else
		InterlockedDecrement(&g_lLocks);
}

// �C�ӂ̃��W�X�g���L�[���쐬���܂�
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
// ���j���[�̏��������s�A�\���ΏۂƂȂ�t�@�C����`���邽�߂Ɏg�p
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

	LockModule(TRUE); // �R���X�g���N�^�Ń��b�N����
}

CContextMenu::~CContextMenu()
{
	LockModule(FALSE); // �f�X�g���N�^�Ń��b�N�̉��
}

STDMETHODIMP CContextMenu::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	// �n���ꂽriid��IUnknown��IContextMenu�ł����
	// ���̗������������Ă�IContextMenu�ɃL���X�g���܂��Ƃ������Ƃ�
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IContextMenu))
		*ppvObject = static_cast<IContextMenu *>(this);
	else if (IsEqualIID(riid, IID_IShellExtInit)) // ShellExtInit�^�ł������Ƃ������̌^��
		*ppvObject = static_cast<IShellExtInit *>(this);
	else
		return E_NOINTERFACE; // �ǂ�����Ȃ������ꍇ�̓G���[�l��ԋp

	// �����͐��������ꍇ�����Ȃ̂ŁA�Q�ƃJ�E���^�𑝂₵�܂�
	AddRef();

	return S_OK;
}

// �Q�ƃJ�E���^�𑝂₷
STDMETHODIMP_(ULONG) CContextMenu::AddRef()
{
	// �C���N�������g���邾���̊֐��A���̃X���b�h�ɂ���ĂԂ����킳��Ȃ����Ƃ��ۏ؂����
	return InterlockedIncrement(&m_cRef);
}

// �Q�ƃJ�E���^�����炷
STDMETHODIMP_(ULONG) CContextMenu::Release()
{
	// ���炵������0�ɂȂ�����A�����s�v���Ă��ƂȂ̂�
	// �������g��delete���ăf�X�g���N�^���Ăт܂�
	if (InterlockedDecrement(&m_cRef) == 0) {	
		delete this;
		return 0;
	}

	return m_cRef;
}

// �E�N�������Ƃ��ɁA���j���[�\������K�v���������Ƃ��ɌĂ΂��֐�
STDMETHODIMP CContextMenu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
	MENUITEMINFO mii;
	
	if (uFlags & CMF_DEFAULTONLY) // �f�t�H���g�̍��ڂ����ł��邱�Ƃ��]�܂��Ƃ�
		return MAKE_SCODE(SEVERITY_SUCCESS, FACILITY_NULL, 0);

	mii.cbSize     = sizeof(MENUITEMINFO);
	mii.fMask      = MIIM_ID | MIIM_TYPE;
	mii.fType      = MFT_STRING;
	mii.wID        = idCmdFirst;
	mii.dwTypeData = L"�o�b�N�A�b�v";
	InsertMenuItem(hmenu, indexMenu, TRUE, &mii);
	
	mii.wID        = ++idCmdFirst;
	mii.dwTypeData = L"�o�b�N�A�b�v���畜��";
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
		if (uFlags == GCS_HELPTEXTA) // ASCII�Ŋi�[����
			lstrcpyA(pszName, "�o�b�N�A�b�v���܂�");
		else if (uFlags == GCS_HELPTEXTW) // WIDE�Ŋi�[����
			lstrcpyW((LPWSTR)pszName, L"�o�b�N�A�b�v���܂�");
		else if (uFlags == GCS_VERBA) // ���ږ�ASCII
			lstrcpyA(pszName, "A");
		else if (uFlags == GCS_VERBW) // ���ږ�WIDE
			lstrcpyW((LPWSTR)pszName, L"A");
		else
			;
	}
	else if (idCmd == 1) {
		if (uFlags == GCS_HELPTEXTA)
			lstrcpyA(pszName, "�o�b�N�A�b�v���畜�����܂�");
		else if (uFlags == GCS_HELPTEXTW)
			lstrcpyW((LPWSTR)pszName, L"�������܂�");
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

	// pdtobj����E�N���b�N�����t�@�C�������擾�\
	STGMEDIUM medium;
	FORMATETC fe = {CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	if( FAILED(pdtobj->GetData(&fe, &medium)) ){
		::ShowLastError();
		return S_OK;
	}
	
	// �I������Ă�t�@�C�������擾
	UINT fileNum = ::DragQueryFile((HDROP)medium.hGlobal, (UINT)-1, NULL, 0);
	for(UINT i=0; i<fileNum; i++){
		TCHAR buf[MAX_PATH]=L"";
		if( FAILED(::DragQueryFile((HDROP)medium.hGlobal, i, buf, MAX_PATH)) ){
			::ShowLastError();
			return S_OK;
		}

		// �Ώۂ��t�@�C���������Ƃ��̂ݏ���
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
// COM�N���X�t�@�N�g���̐錾
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

	// Factory or Unknown�C���^�[�t�F�[�X�Ȃ�L���X�g
	// ����ȊO�̓G���[
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IClassFactory))
		*ppvObject = static_cast<IClassFactory *>(this);
	else
		return E_NOINTERFACE;

	AddRef();
	
	return S_OK;
}

STDMETHODIMP_(ULONG) CClassFactory::AddRef()
{
	LockModule(TRUE); // ���b�N���܂�
	return 2;
}

STDMETHODIMP_(ULONG) CClassFactory::Release()
{
	LockModule(FALSE); // �A�����b�N���܂�
	return 1;
}

// Factory����C���X�^���X���쐬���܂�
STDMETHODIMP CClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
{
	CContextMenu *p;
	HRESULT      hr;
	
	*ppvObject = NULL;

	if (pUnkOuter != NULL)
		return CLASS_E_NOAGGREGATION;

	// �R���e�L�X�g���j���[�N���X�𐶐����܂�
	p = new CContextMenu();
	if (p == NULL)
		return E_OUTOFMEMORY; // ���s�����Ƃ��̓G���[��Ԃ�

	// �R���e�L�X�g���j���[�N���X��QueryInterface�����s
	// riid��IUnknown��IContextMenu�Ȃ琳��ɃC���X�^���X�����\
	hr = p->QueryInterface(riid, ppvObject);
	p->Release();

	return hr; // HRESULT == S_OK�Ȃ�
}

// �T�[�o�[�����b�N��������������肷��Util�֐��H
STDMETHODIMP CClassFactory::LockServer(BOOL fLock)
{
	LockModule(fLock);
	return S_OK;
}

//================================
//	DLL Export
//================================
// g_lLocks�̒l��0�Ȃ�A�܂�ǂ���������b�N����ĂȂ������Ƃ���
// S_OK��Ԃ��A�����炭S_OK��Ԃ���OS���A�����[�h���Ă����񂾂낤
STDAPI DllCanUnloadNow()
{
	return g_lLocks == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	// �ÓI�ϐ��ɃN���X�t�@�N�g�����i�[�A�Ȍ���g�p������Ă��Ƃ�
	static CClassFactory serverFactory;
	HRESULT hr;

	*ppv  = NULL;
	
	// �����̍��V�F���g���o������ꍇ�̂݁A�t�@�N�g���𐶐�����
	// �����łȂ������炻��Ȃ̂���܂�����ĕԎ������Ă�
	if (IsEqualCLSID(rclsid, CLSID_ContextMenuSample))
		hr = serverFactory.QueryInterface(riid, ppv);
	else
		hr = CLASS_E_CLASSNOTAVAILABLE;

	return hr;
}

// regsvr32.dll�ɌĂ΂�܂�
// ��Ƀ��W�X�g���ɓo�^�����ƂȂ͂�
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

	// ���ׂẴt�@�C����ΏۂƂ���
	wsprintf(szKey, TEXT("%s\\shellex\\ContextMenuHandlers\\%s"), L"*", g_szHandlerName);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, (LPTSTR)g_szClsid))
		return E_FAIL;
	return S_OK;
}

// regsvr32.dll /u �ɌĂ΂�܂�
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
			::ErrorMessageBox(L"�t�@�C���̃o�b�N�A�b�v�Ɏ��s���܂���");
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
			::ErrorMessageBox(L"�t�@�C���̕����Ɏ��s���܂���");
		}
		it++;
	}
}
