#include <windows.h>
#include <ocidl.h>
#include <olectl.h>

/* //to usage
	RECT rc;
	long hmWidth=0;
	long hmHeight=0;
	lpPic->get_Height(&hmHeight);
	lpPic->get_Width(&hmWidth);
	GetClientRect(hWnd, &rc);
	int nWidth,nHeight;
	nWidth=rc.right - rc.left;
	nHeight=rc.bottom - rc.top;
	
	long h, w;
	h = hmHeight;
	w = hmWidth;
	h = MulDiv(h, GetDeviceCaps(hdc, LOGPIXELSX), 2540);
	w = MulDiv(w, GetDeviceCaps(hdc, LOGPIXELSY), 2540);
	
	//zero point in bottom right corner
	HRESULT hr = lpPic->Render(hdc,nWidth,0, -w, h,hmWidth,hmHeight,-hmWidth,-hmHeight,NULL);
*/

bool LoadJpg(LPCTSTR name, LPPICTURE* lppi)
{
	HANDLE hFile = CreateFile(name, GENERIC_READ, 0, NULL, OPEN_EXISTING, 0, NULL);
	
	if (hFile == INVALID_HANDLE_VALUE)
	{
		return false;
	}

	DWORD dwFileSize = GetFileSize(hFile, NULL);

	if (-1 == dwFileSize)
	{
		CloseHandle(hFile);
		return false;
	}

	LPVOID pvData;
	HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, dwFileSize);

	if (NULL == hGlobal)
	{
		GlobalUnlock(hGlobal);
		CloseHandle(hFile);
		return false;
	}

	pvData = GlobalLock(hGlobal);

	if (NULL == pvData)
	{
		GlobalUnlock(hGlobal);
		CloseHandle(hFile);
		return false;
	}

	DWORD dwFileRead = 0;
	BOOL bRead = ReadFile(hFile, pvData, dwFileSize, &dwFileSize, NULL);

	GlobalUnlock(hGlobal);
	CloseHandle(hFile);

	if (FALSE == bRead)
	{
		return false;
	}

	LPSTREAM pstm = NULL;
	HRESULT hr = CreateStreamOnHGlobal(hGlobal, TRUE, &pstm);
	
	if (FAILED(hr))
	{
		return false;
	}

	if (*lppi)
	{
		(*lppi)->Release();
	}

	hr = OleLoadPicture(pstm, dwFileSize, FALSE, IID_IPicture, (LPVOID*)lppi);
	pstm->Release();

	if (FAILED(hr) || *lppi == NULL)
	{
		return false;
	}

	return true;
}